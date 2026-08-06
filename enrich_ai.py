#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
라쿠텐 수집 결과 -> AI 번역 + 스마트스토어 카테고리 매핑 + 태그 생성
--------------------------------------------------------------------
rakuten_collector.py 가 만든 엑셀을 입력으로 받아,
판매용 상품명 / 한국어 상세설명 / 카테고리ID / 태그를 붙여서 다시 저장합니다.

구조
  1단계) 상품에 등장한 genreId만 라쿠텐 Genre API로 조회 -> 전체 경로 복원
         -> 경로 번역 -> 스마트스토어 카테고리 후보 추림 -> LLM이 최종 선택
         -> genre_cache.json 에 저장 (다음 실행부터 무료)
  2단계) 상품별 상품명/상세설명 번역 + 판매용 상품명 + 태그 생성 (병렬)

준비
  pip install requests pandas openpyxl xlrd openai
  export RAKUTEN_APP_ID=... RAKUTEN_ACCESS_KEY=... OPENAI_API_KEY=...

  스마트스토어 카테고리 CSV 필요 (네이버 셀러 도움말에서 다운로드).
  컬럼명이 달라도 'ID처럼 보이는 열' + '나머지 텍스트 열'을 자동으로 잡습니다.

실행
  python enrich_ai.py rakuten_items.xlsx --cat naver_categories.csv --out final.xlsx
  python enrich_ai.py rakuten_items.xlsx --no-category      # 번역만
"""

import argparse
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests

GENRE_API = "https://openapi.rakuten.co.jp/ichibagt/api/IchibaGenre/Search/20260401"
APP_ID = os.environ.get("RAKUTEN_APP_ID", "")
ACCESS_KEY = os.environ.get("RAKUTEN_ACCESS_KEY", "")

MODEL = os.environ.get("ENRICH_MODEL", "gpt-4.1-mini")
GENRE_CACHE = "genre_cache.json"
WORKERS = 4

# 표시광고법상 위험한 표현. 상품명에서 걸러냅니다.
BANNED = ["최고", "최상", "1위", "일등", "정품보장", "무조건", "100%",
          "완벽", "최저가", "국내최초", "세계최초", "특허", "공식"]


# ----------------------------------------------------------------------
# LLM
# ----------------------------------------------------------------------
def get_client():
    from openai import OpenAI
    key = os.environ.get("OPENAI_API_KEY")
    if not key:
        sys.exit("OPENAI_API_KEY 환경변수를 설정하세요.")
    return OpenAI(api_key=key)


def ask_json(client, system: str, user: str, retries: int = 3):
    """JSON만 뱉도록 강제하고 파싱까지."""
    for i in range(retries):
        try:
            r = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "system", "content": system},
                          {"role": "user", "content": user}],
                response_format={"type": "json_object"},
                temperature=0.3,
            )
            return json.loads(r.choices[0].message.content)
        except Exception as e:
            if i == retries - 1:
                print(f"  [LLM 실패] {e}", file=sys.stderr)
                return {}
            time.sleep(2 * (i + 1))
    return {}


# ----------------------------------------------------------------------
# 1단계 : 장르 경로 복원
# ----------------------------------------------------------------------
def load_cache(path=GENRE_CACHE) -> dict:
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache: dict, path=GENRE_CACHE):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def fetch_genre_path(session, genre_id) -> str:
    """
    genreId 하나로 '靴 > レディース靴 > スニーカー' 형태의 전체 경로를 복원.
    ancestors(상위 장르) + current(현재 장르)를 이어붙입니다.
    """
    params = {
        "applicationId": APP_ID,
        "accessKey": ACCESS_KEY,
        "genreId": genre_id,
        "format": "json",
    }
    try:
        r = session.get(GENRE_API, params=params, timeout=20)
        if r.status_code != 200:
            return ""
        d = r.json()
    except Exception:
        return ""

    def name_of(node):
        if isinstance(node, dict):
            inner = node.get("parent") or node.get("child") or node
            return inner.get("genreName") or inner.get("nameJa") or ""
        return ""

    parts = []
    for anc in (d.get("parents") or d.get("ancestors") or []):
        n = name_of(anc)
        if n:
            parts.append(n)

    cur = d.get("current") or d.get("genre") or {}
    if isinstance(cur, list):
        cur = cur[0] if cur else {}
    n = cur.get("genreName") or cur.get("nameJa") or ""
    if n:
        parts.append(n)

    return " > ".join(parts)


# ----------------------------------------------------------------------
# 스마트스토어 카테고리
# ----------------------------------------------------------------------
def load_categories(path: str) -> pd.DataFrame:
    """
    컬럼 구조가 제각각이라 자동 판별:
      - 값이 대부분 숫자인 열 -> 카테고리ID
      - 나머지 문자열 열들을 이어붙여 -> 전체경로
    """
    if path.endswith((".xlsx", ".xls")):
        df = pd.read_excel(path, dtype=str)
    else:
        try:
            df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            df = pd.read_csv(path, dtype=str, encoding="cp949")
    df = df.fillna("")

    id_col = None
    for c in df.columns:
        s = df[c].astype(str).str.strip()
        digits = s.str.fullmatch(r"\d{4,}").mean()
        if digits > 0.8:
            id_col = c
            break
    if id_col is None:
        id_col = df.columns[0]

    text_cols = [c for c in df.columns if c != id_col]
    path_series = (df[text_cols].astype(str)
                   .apply(lambda r: " > ".join([x.strip() for x in r if x.strip()]), axis=1))

    out = pd.DataFrame({"cat_id": df[id_col].astype(str).str.strip(),
                        "cat_path": path_series})
    out = out[(out.cat_id != "") & (out.cat_path != "")].drop_duplicates("cat_id")
    print(f"카테고리 {len(out)}건 로드 (ID열: {id_col})")
    return out.reset_index(drop=True)


def _norm(s: str) -> str:
    return re.sub(r"[^가-힣A-Za-z0-9]", "", str(s or ""))


def _bigrams(s: str) -> set:
    s = _norm(s)
    if len(s) < 2:
        return {s} if s else set()
    return {s[i:i + 2] for i in range(len(s) - 1)}


def _dice(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    return 2 * len(a & b) / (len(a) + len(b))


def _leaf(path: str) -> str:
    return str(path).split(">")[-1].strip()


def shortlist(cat_df: pd.DataFrame, query: str, n: int = 30) -> list[tuple]:
    """
    번역된 장르 경로와 가까운 카테고리 후보를 추린다 (LLM 토큰 절약용).

    단어 단위 비교로는 '헤어드라이어' vs '드라이어', '머그컵' vs '머그' 처럼
    한쪽이 다른 쪽을 품는 경우를 놓친다. 문자 2-gram 유사도로 비교하고,
    실제 상품 분류를 좌우하는 말단(leaf) 이름에 가중치를 크게 둔다.
    """
    q_full = _bigrams(query)
    q_leaf_txt = _norm(_leaf(query))
    q_leaf = _bigrams(q_leaf_txt)

    scored = []
    for cid, path in zip(cat_df.cat_id, cat_df.cat_path):
        leaf_txt = _norm(_leaf(path))
        score = (0.35 * _dice(q_full, _bigrams(path))
                 + 0.65 * _dice(q_leaf, _bigrams(leaf_txt)))
        # 한쪽이 다른 쪽을 포함하면 강한 신호
        if q_leaf_txt and leaf_txt:
            if q_leaf_txt in leaf_txt or leaf_txt in q_leaf_txt:
                score += 0.35
        if q_leaf_txt and q_leaf_txt in _norm(path):
            score += 0.15
        if score > 0.05:
            scored.append((score, str(cid), path))

    scored.sort(key=lambda x: -x[0])
    return [(c, p) for _, c, p in scored[:n]]


GENRE_SYS = """너는 일본 라쿠텐 상품을 한국 오픈마켓에 등록하는 구매대행 셀러의 보조다.
일본어 상품 장르 경로를 한국어로 번역하고, 주어진 스마트스토어 카테고리 후보 중
가장 알맞은 하나를 고른다.
반드시 JSON만 출력한다: {"ko_path": "...", "cat_id": "...", "cat_path": "...", "confidence": 0.0~1.0}
후보 중 적절한 게 없으면 cat_id를 빈 문자열로 둔다."""


def map_genres(client, session, genre_ids, cat_df, cache: dict):
    todo = [g for g in genre_ids if str(g) not in cache]
    if not todo:
        return cache
    print(f"\n[1단계] 신규 장르 {len(todo)}개 매핑 중...")

    for i, gid in enumerate(todo, 1):
        ja_path = fetch_genre_path(session, gid)
        time.sleep(1.0)
        if not ja_path:
            cache[str(gid)] = {"ja_path": "", "ko_path": "", "cat_id": "",
                               "cat_path": "", "confidence": 0}
            continue

        if cat_df is None:
            res = ask_json(client, GENRE_SYS,
                           f"장르경로: {ja_path}\n후보: (없음, 번역만 하라)")
            res.setdefault("cat_id", "")
        else:
            # 1차 번역으로 검색 쿼리 확보
            rough = ask_json(client,
                             '일본어를 한국어로 번역. JSON만: {"ko":"..."}',
                             ja_path).get("ko", ja_path)
            cands = shortlist(cat_df, rough)
            cand_txt = "\n".join(f"{c} | {p}" for c, p in cands) or "(없음)"
            res = ask_json(client, GENRE_SYS,
                           f"장르경로(일본어): {ja_path}\n\n카테고리 후보:\n{cand_txt}")

        res["ja_path"] = ja_path
        cache[str(gid)] = res
        print(f"  {i}/{len(todo)}  {ja_path}  ->  {res.get('cat_path','')} "
              f"({res.get('cat_id','')})")
        save_cache(cache)

    return cache


# ----------------------------------------------------------------------
# 2단계 : 상품 번역
# ----------------------------------------------------------------------
ITEM_SYS = f"""너는 일본 구매대행 상품을 한국 오픈마켓에 등록하는 전문 MD다.
일본어 상품 정보를 받아 한국 판매용으로 가공한다.

상품명 규칙:
- 형식: [브랜드] [제품명/모델] [핵심속성] [상품유형]
- 40~50자 내외. 같은 단어 반복 금지. 특수문자와 이모지 금지.
- 일본어를 남기지 말 것. 브랜드는 널리 쓰이는 한글 표기 우선, 애매하면 영문 유지.
- 다음 표현은 절대 쓰지 말 것(표시광고법 위반 소지): {", ".join(BANNED)}
- 검색될 만한 실제 소비자 용어를 넣을 것.

상세설명 규칙:
- 자연스러운 한국어. 원문에 없는 사실을 지어내지 말 것.
- 사이즈/소재/원산지/수량 등 구매 판단에 필요한 정보 위주로 정리.
- 일본 내수 사양(전압 100V, 사이즈 표기 차이 등)이 원문에 있으면 반드시 살릴 것.
- 800자 이내.

태그: 검색 유입용 키워드 10개. 상품명과 중복되는 단어는 피하고 연관검색어 위주로.

반드시 JSON만 출력:
{{"name": "...", "desc": "...", "tags": ["..."], "brand": "...", "warning": "..."}}
warning에는 통관/전압/사이즈 등 셀러가 주의할 점을 한 줄로. 없으면 빈 문자열."""


def strip_banned(text: str) -> str:
    """
    금지어를 부분 치환하면 '최고급' -> '급' 처럼 깨진 조각이 남는다.
    금지어를 포함한 어절 자체를 통째로 제거한다.
    """
    keep = [t for t in text.split()
            if not any(b in t for b in BANNED)]
    return re.sub(r"\s+", " ", " ".join(keep)).strip()


def enrich_one(client, row: dict) -> dict:
    src = (f"상품명: {row.get('상품명_원문','')}\n"
           f"캐치카피: {row.get('캐치카피','')}\n"
           f"샵: {row.get('샵명','')}\n"
           f"가격: {row.get('가격_엔','')}엔\n"
           f"요약: {str(row.get('요약설명',''))[:800]}\n"
           f"상세: {str(row.get('상세설명',''))[:1500]}")
    res = ask_json(client, ITEM_SYS, src)

    name = strip_banned(str(res.get("name", "")))[:100]

    tags = res.get("tags", [])
    if isinstance(tags, str):
        tags = [t.strip() for t in tags.split(",")]

    return {
        "판매상품명": name,
        "브랜드": str(res.get("brand", "")).strip(),
        "상세설명_한글": str(res.get("desc", "")).strip(),
        "태그": ",".join([str(t).strip() for t in tags if str(t).strip()][:10]),
        "주의사항": str(res.get("warning", "")).strip(),
    }


def enrich_all(client, rows: list[dict]) -> list[dict]:
    print(f"\n[2단계] 상품 {len(rows)}건 번역 중 (병렬 {WORKERS})...")
    out = [None] * len(rows)
    done = 0
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(enrich_one, client, r): i for i, r in enumerate(rows)}
        for f in as_completed(futs):
            i = futs[f]
            try:
                out[i] = f.result()
            except Exception as e:
                print(f"  [실패 {i}] {e}", file=sys.stderr)
                out[i] = {}
            done += 1
            if done % 10 == 0 or done == len(rows):
                print(f"  {done}/{len(rows)}")
    return out


# ----------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser()
    p.add_argument("infile", help="rakuten_collector.py 출력 엑셀")
    p.add_argument("--cat", help="스마트스토어 카테고리 CSV/XLSX")
    p.add_argument("--no-category", action="store_true", help="번역만 수행")
    p.add_argument("--out", default="final.xlsx")
    p.add_argument("--limit", type=int, default=None, help="테스트용 상위 N건만")
    a = p.parse_args()

    df = pd.read_excel(a.infile, dtype=str).fillna("")
    if a.limit:
        df = df.head(a.limit)
    print(f"입력 {len(df)}건")

    client = get_client()
    session = requests.Session()

    # 1단계
    if not a.no_category:
        cat_df = load_categories(a.cat) if a.cat else None
        if cat_df is None:
            print("경고: --cat 미지정. 장르 번역만 하고 카테고리ID는 비웁니다.")
        cache = load_cache()
        gids = sorted({str(g) for g in df.get("장르ID", []) if str(g).strip()})
        cache = map_genres(client, session, gids, cat_df, cache)

        df["장르_한글"] = df["장르ID"].astype(str).map(
            lambda g: cache.get(g, {}).get("ko_path", ""))
        df["카테고리ID"] = df["장르ID"].astype(str).map(
            lambda g: cache.get(g, {}).get("cat_id", ""))
        df["카테고리경로"] = df["장르ID"].astype(str).map(
            lambda g: cache.get(g, {}).get("cat_path", ""))
        df["매핑신뢰도"] = df["장르ID"].astype(str).map(
            lambda g: cache.get(g, {}).get("confidence", 0))

    # 2단계
    enriched = enrich_all(client, df.to_dict("records"))
    for key in ["판매상품명", "브랜드", "상세설명_한글", "태그", "주의사항"]:
        df[key] = [e.get(key, "") for e in enriched]

    # 보기 좋은 순서로 재배치
    head = ["판매상품명", "브랜드", "카테고리ID", "카테고리경로", "매핑신뢰도",
            "판매가_원", "가격_엔", "총원가_원", "예상마진_원", "태그", "주의사항"]
    cols = [c for c in head if c in df.columns] + \
           [c for c in df.columns if c not in head]
    df = df[cols]

    df.to_excel(a.out, index=False)
    print(f"\n저장 완료: {a.out}")

    if "매핑신뢰도" in df.columns:
        low = df[pd.to_numeric(df["매핑신뢰도"], errors="coerce").fillna(0) < 0.7]
        if len(low):
            print(f"주의: 카테고리 신뢰도 0.7 미만 {len(low)}건 -> 수동 검수 권장")


if __name__ == "__main__":
    main()
