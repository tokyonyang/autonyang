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

# 라쿠텐 요청에 필요한 세션(Referer/Origin 헤더 포함)과 앱 키를 재사용한다.
# 2026년 API 개편으로 이 헤더가 없으면 조용히 403이 나기 때문에 중요하다.
try:
    from rakuten_collector import _session as rk_session, APP_ID, ACCESS_KEY
except ImportError:
    sys.exit("rakuten_collector.py 를 같은 폴더에 두세요.")

GENRE_API = "https://openapi.rakuten.co.jp/ichibagt/api/IchibaGenre/Search/20260401"

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

    실제 응답 구조 (2026-04 버전):
      {"genre": {"genreId":.., "nameJa":..},
       "ancestors": [{"genreId":.., "nameJa":..}, ...],  # 최상위부터 순서대로
       "children": [...]}
    """
    params = {"applicationId": APP_ID, "accessKey": ACCESS_KEY,
              "genreId": genre_id, "format": "json"}
    try:
        r = session.get(GENRE_API, params=params, timeout=20)
        if r.status_code != 200:
            return ""
        d = r.json()
    except Exception:
        return ""

    parts = [a.get("nameJa", "") for a in (d.get("ancestors") or []) if a.get("nameJa")]
    cur_name = (d.get("genre") or {}).get("nameJa", "")
    if cur_name:
        parts.append(cur_name)

    return " > ".join(parts)


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

상품명 규칙 (네이버 쇼핑 검색 적합도 최우선):
- 형식: [브랜드] [카테고리 대표명사] [제품명/모델] [핵심속성]
- 카테고리 대표명사는 실제 소비자가 검색창에 치는 일반명사를 맨 앞쪽에 둘 것.
  네이버 적합도 알고리즘은 상품명 앞부분 단어에 가중치를 더 준다.
  (예: "나이키 스니커즈 에어맥스 화이트" O, "나이키 에어맥스 화이트 신발" X)
  입력으로 주어지는 '카테고리' 정보를 참고해 그 카테고리에서 실제로
  검색될 법한 명사를 고를 것. 카테고리 정보가 없으면 제품 유형에서 유추한다.
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
(현재 CSV 일괄등록 템플릿은 태그 컬럼이 없어 업로드에는 반영되지 않지만,
추후 커머스 API 연동 시 바로 쓸 수 있도록 계속 생성해 둔다.)

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
    # 1단계에서 확정된 스마트스토어 카테고리 경로(없으면 라쿠텐 장르 번역)를
    # 상품명 생성 프롬프트에 같이 넘긴다. 이게 있어야 LLM이 그 카테고리에서
    # 실제로 검색되는 대표명사(예: 스니커즈, 원피스)를 상품명 앞쪽에 넣는다.
    category = row.get("카테고리경로") or row.get("장르_한글") or ""

    src = (f"상품명: {row.get('상품명_원문','')}\n"
           f"캐치카피: {row.get('캐치카피','')}\n"
           f"카테고리: {category}\n"
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
    session = rk_session()

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
        vals = [e.get(key, "") for e in enriched]
        if key == "브랜드" and "브랜드" in df.columns:
            # 수집 단계에서 화이트리스트로 확정한 브랜드가 있으면 그쪽을 신뢰한다.
            # LLM 추측보다 정확하고 표기도 일관된다.
            vals = [old.strip() or new
                    for old, new in zip(df["브랜드"].astype(str), vals)]
        df[key] = vals

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
