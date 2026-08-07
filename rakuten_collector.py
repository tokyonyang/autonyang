#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
라쿠텐(楽天市場) 구매대행 상품 수집기
--------------------------------------------------
- 공식 Rakuten Ichiba Item Search API (version 2026-07-01) 사용
- API로 안 나오는 고화질 이미지 / 상세설명 / 옵션은 상품페이지에서 보강
- 결과를 엑셀(xlsx)로 저장 -> 마켓 대량등록 양식으로 가공

사용 전:
  1) https://webservice.rakuten.co.jp/app/create 에서 앱 생성
  2) applicationId + accessKey 발급
  3) 환경변수 설정
     export RAKUTEN_APP_ID="..."
     export RAKUTEN_ACCESS_KEY="..."

설치:
  pip install requests beautifulsoup4 pandas openpyxl lxml

예시:
  # 키워드로 100개 수집
  python rakuten_collector.py search "スニーカー" --max 100 --out out.xlsx

  # 특정 샵 전체 수집
  python rakuten_collector.py shop "shopcode" --max 300

  # URL 목록 파일로 개별 수집 (한 줄에 하나씩)
  python rakuten_collector.py urls urls.txt --images
"""

import argparse
import json
import os
import re
import sys
import time
import unicodedata
from dataclasses import dataclass, field, asdict

import requests
from bs4 import BeautifulSoup

API_URL = "https://openapi.rakuten.co.jp/ichibams/api/IchibaItem/Search/20260701"
RANK_URL = "https://openapi.rakuten.co.jp/ichibaranking/api/IchibaItem/Ranking/20220601"

APP_ID = os.environ.get("RAKUTEN_APP_ID", "")
ACCESS_KEY = os.environ.get("RAKUTEN_ACCESS_KEY", "")
# 앱 등록 시 입력한 URL과 정확히 일치해야 한다. 다르면 403 HTTP_REFERRER_NOT_ALLOWED.
# https://webservice.rakuten.co.jp/app/list 에서 본인 앱의 등록 URL을 확인해 설정.
REFERRER = os.environ.get("RAKUTEN_REFERRER", "")

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# API 권장 호출 간격(초). 너무 빠르면 429가 납니다.
API_DELAY = 1.0
PAGE_DELAY = 1.5  # 상품페이지 스크래핑 간격


# ----------------------------------------------------------------------
# 환율 / 마진 설정  (여기 숫자만 본인 기준으로 바꾸세요)
# ----------------------------------------------------------------------
@dataclass
class PriceConfig:
    jpy_krw: float = 9.5          # 100엔당이 아니라 1엔당 원화
    domestic_ship_jpy: int = 800  # 일본 내 배송비(엔)
    intl_ship_krw: int = 9000     # 국제배송비(원) - 무게별로 조정
    duty_rate: float = 0.0        # 관세율 (품목별. 목록통관 면세면 0)
    vat_rate: float = 0.10        # 부가세
    margin_rate: float = 0.30     # 목표 마진율
    market_fee: float = 0.06      # 마켓 수수료 (스마트스토어 등)
    round_to: int = 100           # 판매가 반올림 단위


def calc_sale_price(jpy: int, cfg: PriceConfig) -> dict:
    """엔화 원가 -> 한국 판매가 계산."""
    cost_jpy = jpy + cfg.domestic_ship_jpy
    cost_krw = cost_jpy * cfg.jpy_krw
    duty = cost_krw * cfg.duty_rate
    vat = (cost_krw + duty) * cfg.vat_rate
    landed = cost_krw + duty + vat + cfg.intl_ship_krw

    # 마진과 마켓수수료를 판매가 기준으로 역산
    denom = 1 - cfg.margin_rate - cfg.market_fee
    sale = landed / denom if denom > 0 else landed * 2

    sale = int(round(sale / cfg.round_to) * cfg.round_to)
    return {
        "원가_원": int(round(cost_krw)),
        "관세_원": int(round(duty)),
        "부가세_원": int(round(vat)),
        "총원가_원": int(round(landed)),
        "판매가_원": sale,
        "예상마진_원": int(round(sale * (1 - cfg.market_fee) - landed)),
    }


# ----------------------------------------------------------------------
# API 호출
# ----------------------------------------------------------------------
def _session() -> requests.Session:
    if not REFERRER:
        sys.exit(
            "RAKUTEN_REFERRER 환경변수가 설정되지 않았습니다.\n"
            "  2026년 API 개편으로 Referer 헤더가 필수가 됐고,\n"
            "  값은 앱을 만들 때 입력한 URL과 정확히 일치해야 합니다.\n"
            "  https://webservice.rakuten.co.jp/app/list 에서\n"
            "  본인 앱의 '이용할 URL'을 확인한 뒤:\n"
            '    export RAKUTEN_REFERRER="https://example.com"\n'
            "  처럼 그대로 설정해 주세요. (뒤에 슬래시 유무까지 동일하게)")
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        # 2026년 API 개편으로 Referer/Origin 헤더가 없으면 403이 난다.
        # 등록 URL과 다르면 HTTP_REFERRER_NOT_ALLOWED 로 역시 403.
        "Referer": REFERRER,
        "Origin": REFERRER.rstrip("/"),
    })
    return s


def api_call(session: requests.Session, _url: str = API_URL, **params) -> dict:
    """Item Search / Ranking API 호출. 429는 백오프 후 재시도."""
    _check_keys()

    base = {
        "applicationId": APP_ID,
        "accessKey": ACCESS_KEY,   # 신규 필수 파라미터
        "format": "json",
        "formatVersion": 2,        # items[0].itemName 형태로 평탄화
        "hits": 30,                # 페이지당 최대 30
    }
    base.update({k: v for k, v in params.items() if v is not None})

    for attempt in range(5):
        r = session.get(_url, params=base, timeout=20)
        if r.status_code == 200:
            return r.json()
        if r.status_code == 429:
            wait = 5 * (attempt + 1)
            print(f"  [429] 요청 과다. {wait}초 대기...", file=sys.stderr)
            time.sleep(wait)
            continue
        if r.status_code == 404:
            return {"items": [], "count": 0}
        sys.exit(f"API 오류 [{r.status_code}] {_api_error_text(r)}")
    raise RuntimeError("재시도 한도 초과")


def collect(session, max_items=100, **params):
    """페이지네이션 수집. page는 최대 100 -> 검색당 이론상 3000개."""
    out, page = [], 1
    while len(out) < max_items and page <= 100:
        data = api_call(session, page=page, **params)
        raw = data.get("items", [])
        if not raw:
            break
        # 방어적 언랩: formatVersion=2 를 줘도 {"item": {...}} 로 감싸져
        # 오는 경우가 실제로 있었다 (랭킹 API에서 확인됨). 검색 API도 만약을
        # 대비해 같은 방식으로 처리한다.
        items = [dict(e["item"]) if isinstance(e, dict) and "item" in e else e
                 for e in raw]
        out.extend(items)
        total = data.get("count", 0)
        print(f"  page {page} / 누적 {len(out)}건 (전체 {total}건)", file=sys.stderr)
        if page >= data.get("pageCount", 1):
            break
        page += 1
        time.sleep(API_DELAY)
    return out[:max_items]


def collect_ranking(session, max_items=100, genre=None, age=None,
                    sex=None, realtime=False):
    """
    오늘의 베스트셀러(랭킹 시장) 수집.

    API 제약 (문서 기준):
      - genreId 와 age/sex 는 동시 지정 불가
      - page 는 1~34 (페이지당 30위, 최대 1000위)
      - period=realtime 을 age 와 같이 쓰면 age 는 20/30/40 만 허용

    응답 구조 주의: 검색 API와 달리 랭킹 API는 formatVersion=2 를 줘도
    각 항목이 {"rank": N, "item": {...}} 형태로 한 겹 더 감싸져 온다.
    검색 API처럼 바로 item.get("itemName") 이 안 되고 래핑을 벗겨야 한다.
    """
    if genre and (age or sex is not None):
        sys.exit("장르와 연령/성별은 동시에 지정할 수 없습니다. 하나만 쓰세요.")
    if realtime and age and int(age) not in (20, 30, 40):
        sys.exit("realtime 모드에서 연령은 20/30/40 만 가능합니다.")

    out, page = [], 1
    while len(out) < max_items and page <= 34:
        params = {"page": page}
        if genre:
            params["genreId"] = genre
        if age:
            params["age"] = age
        if sex is not None:
            params["sex"] = sex
        if realtime:
            params["period"] = "realtime"

        data = api_call(session, RANK_URL, **params)
        raw = data.get("items", [])
        if not raw:
            break

        # {"rank":N, "item":{...}} 래핑을 벗기고 rank를 item 안으로 합친다
        page_items = []
        for entry in raw:
            if isinstance(entry, dict) and "item" in entry:
                it = dict(entry["item"])
                it["rank"] = entry.get("rank", it.get("rank", ""))
                page_items.append(it)
            else:
                page_items.append(entry)

        out.extend(page_items)
        print(f"  page {page} / 누적 {len(out)}건", file=sys.stderr)
        page += 1
        time.sleep(API_DELAY)

    return out[:max_items]


# ----------------------------------------------------------------------
# 브랜드 필터
# ----------------------------------------------------------------------
# 정품 브랜드 상품으로 보기 어려운 상품을 걸러냅니다.
EXCLUDE_WORDS = [
    "中古", "ジャンク", "訳あり", "わけあり", "福袋", "まとめ売り",
    "詰め合わせ", "レプリカ", "コピー", "互換", "非純正", "模造",
    "ノーブランド", "アウトレット品", "USED", "リユース",
]


def _nfkc(s: str) -> str:
    """전각/반각, 대소문자, 공백 차이를 없애고 비교용으로 정규화."""
    return unicodedata.normalize("NFKC", str(s or "")).lower().replace(" ", "")


def load_brands(path: str):
    """
    브랜드 목록 파일. 한 줄에 브랜드 하나, 별칭은 | 로 구분.
      유니클로|UNIQLO|ユニクロ
    첫 항목을 대표 표기로 씁니다. # 로 시작하면 주석.
    """
    brands = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            names = [x.strip() for x in line.split("|") if x.strip()]
            if names:
                brands.append((names[0], [_nfkc(x) for x in names]))
    return brands


def match_brand(item: dict, brands):
    """
    상품명/캐치카피에서 브랜드를 찾아 대표 표기를 반환. 없으면 None.
    라쿠텐 상품명은 【】[] 같은 괄호와 전각문자가 뒤섞여 있어 NFKC 정규화가 필수입니다.
    """
    hay = _nfkc(f"{item.get('itemName','')} {item.get('catchcopy','')}")
    for canonical, aliases in brands:
        if any(a and a in hay for a in aliases):
            return canonical
    return None


def is_excluded(item: dict) -> bool:
    """중고/훼손품/럭키백 등 정품 신품으로 팔 수 없는 상품인지."""
    hay = _nfkc(f"{item.get('itemName','')} {item.get('catchcopy','')}")
    return any(_nfkc(w) in hay for w in EXCLUDE_WORDS)


def filter_items(items, brands=None, exclude_used=True):
    """브랜드 화이트리스트 + 제외어 필터. (통과목록, 통계) 반환."""
    kept, stat = [], {"제외어": 0, "브랜드없음": 0}
    for it in items:
        if exclude_used and is_excluded(it):
            stat["제외어"] += 1
            continue
        if brands:
            b = match_brand(it, brands)
            if not b:
                stat["브랜드없음"] += 1
                continue
            it["_brand"] = b
        kept.append(it)
    return kept, stat


# ----------------------------------------------------------------------
# 장르 필터
# ----------------------------------------------------------------------
GENRE_URL = "https://openapi.rakuten.co.jp/ichibagt/api/IchibaGenre/Search/20260401"
ANCESTRY_CACHE = "genre_ancestry.json"


def _api_error_text(r: requests.Response) -> str:
    """
    라쿠텐 에러 응답은 두 형식이 섞여 있다.
      신형: {"errors": {"errorCode": 403, "errorMessage": "..."}}
      구형: {"error": "wrong_parameter", "error_description": "..."}
    """
    try:
        e = r.json()
    except Exception:
        return r.text[:200]

    if "errors" in e and isinstance(e["errors"], dict):
        inner = e["errors"]
        return f"{inner.get('errorCode','')}: {inner.get('errorMessage','')}"
    if "error" in e:
        return f"{e.get('error','')}: {e.get('error_description','')}"
    return str(e)[:200]


def _check_keys():
    """키가 비었거나 자리표시자 그대로면 즉시 알려준다."""
    bad = {"", "...", "your_app_id", "xxx"}
    if APP_ID.strip() in bad or ACCESS_KEY.strip() in bad:
        sys.exit(
            "RAKUTEN_APP_ID / RAKUTEN_ACCESS_KEY 가 설정되지 않았습니다.\n"
            "  https://webservice.rakuten.co.jp/app/list 에서 실제 값을 확인해\n"
            '  export RAKUTEN_APP_ID="1234567890123456789"\n'
            '  export RAKUTEN_ACCESS_KEY="abcdef..."\n'
            "  처럼 따옴표 안에 실제 값을 넣어 주세요.")


def _genre_call(session, genre_id):
    _check_keys()
    params = {"applicationId": APP_ID, "accessKey": ACCESS_KEY,
              "genreId": genre_id, "format": "json"}
    r = session.get(GENRE_URL, params=params, timeout=20)
    if r.status_code != 200:
        sys.exit(f"장르 API 오류 [{r.status_code}] {_api_error_text(r)}")
    return r.json()


def list_genres(session, genre_id=0, raw=False):
    """
    장르 트리 한 단계를 출력. 허용할 장르ID를 눈으로 고르기 위한 용도.

    실제 응답 구조 (2026-04 버전):
      {"genre": {"genreId":.., "nameJa":.., "level":..},
       "ancestors": [{"genreId":.., "nameJa":..}, ...],
       "children":  [{"genreId":.., "nameJa":.., "level":..}, ...]}
    문서상의 genreName/current/child 같은 필드명이 아니라 nameJa/genre 를 쓴다.
    """
    d = _genre_call(session, genre_id)

    if raw:
        print(json.dumps(d, ensure_ascii=False, indent=2)[:4000])
        return

    cur = d.get("genre") or {}
    ancestors = d.get("ancestors") or []
    if ancestors:
        trail = " > ".join(a.get("nameJa", "") for a in ancestors)
        print(f"\n경로: {trail}")
    if cur.get("genreId") is not None:
        print(f"현재: [{cur.get('genreId','')}] {cur.get('nameJa','') or '(최상위)'}")

    children = d.get("children") or []
    if not children:
        print("  하위 장르 없음 (여기가 말단 장르입니다)")
        return

    print(f"\n{'장르ID':>10}  장르명")
    print("  " + "-" * 44)
    for n in children:
        print(f"{str(n.get('genreId','')):>10}  {n.get('nameJa','')}")
    print(f"\n총 {len(children)}개. 더 내려가려면 --genre <장르ID> 를 붙여 다시 실행하세요.")


def _load_ancestry_cache():
    if os.path.exists(ANCESTRY_CACHE):
        try:
            with open(ANCESTRY_CACHE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_ancestry_cache(cache):
    with open(ANCESTRY_CACHE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def genre_ancestry(session, genre_id, cache):
    """
    genreId의 조상 ID 목록(자기 자신 포함)을 반환. 조회 결과는 캐시.
    응답의 ancestors 는 최상위부터 순서대로 오는 평면 리스트다.
    """
    key = str(genre_id)
    if key in cache:
        return cache[key]

    d = _genre_call(session, genre_id)
    ids = [str(a["genreId"]) for a in (d.get("ancestors") or []) if a.get("genreId") is not None]
    ids.append(key)

    cache[key] = ids
    _save_ancestry_cache(cache)
    time.sleep(API_DELAY)
    return ids


def filter_by_genre(session, items, allowed):
    """허용 장르(및 그 하위)에 속한 상품만 남긴다."""
    allowed = {str(x).strip() for x in allowed if str(x).strip()}
    if not allowed:
        return items, 0

    cache = _load_ancestry_cache()
    kept, dropped = [], 0
    for it in items:
        gid = it.get("genreId")
        if not gid:
            dropped += 1
            continue
        if set(genre_ancestry(session, gid, cache)) & allowed:
            kept.append(it)
        else:
            dropped += 1
    return kept, dropped


# ----------------------------------------------------------------------
# 이미지 / URL 유틸
# ----------------------------------------------------------------------
_EX = re.compile(r"_ex=\d+x\d+")


def upscale(url: str, size: int = 700) -> str:
    """
    라쿠텐 썸네일 서버는 _ex 파라미터로 크기를 지정합니다.
    API가 주는 128x128을 700x700 등으로 올려서 받습니다.
    """
    if not url:
        return ""
    if "_ex=" in url:
        return _EX.sub(f"_ex={size}x{size}", url)
    sep = "&" if "?" in url else "?"
    return f"{url}{sep}_ex={size}x{size}"


def item_code_from_url(url: str) -> str | None:
    """https://item.rakuten.co.jp/{shop}/{id}/ -> 'shop:id'"""
    m = re.search(r"item\.rakuten\.co\.jp/([^/]+)/([^/?#]+)", url)
    return f"{m.group(1)}:{m.group(2)}" if m else None


def enrich_from_page(session, item: dict, max_images: int = 20) -> dict:
    """
    상품 상세페이지에서 API에 없는 정보를 보강.
    - 고화질 이미지 전체 (해당 샵 cabinet 경로 이미지만 필터)
    - 상세설명 텍스트
    라쿠텐은 샵마다 HTML이 제각각이라 '샵 cabinet 이미지 필터' 방식이 가장 안정적입니다.
    """
    url = item.get("itemUrl", "")
    shop = item.get("shopCode", "")
    if not url:
        return {}

    try:
        r = session.get(url, timeout=20)
        r.encoding = r.apparent_encoding
        soup = BeautifulSoup(r.text, "lxml")
    except Exception as e:
        print(f"  [경고] 페이지 로드 실패 {url}: {e}", file=sys.stderr)
        return {}

    imgs, seen = [], set()
    for tag in soup.find_all("img"):
        src = tag.get("src") or tag.get("data-src") or ""
        if not src.startswith("http"):
            continue
        # 해당 샵의 상품 이미지 저장소만 채택 (배너/아이콘 제외)
        if "image.rakuten.co.jp" not in src:
            continue
        if shop and f"/{shop}/" not in src:
            continue
        clean = _EX.sub("", src).rstrip("?&")
        if clean in seen:
            continue
        seen.add(clean)
        imgs.append(upscale(src, 700))
        if len(imgs) >= max_images:
            break

    desc = ""
    for sel in ("#item_desc", ".item_desc", "#sales-description",
                ".sales-description-text", "#itemDescription"):
        node = soup.select_one(sel)
        if node:
            desc = node.get_text("\n", strip=True)[:3000]
            break

    return {"detail_images": imgs, "detail_text": desc}


# ----------------------------------------------------------------------
# 정규화
# ----------------------------------------------------------------------
def normalize(item: dict, cfg: PriceConfig, extra: dict | None = None) -> dict:
    extra = extra or {}
    api_imgs = [upscale(u, 700) for u in (item.get("mediumImageUrls") or [])]
    detail_imgs = extra.get("detail_images") or []

    # 대표이미지는 API 것, 나머지는 상세페이지에서 보강
    all_imgs, seen = [], set()
    for u in api_imgs + detail_imgs:
        key = _EX.sub("", u)
        if key not in seen:
            seen.add(key)
            all_imgs.append(u)

    price = item.get("itemPrice", 0) or 0
    calc = calc_sale_price(price, cfg)

    return {
        "상품코드": item.get("itemCode", ""),
        "순위": item.get("rank", ""),
        "브랜드": item.get("_brand", ""),
        "상품명_원문": item.get("itemName", ""),
        "캐치카피": item.get("catchcopy", ""),
        "가격_엔": price,
        **calc,
        "재고": "있음" if item.get("availability") == 1 else "품절",
        "해외배송가능": "Y" if item.get("shipOverseasFlag") == 1 else "N",
        "리뷰수": item.get("reviewCount", 0),
        "평점": item.get("reviewAverage", 0),
        "샵명": item.get("shopName", ""),
        "샵코드": item.get("shopCode", ""),
        "장르ID": item.get("genreId", ""),
        "상품URL": item.get("itemUrl", ""),
        "대표이미지": all_imgs[0] if all_imgs else "",
        "이미지수": len(all_imgs),
        "이미지목록": "\n".join(all_imgs),
        "요약설명": (item.get("itemCaption") or "")[:1000],
        "상세설명": extra.get("detail_text", ""),
    }


def save(rows: list[dict], path: str):
    import pandas as pd
    df = pd.DataFrame(rows)
    if path.endswith(".csv"):
        df.to_csv(path, index=False, encoding="utf-8-sig")
    else:
        df.to_excel(path, index=False)
    print(f"\n저장 완료: {path}  ({len(rows)}건)")


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser(description="라쿠텐 구매대행 상품 수집기")
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(sp):
        sp.add_argument("--max", type=int, default=100, help="최대 수집 개수")
        sp.add_argument("--out", default="rakuten_items.xlsx")
        sp.add_argument("--images", action="store_true",
                        help="상품페이지까지 방문해 고화질 이미지/상세설명 보강 (느림)")
        sp.add_argument("--min-price", type=int, default=None)
        sp.add_argument("--max-price", type=int, default=None)
        sp.add_argument("--rate", type=float, default=9.5, help="엔/원 환율")
        sp.add_argument("--brands", default=None,
                        help="브랜드 목록 파일. 지정하면 해당 브랜드 상품만 남김")
        sp.add_argument("--allow-genres", default=None,
                        help="허용 장르ID 목록. 쉼표로 구분. 하위 장르도 자동 포함")
        sp.add_argument("--allow-used", action="store_true",
                        help="중고/훼손품/럭키백도 포함 (기본은 제외)")

    s1 = sub.add_parser("search", help="키워드 검색")
    s1.add_argument("keyword")
    s1.add_argument("--genre", type=int, default=None)
    common(s1)

    s2 = sub.add_parser("shop", help="특정 샵 상품 수집")
    s2.add_argument("shopcode")
    common(s2)

    s4 = sub.add_parser("ranking", help="오늘의 베스트셀러(랭킹) 수집")
    s4.add_argument("--genre", type=int, default=None, help="장르ID (연령/성별과 동시 사용 불가)")
    s4.add_argument("--age", type=int, default=None, choices=[10, 20, 30, 40, 50])
    s4.add_argument("--sex", type=int, default=None, choices=[0, 1], help="0 남성 / 1 여성")
    s4.add_argument("--realtime", action="store_true", help="실시간 랭킹")
    common(s4)

    s5 = sub.add_parser("genres", help="장르 트리 조회 (허용 장르ID 확인용)")
    s5.add_argument("--genre", type=int, default=0, help="비우면 최상위 장르 목록")
    s5.add_argument("--raw", action="store_true", help="원본 JSON 응답을 그대로 출력 (디버깅용)")

    s3 = sub.add_parser("urls", help="URL 목록 파일로 개별 수집")
    s3.add_argument("file")
    common(s3)

    a = p.parse_args()
    session = _session()

    if a.cmd == "genres":
        list_genres(session, a.genre, raw=a.raw)
        return

    cfg = PriceConfig(jpy_krw=a.rate)

    if a.cmd == "search":
        print(f"[검색] {a.keyword}")
        items = collect(session, max_items=a.max, keyword=a.keyword,
                        genreId=a.genre, minPrice=a.min_price,
                        maxPrice=a.max_price, imageFlag=1, availability=1)
    elif a.cmd == "ranking":
        print("[랭킹] " + (f"장르 {a.genre}" if a.genre else
                          f"연령 {a.age or '전체'} / 성별 {a.sex if a.sex is not None else '전체'}"))
        items = collect_ranking(session, max_items=a.max, genre=a.genre,
                                age=a.age, sex=a.sex, realtime=a.realtime)
        # 랭킹 API는 가격 파라미터를 지원하지 않아 받은 뒤에 직접 거른다
        if a.min_price or a.max_price:
            n0 = len(items)
            lo, hi = a.min_price or 0, a.max_price or 10**9
            items = [x for x in items if lo <= (x.get("itemPrice") or 0) <= hi]
            print(f"  가격대 필터 {len(items)}/{n0}건", file=sys.stderr)
    elif a.cmd == "shop":
        print(f"[샵] {a.shopcode}")
        items = collect(session, max_items=a.max, shopCode=a.shopcode,
                        minPrice=a.min_price, maxPrice=a.max_price,
                        imageFlag=1, availability=1)
    else:
        codes = []
        with open(a.file, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                code = item_code_from_url(line) or line
                codes.append(code)
        print(f"[URL] {len(codes)}건")
        items = []
        for i, code in enumerate(codes, 1):
            d = api_call(session, itemCode=code)
            got = d.get("items", [])
            if got:
                first = got[0]
                if isinstance(first, dict) and "item" in first:
                    first = dict(first["item"])
                items.append(first)
            else:
                print(f"  [미발견] {code}", file=sys.stderr)
            print(f"  {i}/{len(codes)}", file=sys.stderr)
            time.sleep(API_DELAY)

    print(f"\nAPI 수집 {len(items)}건")

    if getattr(a, "allow_genres", None):
        allowed = a.allow_genres.split(",")
        n0 = len(items)
        items, dropped = filter_by_genre(session, items, allowed)
        print(f"장르 필터 통과 {len(items)}/{n0}건 (제외 {dropped})")

    brands = load_brands(a.brands) if getattr(a, "brands", None) else None
    if brands:
        print(f"브랜드 목록 {len(brands)}개 적용")
    before = len(items)
    items, stat = filter_items(items, brands,
                               exclude_used=not getattr(a, "allow_used", False))
    if before != len(items):
        print(f"필터 통과 {len(items)}/{before}건 "
              f"(제외어 {stat['제외어']} / 브랜드불일치 {stat['브랜드없음']})")
    if not items:
        print("\n조건에 맞는 상품이 없습니다. 브랜드 목록이나 필터를 확인하세요.")
        return

    rows = []
    for i, it in enumerate(items, 1):
        extra = {}
        if a.images:
            print(f"  이미지 보강 {i}/{len(items)}", file=sys.stderr)
            extra = enrich_from_page(session, it)
            time.sleep(PAGE_DELAY)
        rows.append(normalize(it, cfg, extra))

    save(rows, a.out)


if __name__ == "__main__":
    main()
