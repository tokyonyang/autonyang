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
import os
import re
import sys
import time
from dataclasses import dataclass, field, asdict

import requests
from bs4 import BeautifulSoup

API_URL = "https://openapi.rakuten.co.jp/ichibams/api/IchibaItem/Search/20260701"

APP_ID = os.environ.get("RAKUTEN_APP_ID", "")
ACCESS_KEY = os.environ.get("RAKUTEN_ACCESS_KEY", "")

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
    s = requests.Session()
    s.headers.update({"User-Agent": UA})
    return s


def api_call(session: requests.Session, **params) -> dict:
    """Item Search API 호출. 429는 백오프 후 재시도."""
    if not APP_ID or not ACCESS_KEY:
        sys.exit("RAKUTEN_APP_ID / RAKUTEN_ACCESS_KEY 환경변수를 설정하세요.")

    base = {
        "applicationId": APP_ID,
        "accessKey": ACCESS_KEY,   # 신규 필수 파라미터
        "format": "json",
        "formatVersion": 2,        # items[0].itemName 형태로 평탄화
        "hits": 30,                # 페이지당 최대 30
    }
    base.update({k: v for k, v in params.items() if v is not None})

    for attempt in range(5):
        r = session.get(API_URL, params=base, timeout=20)
        if r.status_code == 200:
            return r.json()
        if r.status_code == 429:
            wait = 5 * (attempt + 1)
            print(f"  [429] 요청 과다. {wait}초 대기...", file=sys.stderr)
            time.sleep(wait)
            continue
        if r.status_code == 404:
            return {"items": [], "count": 0}
        raise RuntimeError(f"API {r.status_code}: {r.text[:300]}")
    raise RuntimeError("재시도 한도 초과")


def collect(session, max_items=100, **params):
    """페이지네이션 수집. page는 최대 100 -> 검색당 이론상 3000개."""
    out, page = [], 1
    while len(out) < max_items and page <= 100:
        data = api_call(session, page=page, **params)
        items = data.get("items", [])
        if not items:
            break
        out.extend(items)
        total = data.get("count", 0)
        print(f"  page {page} / 누적 {len(out)}건 (전체 {total}건)", file=sys.stderr)
        if page >= data.get("pageCount", 1):
            break
        page += 1
        time.sleep(API_DELAY)
    return out[:max_items]


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

    s1 = sub.add_parser("search", help="키워드 검색")
    s1.add_argument("keyword")
    s1.add_argument("--genre", type=int, default=None)
    common(s1)

    s2 = sub.add_parser("shop", help="특정 샵 상품 수집")
    s2.add_argument("shopcode")
    common(s2)

    s3 = sub.add_parser("urls", help="URL 목록 파일로 개별 수집")
    s3.add_argument("file")
    common(s3)

    a = p.parse_args()
    cfg = PriceConfig(jpy_krw=a.rate)
    session = _session()

    if a.cmd == "search":
        print(f"[검색] {a.keyword}")
        items = collect(session, max_items=a.max, keyword=a.keyword,
                        genreId=a.genre, minPrice=a.min_price,
                        maxPrice=a.max_price, imageFlag=1, availability=1)
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
                items.append(got[0])
            else:
                print(f"  [미발견] {code}", file=sys.stderr)
            print(f"  {i}/{len(codes)}", file=sys.stderr)
            time.sleep(API_DELAY)

    print(f"\nAPI 수집 {len(items)}건")

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
