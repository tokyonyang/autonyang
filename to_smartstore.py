#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
enrich_ai.py 결과 -> 스마트스토어 '상품 일괄등록' 양식 변환기
--------------------------------------------------------------
네이버 템플릿 CSV의 93개 컬럼 순서/이름을 그대로 읽어와서,
수집·번역된 데이터를 해당 위치에 채워 넣습니다.

핵심 처리
  - 상품명/상품코드 금지문자 제거 및 길이 제한
  - 판매가 10원 단위 보정
  - 관부가세 자동 판정 (일본발 USD 150 이하 -> '부과 대상 아님')
  - 상세설명 HTML 자동 조립 + 해외구매대행 필수 고지문 삽입
  - 작성가이드(3~6행) 제거 후 데이터행부터 기록

실행
  python to_smartstore.py final.xlsx \\
      --template ExcelSaveTemplate_20260324_일괄등록_.csv \\
      --origin-code 0200038 \\
      --importer "디깅스팟" \\
      --ship-template 12345 \\
      --out upload.csv

주의(업로드 전 1회 설정 필요)
  1) 관부가세 항목은 '출고지 주소가 해외주소'일 때만 유효합니다.
     스마트스토어센터 > 배송비 템플릿에서 출고지를 일본으로 지정한 뒤
     그 템플릿코드를 --ship-template 으로 넘기세요.
  2) 원산지코드는 상품등록 화면의 '원산지 찾기' 팝업에서 일본 코드를 한 번
     확인해서 --origin-code 로 넘기세요. (기본값은 임의값이라 반드시 교체)
"""

import argparse
import csv
import html
import re
import sys

import pandas as pd

# 상품명/코드에 쓸 수 없는 문자 (템플릿 가이드 기준)
BAD_NAME = r'[\\*?"<>]'
BAD_CODE = r'[\\*?"<>,]'

# 관부가세 면세 기준 (미국발 200불, 그 외 150불)
DUTY_FREE_USD = 150


def clean_name(s: str, limit: int = 100) -> str:
    s = re.sub(BAD_NAME, "", str(s or ""))
    s = re.sub(r"\s+", " ", s).strip()
    return s[:limit]


def clean_code(s: str, limit: int = 30) -> str:
    s = re.sub(BAD_CODE, "", str(s or "")).strip()
    s = s.replace(":", "_")
    return s[:limit]


def round10(v) -> str:
    """판매가는 10원 단위, 10~999,999,990 범위."""
    try:
        n = int(round(float(v) / 10) * 10)
    except (TypeError, ValueError):
        return ""
    n = max(10, min(n, 999_999_990))
    return str(n)


def judge_duty(total_cost_krw, usd_krw: float) -> str:
    """
    판매가에 관부가세를 이미 녹여 팔 것을 전제로 판정.
    일본발이므로 USD 150 이하면 '부과 대상 아님'.
    """
    try:
        usd = float(total_cost_krw) / usd_krw
    except (TypeError, ValueError, ZeroDivisionError):
        return "부과 대상 아님"
    return "부과 대상 아님" if usd <= DUTY_FREE_USD else "관부가세 포함"


NOTICE = """
<div style="margin-top:30px;padding:20px;border:1px solid #ddd;font-size:14px;line-height:1.8">
<b>해외구매대행 안내</b><br>
· 본 상품은 해외(일본)에서 판매되는 상품을 구매대행하는 상품입니다.<br>
· 주문 후 현지 구매·검수·국제배송 절차를 거치며, 통상 7~14일 소요됩니다.<br>
· 현지 재고 상황에 따라 품절 시 주문이 취소될 수 있습니다.<br>
· 해외 배송 상품 특성상 단순 변심에 의한 반품·교환 시 왕복 국제배송비가 부과됩니다.<br>
· 상품 사양은 일본 내수 기준입니다. 전압(100V), 사이즈 표기, 언어 등을
구매 전 반드시 확인해 주세요.<br>
· 모니터 환경에 따라 색상이 실제와 다르게 보일 수 있습니다.
</div>
""".strip()


def build_detail(row: dict, img_base: str = "") -> str:
    """상세설명 HTML 조립. 이미지 + 번역 본문 + 고지문."""
    imgs = [u.strip() for u in str(row.get("이미지목록", "")).split("\n") if u.strip()]
    if img_base:
        # 자체 호스팅으로 치환할 경우 파일명만 떼어 붙임
        imgs = [f"{img_base.rstrip('/')}/{u.split('/')[-1].split('?')[0]}" for u in imgs]

    parts = ['<div style="max-width:860px;margin:0 auto">']

    desc = str(row.get("상세설명_한글", "")).strip()
    if desc:
        body = "<br>".join(html.escape(l) for l in desc.split("\n") if l.strip())
        parts.append(f'<div style="font-size:15px;line-height:1.9;padding:20px 10px">{body}</div>')

    for u in imgs:
        parts.append(f'<img src="{u}" style="max-width:100%;display:block" />')

    warn = str(row.get("주의사항", "")).strip()
    if warn:
        parts.append(
            '<div style="margin-top:20px;padding:14px;background:#fff8e1;'
            f'border-left:4px solid #ffb300;font-size:14px">{html.escape(warn)}</div>')

    parts.append(NOTICE)
    parts.append("</div>")
    return "".join(parts)


def as_text(v) -> str:
    """
    엑셀에서 읽힌 숫자형 코드가 '50000167.0' 처럼 되는 것을 막는다.
    카테고리코드/원산지코드가 이렇게 들어가면 등록이 전부 반려된다.
    """
    s = str(v if v is not None else "").strip()
    return s[:-2] if re.fullmatch(r"\d+\.0", s) else s


def load_template(path: str):
    """템플릿에서 1행(그룹), 2행(컬럼명)을 원본 그대로 보존해 반환."""
    if path.endswith((".xlsx", ".xls")):
        raw = pd.read_excel(path, header=None, dtype=str).fillna("")
        rows = raw.values.tolist()
    else:
        for enc in ("utf-8-sig", "cp949"):
            try:
                with open(path, encoding=enc, newline="") as f:
                    rows = list(csv.reader(f))
                break
            except UnicodeDecodeError:
                continue
        else:
            sys.exit("템플릿 인코딩을 읽을 수 없습니다. UTF-8 CSV 또는 xlsx로 저장해 주세요.")

    if len(rows) < 2:
        sys.exit("템플릿 형식이 올바르지 않습니다.")

    group_row, header_row = rows[0], rows[1]
    # 조회용 정규화 키 (줄바꿈/공백 제거)
    keys = [re.sub(r"\s+", "", c) for c in header_row]
    return group_row, header_row, keys


def main():
    p = argparse.ArgumentParser()
    p.add_argument("infile", help="enrich_ai.py 출력 엑셀 (final.xlsx)")
    p.add_argument("--template", required=True, help="네이버 일괄등록 템플릿 CSV/XLSX")
    p.add_argument("--out", default="upload.csv")
    p.add_argument("--origin-code", default="0200038",
                   help="일본 원산지코드. '원산지 찾기' 팝업에서 확인 후 교체 필수")
    p.add_argument("--importer", default="", help="수입사(구매대행 사업자 상호)")
    p.add_argument("--ship-template", default="", help="배송비 템플릿코드(출고지=일본)")
    p.add_argument("--as-phone", default="", help="A/S 전화번호")
    p.add_argument("--as-guide", default="상품 수령 후 7일 이내 고객센터로 문의해 주세요.")
    p.add_argument("--stock", type=int, default=10, help="기본 재고수량")
    p.add_argument("--usd-krw", type=float, default=1380.0, help="관부가세 판정용 환율")
    p.add_argument("--image-base", default="",
                   help="자체 호스팅 이미지 URL 접두사 (미지정 시 원본 URL 사용)")
    a = p.parse_args()

    df = pd.read_excel(a.infile, dtype=str).fillna("")
    group_row, header_row, keys = load_template(a.template)
    ncol = len(header_row)
    print(f"템플릿 컬럼 {ncol}개 / 입력 상품 {len(df)}건")

    def idx(name: str):
        k = re.sub(r"\s+", "", name)
        return keys.index(k) if k in keys else None

    out_rows = []
    skipped = []

    for i, row in enumerate(df.to_dict("records")):
        r = [""] * ncol

        def put(col, val):
            j = idx(col)
            if j is not None:
                r[j] = "" if val is None else str(val)

        name = clean_name(row.get("판매상품명") or row.get("상품명_원문"))
        price = round10(row.get("판매가_원"))
        cat = as_text(row.get("카테고리ID"))
        imgs = [u.strip() for u in str(row.get("이미지목록", "")).split("\n") if u.strip()]

        # 필수값 검증
        missing = []
        if not name:
            missing.append("상품명")
        if not price:
            missing.append("판매가")
        if not cat:
            missing.append("카테고리코드")
        if not imgs:
            missing.append("대표이미지")
        if missing:
            skipped.append((i + 1, name[:30], ",".join(missing)))
            continue

        if a.image_base:
            base = a.image_base.rstrip("/")
            imgs = [f"{base}/{u.split('/')[-1].split('?')[0]}" for u in imgs]

        # 상품 기본정보
        put("판매자 상품코드", clean_code(as_text(row.get("상품코드"))))
        put("카테고리코드", cat)
        put("상품명", name)
        put("상품상태", "신상품")
        put("판매가", price)
        put("부가세", "과세상품")
        put("관부가세", judge_duty(row.get("총원가_원"), a.usd_krw))
        put("재고수량", a.stock)
        put("대표이미지", imgs[0])
        put("추가이미지", "\n".join(imgs[1:10]))   # 대표 제외 최대 9장
        put("상세설명", build_detail(row, a.image_base))

        # 상품 주요정보
        put("브랜드", clean_name(row.get("브랜드"), 50))
        put("원산지코드", as_text(a.origin_code))
        put("수입사", a.importer)
        put("복수원산지여부", "N")
        put("미성년자 구매", "Y")

        # 배송정보 (템플릿코드가 있으면 나머지는 무시됨)
        if a.ship_template:
            put("배송비 템플릿코드", a.ship_template)
        else:
            put("배송방법", "택배, 소포, 등기")
            put("배송비유형", "무료")

        # A/S
        if a.as_phone:
            put("A/S 전화번호", a.as_phone)
        put("A/S 안내", a.as_guide)

        out_rows.append(r)

    # 1행 그룹 + 2행 컬럼명 + 데이터 (작성가이드 3~6행은 제외)
    with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(group_row)
        w.writerow(header_row)
        w.writerows(out_rows)

    print(f"\n저장 완료: {a.out}  (등록 {len(out_rows)}건)")
    if skipped:
        print(f"\n제외 {len(skipped)}건 (필수값 누락):")
        for n, nm, why in skipped[:20]:
            print(f"  {n}행 {nm} -> {why}")

    if a.origin_code == "0200038":
        print("\n[확인] --origin-code 가 기본값입니다. "
              "'원산지 찾기' 팝업에서 일본 코드를 확인해 교체하세요.")
    if not a.ship_template:
        print("[확인] 배송비 템플릿코드가 없습니다. "
              "출고지가 국내로 잡히면 관부가세 값이 무시됩니다.")


if __name__ == "__main__":
    main()
