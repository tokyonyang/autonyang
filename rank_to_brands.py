#!/usr/bin/env python3
"""소싱 랭킹(soursing_item) 상위 브랜드 → brands.txt 별칭과 매칭 → brands_active.txt 생성.

- brands.txt 에서 '주석 처리되지 않은' 브랜드만 후보로 씁니다.
  (C등급처럼 주석 처리한 브랜드는 랭킹 상위여도 자동으로 켜지지 않습니다.)
- 매칭 안 된 랭킹 브랜드는 리포트로 출력합니다. 수집하고 싶으면 brands.txt 에
  별칭(한글|영문|가타카나)을 직접 추가하세요.

사용:
  python rank_to_brands.py                      # 기본: soursing_item 공개 CSV, 상위 30개
  python rank_to_brands.py --source ranking.csv --top 15 --out brands_active.txt
"""
import argparse
import csv
import io
import sys
import unicodedata
import urllib.request

DEFAULT_SOURCE = ("https://raw.githubusercontent.com/tokyonyang/soursing_item/"
                  "main/dashboard_data/latest_product_keywords.csv")


def norm(s):
    return unicodedata.normalize("NFKC", s or "").strip().lower().replace(" ", "")


def read_text(src):
    if src.startswith(("http://", "https://")):
        with urllib.request.urlopen(src, timeout=30) as r:
            return r.read().decode("utf-8-sig")
    with open(src, encoding="utf-8-sig") as f:
        return f.read()


def load_ranking(src, recs):
    """브랜드별 최고 final_score. recs: 허용할 recommendation 값."""
    rows = csv.DictReader(io.StringIO(read_text(src), newline=""))
    best = {}
    for r in rows:
        brand = (r.get("brand") or "").strip()
        if not brand or r.get("recommendation") not in recs:
            continue
        try:
            score = float(r.get("final_score") or 0)
        except ValueError:
            continue
        if brand not in best or score > best[brand]:
            best[brand] = score
    return sorted(best.items(), key=lambda kv: -kv[1])


def load_brand_lines(path):
    """(원본줄, [정규화 별칭], 활성여부). 주석 처리된 브랜드 줄('# 이름|...')도 비활성으로 읽는다."""
    out = []
    with open(path, encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            active = True
            if line.startswith("#"):
                line = line.lstrip("#").strip()
                active = False
            if "|" not in line:
                continue  # 설명 주석 등
            aliases = [norm(a) for a in line.split("|") if a.strip()]
            if aliases:
                out.append((line, aliases, active))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=DEFAULT_SOURCE, help="랭킹 CSV 경로 또는 URL")
    ap.add_argument("--brands", default="brands.txt")
    ap.add_argument("--out", default="brands_active.txt")
    ap.add_argument("--top", type=int, default=30, help="랭킹 상위 N개 브랜드까지만 검토")
    ap.add_argument("--rec", default="테스트,관찰", help="포함할 추천 등급(쉼표). 기본 테스트,관찰")
    args = ap.parse_args()

    ranking = load_ranking(args.source, set(args.rec.split(",")))[: args.top]
    lines = load_brand_lines(args.brands)

    matched, disabled, unmatched, seen = [], [], [], set()
    for brand, score in ranking:
        key = norm(brand)
        hit = next((l for l in lines if key in l[1]), None)
        if hit is None:
            unmatched.append((brand, score))
        elif hit[2]:
            if hit[0] not in seen:
                seen.add(hit[0])
                matched.append((brand, score, hit[0]))
        else:
            disabled.append((brand, score))

    print(f"랭킹 상위 {len(ranking)}개 브랜드 검토 (기준: {args.rec})")
    print(f"\n[수집 대상] {len(matched)}개")
    for b, s, line in matched:
        print(f"  {s:5.1f}  {b}  ←  {line}")
    print(f"\n[brands.txt에서 주석 처리됨 → 자동으로 켜지 않음] {len(disabled)}개")
    for b, s in disabled:
        print(f"  {s:5.1f}  {b}")
    print(f"\n[brands.txt에 없음 → 별칭 추가 필요] {len(unmatched)}개")
    for b, s in unmatched:
        print(f"  {s:5.1f}  {b}")

    if not matched:
        print("\n매칭된 브랜드가 없어 brands_active.txt 를 만들지 않습니다.", file=sys.stderr)
        sys.exit(2)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write("# rank_to_brands.py 자동 생성 — 직접 수정하지 마세요\n")
        for _, _, line in matched:
            f.write(line + "\n")
    print(f"\n→ {args.out} 저장")


if __name__ == "__main__":
    main()
