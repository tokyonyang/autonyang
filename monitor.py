#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
라쿠텐 등록상품 가격·재고 감시기
--------------------------------------------------
등록해둔 상품의 현지 가격과 재고를 주기적으로 다시 조회해서
역마진·품절·판매종료를 잡아냅니다. 엔화 가격뿐 아니라 환율 변동도 함께 봅니다.

준비
  1) schema.sql 을 Supabase SQL Editor에서 실행
  2) pip install requests pandas openpyxl
  3) 환경변수
     export RAKUTEN_APP_ID=... RAKUTEN_ACCESS_KEY=...
     export SUPABASE_URL="https://xxxx.supabase.co"
     export SUPABASE_KEY="service_role 키"
     export TELEGRAM_BOT_TOKEN=...   # 선택
     export TELEGRAM_CHAT_ID=...     # 선택

사용
  # 등록한 상품을 DB에 적재 (to_smartstore.py 직전 파일 사용)
  python monitor.py import final.xlsx

  # 감시 실행 (오래된 것부터 200건)
  python monitor.py check --limit 200

  # 매일 새벽 3시 cron 예시
  # 0 3 * * * cd /home/ubuntu/rakuten && /usr/bin/python3 monitor.py check --limit 500 >> monitor.log 2>&1
"""

import argparse
import os
import sys
import time
from datetime import datetime, timezone

import requests

# 같은 폴더의 수집기에서 가격 계산 로직 재사용
try:
    from rakuten_collector import PriceConfig, calc_sale_price, api_call, _session
except ImportError:
    sys.exit("rakuten_collector.py 를 같은 폴더에 두세요.")

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "")
TG_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TG_CHAT = os.environ.get("TELEGRAM_CHAT_ID", "")

API_DELAY = 1.0


# ----------------------------------------------------------------------
# Supabase REST
# ----------------------------------------------------------------------
def sb_headers(extra=None):
    if not SUPABASE_URL or not SUPABASE_KEY:
        sys.exit("SUPABASE_URL / SUPABASE_KEY 환경변수를 설정하세요.")
    h = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
    }
    if extra:
        h.update(extra)
    return h


def sb_get(table, params):
    r = requests.get(f"{SUPABASE_URL}/rest/v1/{table}",
                     headers=sb_headers(), params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def sb_upsert(table, rows, on_conflict=None):
    if not rows:
        return
    params = {"on_conflict": on_conflict} if on_conflict else {}
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/{table}",
        headers=sb_headers({"Prefer": "resolution=merge-duplicates,return=minimal"}),
        params=params, json=rows, timeout=60)
    if r.status_code >= 300:
        raise RuntimeError(f"upsert {table} 실패 {r.status_code}: {r.text[:300]}")


def sb_patch(table, match: dict, values: dict):
    params = {k: f"eq.{v}" for k, v in match.items()}
    r = requests.patch(f"{SUPABASE_URL}/rest/v1/{table}",
                       headers=sb_headers({"Prefer": "return=minimal"}),
                       params=params, json=values, timeout=30)
    if r.status_code >= 300:
        raise RuntimeError(f"patch {table} 실패 {r.status_code}: {r.text[:300]}")


# ----------------------------------------------------------------------
# 환율
# ----------------------------------------------------------------------
def fetch_fx(fallback: float) -> float:
    """엔/원 환율. 실패하면 fallback 사용."""
    try:
        r = requests.get("https://open.er-api.com/v6/latest/JPY", timeout=10)
        v = r.json().get("rates", {}).get("KRW")
        if v and 5 < float(v) < 20:
            return round(float(v), 4)
    except Exception:
        pass
    print(f"  환율 조회 실패 -> 지정값 {fallback} 사용", file=sys.stderr)
    return fallback


# ----------------------------------------------------------------------
# 알림
# ----------------------------------------------------------------------
def send_telegram(text: str):
    if not TG_TOKEN or not TG_CHAT:
        return
    try:
        requests.post(f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage",
                      json={"chat_id": TG_CHAT, "text": text,
                            "parse_mode": "HTML",
                            "disable_web_page_preview": True},
                      timeout=15)
    except Exception as e:
        print(f"  텔레그램 전송 실패: {e}", file=sys.stderr)


# ----------------------------------------------------------------------
# import : 등록 상품 적재
# ----------------------------------------------------------------------
def cmd_import(a):
    import pandas as pd
    df = pd.read_excel(a.infile, dtype=str).fillna("")
    fx = fetch_fx(a.rate)
    rows = []
    for r in df.to_dict("records"):
        code = str(r.get("상품코드", "")).strip()
        if not code:
            continue

        def num(key):
            try:
                return int(float(str(r.get(key, "") or 0)))
            except ValueError:
                return 0

        rows.append({
            "item_code": code,
            "shop_code": str(r.get("샵코드", "")).strip(),
            "item_url": str(r.get("상품URL", "")).strip(),
            "item_name_ja": str(r.get("상품명_원문", ""))[:500],
            "sale_name": str(r.get("판매상품명", ""))[:200],
            "seller_code": code.replace(":", "_")[:30],
            "category_id": str(r.get("카테고리ID", "")).strip().removesuffix(".0"),
            "base_price_jpy": num("가격_엔"),
            "base_fx": fx,
            "listed_price_krw": num("판매가_원"),
            "target_margin": a.margin,
            "current_price_jpy": num("가격_엔"),
            "availability": 1,
            "status": "active",
            "last_checked": datetime.now(timezone.utc).isoformat(),
        })

    for i in range(0, len(rows), 200):
        sb_upsert("rakuten_products", rows[i:i + 200], on_conflict="item_code")
        print(f"  적재 {min(i+200, len(rows))}/{len(rows)}")
    print(f"\n적재 완료 {len(rows)}건 (기준 환율 {fx})")


# ----------------------------------------------------------------------
# check : 감시
# ----------------------------------------------------------------------
def cmd_check(a):
    fx = fetch_fx(a.rate)
    print(f"기준 환율 {fx} 엔/원")

    items = sb_get("rakuten_products", {
        "select": "*",
        "status": "in.(active,soldout)",
        "order": "last_checked.asc.nullsfirst",
        "limit": a.limit,
    })
    if not items:
        print("감시 대상이 없습니다.")
        return
    print(f"감시 대상 {len(items)}건\n")

    session = _session()
    history, alerts, updates, sheet = [], [], [], []
    now = datetime.now(timezone.utc).isoformat()

    for n, it in enumerate(items, 1):
        code = it["item_code"]
        cfg = PriceConfig(jpy_krw=fx,
                          margin_rate=float(it.get("target_margin") or 0.30))

        try:
            data = api_call(session, itemCode=code)
            found = (data.get("items") or [None])[0]
        except Exception as e:
            print(f"  [{n}] {code} 조회 실패: {e}", file=sys.stderr)
            time.sleep(API_DELAY)
            continue
        time.sleep(API_DELAY)

        # --- 상품이 사라진 경우 ---
        if not found:
            miss = int(it.get("consecutive_miss") or 0) + 1
            upd = {"item_code": code, "consecutive_miss": miss,
                   "last_checked": now, "updated_at": now}
            # 일시적 누락일 수 있어 3회 연속일 때만 판매종료 처리
            if miss >= 3:
                upd["status"] = "delisted"
                alerts.append({
                    "item_code": code, "kind": "delisted", "severity": "critical",
                    "message": f"{it.get('sale_name','')} 상품이 라쿠텐에서 사라졌습니다. 판매중지 필요.",
                })
                sheet.append((it, "판매중지", "판매종료", "", ""))
            updates.append(upd)
            continue

        price = int(found.get("itemPrice") or 0)
        avail = int(found.get("availability") or 0)
        calc = calc_sale_price(price, cfg)
        cost = calc["총원가_원"]

        history.append({"item_code": code, "price_jpy": price,
                        "availability": avail, "fx_jpy_krw": fx,
                        "cost_krw": cost})

        old_price = int(it.get("current_price_jpy") or it.get("base_price_jpy") or 0)
        old_avail = int(it.get("availability") or 1)
        listed = int(it.get("listed_price_krw") or 0)
        name = it.get("sale_name") or it.get("item_name_ja") or code

        # --- 재고 변동 ---
        if old_avail == 1 and avail == 0:
            alerts.append({"item_code": code, "kind": "soldout", "severity": "critical",
                           "message": f"{name} 품절. 판매중지 필요."})
            sheet.append((it, "판매중지", "현지 품절", "", ""))
        elif old_avail == 0 and avail == 1:
            alerts.append({"item_code": code, "kind": "restock", "severity": "info",
                           "message": f"{name} 재입고. 판매재개 가능."})

        # --- 실질 마진 (현재 판매가 기준) ---
        margin_now = listed * (1 - cfg.market_fee) - cost
        margin_pct = margin_now / listed if listed else 0

        if listed and margin_pct < a.min_margin:
            sev = "critical" if margin_pct <= 0 else "warn"
            alerts.append({
                "item_code": code, "kind": "margin_low", "severity": sev,
                "old_value": str(listed), "new_value": str(calc["판매가_원"]),
                "message": (f"{name} 마진 {margin_pct*100:.1f}% "
                            f"(원가 {cost:,}원 / 판매가 {listed:,}원) "
                            f"-> 권장 {calc['판매가_원']:,}원"),
            })
            sheet.append((it, "가격인상", f"마진 {margin_pct*100:.1f}%",
                          listed, calc["판매가_원"]))

        # --- 엔화 가격 변동 ---
        if old_price and price != old_price:
            pct = (price - old_price) / old_price * 100
            if abs(pct) >= a.price_threshold:
                kind = "price_up" if pct > 0 else "price_down"
                alerts.append({
                    "item_code": code, "kind": kind,
                    "severity": "warn" if pct > 0 else "info",
                    "old_value": str(old_price), "new_value": str(price),
                    "message": f"{name} 현지가 {old_price:,}엔 -> {price:,}엔 ({pct:+.1f}%)",
                })
                # 가격이 내렸으면 판매가도 낮춰 경쟁력 확보
                if pct < 0 and listed and calc["판매가_원"] < listed:
                    sheet.append((it, "가격인하", f"현지가 {pct:+.1f}%",
                                  listed, calc["판매가_원"]))

        updates.append({
            "item_code": code, "current_price_jpy": price, "availability": avail,
            "status": "soldout" if avail == 0 else "active",
            "consecutive_miss": 0, "last_checked": now, "updated_at": now,
        })

        if n % 25 == 0:
            print(f"  {n}/{len(items)}")

    # --- 저장 ---
    for i in range(0, len(history), 200):
        sb_upsert("price_history", history[i:i + 200])
    for i in range(0, len(updates), 200):
        sb_upsert("rakuten_products", updates[i:i + 200], on_conflict="item_code")
    for i in range(0, len(alerts), 200):
        sb_upsert("alerts", alerts[i:i + 200])

    # --- 조치 시트 ---
    if sheet:
        import csv as _csv
        with open(a.out, "w", encoding="utf-8-sig", newline="") as f:
            w = _csv.writer(f)
            w.writerow(["판매자상품코드", "상품명", "조치", "사유",
                        "현재판매가", "권장판매가", "상품URL"])
            for it, act, why, cur, new in sheet:
                w.writerow([it.get("seller_code", ""), it.get("sale_name", ""),
                            act, why, cur, new, it.get("item_url", "")])
        print(f"\n조치 필요 목록: {a.out} ({len(sheet)}건)")

    # --- 요약 ---
    def cnt(k):
        return sum(1 for x in alerts if x["kind"] == k)

    crit = sum(1 for x in alerts if x["severity"] == "critical")
    summary = (
        f"<b>라쿠텐 감시 리포트</b>\n"
        f"점검 {len(items)}건 · 환율 {fx}\n\n"
        f"긴급 {crit}건\n"
        f"품절 {cnt('soldout')} / 판매종료 {cnt('delisted')} / 재입고 {cnt('restock')}\n"
        f"마진경고 {cnt('margin_low')}\n"
        f"현지가 상승 {cnt('price_up')} / 하락 {cnt('price_down')}"
    )
    print("\n" + summary.replace("<b>", "").replace("</b>", ""))

    top = [x for x in alerts if x["severity"] == "critical"][:10]
    if top:
        summary += "\n\n<b>즉시 조치</b>\n" + "\n".join(
            f"· {x['message'][:90]}" for x in top)
    send_telegram(summary)


# ----------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser(description="라쿠텐 등록상품 감시기")
    sub = p.add_subparsers(dest="cmd", required=True)

    s1 = sub.add_parser("import", help="등록 상품을 DB에 적재")
    s1.add_argument("infile", help="enrich_ai.py 출력 엑셀")
    s1.add_argument("--rate", type=float, default=9.5)
    s1.add_argument("--margin", type=float, default=0.30)

    s2 = sub.add_parser("check", help="가격·재고 감시")
    s2.add_argument("--limit", type=int, default=200, help="회당 점검 건수")
    s2.add_argument("--rate", type=float, default=9.5, help="환율 조회 실패 시 사용")
    s2.add_argument("--min-margin", type=float, default=0.10,
                    help="이 마진율 아래로 떨어지면 경고")
    s2.add_argument("--price-threshold", type=float, default=3.0,
                    help="현지가 변동 알림 기준 %%")
    s2.add_argument("--out", default="actions.csv")

    a = p.parse_args()
    (cmd_import if a.cmd == "import" else cmd_check)(a)


if __name__ == "__main__":
    main()
