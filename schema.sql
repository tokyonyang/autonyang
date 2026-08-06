-- ============================================================
-- 라쿠텐 구매대행 상품 감시 스키마
-- Supabase SQL Editor에 붙여넣고 실행하세요.
-- ============================================================

-- 등록 상품 마스터
create table if not exists rakuten_products (
  item_code        text primary key,          -- 라쿠텐 'shop:1234'
  shop_code        text,
  item_url         text,
  item_name_ja     text,                      -- 원문 상품명
  sale_name        text,                      -- 한국 판매상품명
  seller_code      text,                      -- 스마트스토어 판매자 상품코드
  category_id      text,

  base_price_jpy   integer,                   -- 등록 시점 엔화 원가
  base_fx           numeric(10,4),            -- 등록 시점 엔/원 환율
  listed_price_krw integer,                   -- 현재 스마트스토어 판매가
  target_margin    numeric(5,4) default 0.30, -- 목표 마진율

  current_price_jpy integer,
  availability     smallint default 1,        -- 1 판매중 / 0 품절
  status           text default 'active',     -- active / soldout / delisted / paused
  consecutive_miss smallint default 0,        -- API 미조회 연속 횟수

  first_seen       timestamptz default now(),
  last_checked     timestamptz,
  updated_at       timestamptz default now()
);

create index if not exists idx_products_check
  on rakuten_products (last_checked nulls first)
  where status in ('active', 'soldout');

create index if not exists idx_products_shop
  on rakuten_products (shop_code);


-- 가격/재고 변동 이력
create table if not exists price_history (
  id          bigserial primary key,
  item_code   text not null references rakuten_products(item_code) on delete cascade,
  price_jpy   integer,
  availability smallint,
  fx_jpy_krw  numeric(10,4),
  cost_krw    integer,          -- 환율 반영 총원가
  checked_at  timestamptz default now()
);

create index if not exists idx_history_item
  on price_history (item_code, checked_at desc);


-- 조치가 필요한 알림
create table if not exists alerts (
  id          bigserial primary key,
  item_code   text not null references rakuten_products(item_code) on delete cascade,
  kind        text not null,    -- price_up / price_down / soldout / restock / delisted / margin_low
  severity    text default 'info',   -- info / warn / critical
  old_value   text,
  new_value   text,
  message     text,
  resolved    boolean default false,
  created_at  timestamptz default now()
);

create index if not exists idx_alerts_open
  on alerts (resolved, created_at desc);


-- 최신 상태 한눈에 보기
create or replace view v_product_status as
select
  p.item_code,
  p.sale_name,
  p.seller_code,
  p.status,
  p.base_price_jpy,
  p.current_price_jpy,
  case when p.base_price_jpy > 0
       then round((p.current_price_jpy - p.base_price_jpy)::numeric
                  / p.base_price_jpy * 100, 1)
       else null end                              as price_change_pct,
  p.listed_price_krw,
  p.last_checked,
  (select count(*) from alerts a
    where a.item_code = p.item_code and not a.resolved) as open_alerts
from rakuten_products p
order by open_alerts desc, p.last_checked nulls first;
