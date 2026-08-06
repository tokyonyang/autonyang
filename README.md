# 라쿠텐 구매대행 자동화

로컬 개발 환경 없이 GitHub 웹 UI만으로 운영합니다.
수집 → 번역 → 스마트스토어 양식 변환 → 가격·재고 감시까지 자동화합니다.

---

## 1. 리포 만들기

GitHub에서 **New repository** → 이름 `rakuten-auto` → **Private**로 생성합니다.
(API 키를 다루므로 반드시 Private)

리포 화면에서 **Add file → Create new file**로 아래 파일들을 하나씩 만듭니다.
파일명 칸에 `.github/workflows/collect.yml` 처럼 슬래시를 넣으면 폴더가 자동으로 생깁니다.

```
rakuten-auto/
├── .github/workflows/
│   ├── collect.yml          수집 (버튼 실행)
│   ├── monitor.yml          감시 (매일 자동)
│   └── import.yml           DB 적재 (버튼 실행)
├── rakuten_collector.py
├── enrich_ai.py
├── to_smartstore.py
├── monitor.py
├── schema.sql
├── requirements.txt
├── template.csv             ★ 네이버 일괄등록 템플릿 (직접 업로드)
├── naver_categories.csv     ★ 카테고리 목록 (직접 업로드)
└── urls.txt                 (선택) URL 개별 수집용
```

★ 표시된 두 파일은 **Add file → Upload files**로 올리세요.
`template.csv`는 반드시 **CSV UTF-8**로 저장한 것을 올려야 합니다.

---

## 2. 키 등록

리포 → **Settings → Secrets and variables → Actions**

### Secrets 탭 (New repository secret)

| 이름 | 발급처 |
|---|---|
| `RAKUTEN_APP_ID` | https://webservice.rakuten.co.jp/app/create |
| `RAKUTEN_ACCESS_KEY` | 같은 화면 (신규 필수) |
| `OPENAI_API_KEY` | platform.openai.com |
| `SUPABASE_URL` | Supabase 프로젝트 설정 |
| `SUPABASE_KEY` | Supabase `service_role` 키 |
| `TELEGRAM_BOT_TOKEN` | @BotFather (선택) |
| `TELEGRAM_CHAT_ID` | (선택) |

### Variables 탭 (New repository variable)

민감하지 않은 값이라 Variables에 둡니다.

| 이름 | 예시 | 설명 |
|---|---|---|
| `ORIGIN_CODE` | `0200038` | 일본 원산지코드. **'원산지 찾기' 팝업에서 확인 후 교체 필수** |
| `IMPORTER` | `디깅스팟` | 수입사(사업자 상호) |
| `SHIP_TEMPLATE` | `12345` | 배송비 템플릿코드 (출고지=일본) |
| `AS_PHONE` | `0507-0000-0000` | A/S 전화번호 |

---

## 3. Supabase 준비

1. supabase.com에서 프로젝트 생성 (리전은 **Northeast Asia (Seoul)** 권장)
2. 좌측 **SQL Editor** → `schema.sql` 내용 붙여넣고 **Run**
3. **Settings → API**에서 Project URL과 `service_role` 키를 복사해 Secrets에 등록

---

## 4. 사용하기

### 상품 수집

**Actions 탭 → 상품 수집 → Run workflow**

입력값을 채우고 실행합니다.

| 항목 | 설명 |
|---|---|
| mode | `search`(키워드) / `shop`(샵 통째로) / `urls`(URL 목록) |
| query | 키워드 또는 샵코드 |
| max | 최대 수집 개수 |
| images | 고화질 이미지 수집 여부 (켜면 느림) |
| translate | AI 번역 + 카테고리 매핑 |

**처음엔 max를 10으로 두고 한 번 돌려보세요.** 결과가 마음에 들면 늘리면 됩니다.

실행이 끝나면 해당 run 화면 하단 **Artifacts**에서 `result-N.zip`을 받습니다.
안에 `upload.csv`가 스마트스토어에 바로 올릴 파일입니다.

### 스마트스토어 등록

스마트스토어센터 → 상품관리 → **상품 일괄등록** → `upload.csv` 업로드

### DB 적재

등록을 마쳤으면 `final.xlsx`를 리포에 업로드한 뒤
**Actions → 등록상품 DB 적재 → Run workflow**

이걸 해야 감시 대상이 됩니다.

### 감시

매일 한국시간 새벽 3시에 자동 실행됩니다.
결과 요약은 텔레그램으로 오고, 조치 목록은 Artifacts의 `actions.csv`에 있습니다.

---

## 알아둘 것

**스케줄이 밀릴 수 있습니다.** GitHub의 무료 러너는 부하가 몰리면 cron 실행을 수십 분 늦추거나 건너뜁니다. 정확한 시각이 중요하면 Oracle Cloud 인스턴스에서 cron으로 돌리는 편이 안정적입니다.

```bash
0 3 * * * cd ~/rakuten && python3 monitor.py check --limit 500 >> monitor.log 2>&1
```

**감시 소요 시간.** 라쿠텐 API를 1초 간격으로 부르므로 300건이면 약 6분, 1,000건이면 20분쯤 걸립니다. Actions 무료 한도(Private 리포 월 2,000분)를 감안해 `limit`을 조절하세요.

**장르 캐시.** `genre_cache.json`이 실행 후 자동으로 리포에 커밋됩니다. 한 번 매핑한 장르는 다시 LLM을 부르지 않아서, 며칠 쓰면 번역 비용이 크게 줄어듭니다.

**리포는 반드시 Private.** 수집한 상품 데이터와 캐시가 그대로 올라갑니다.
