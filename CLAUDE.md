# CLAUDE.md — Project Context for Future Claude Sessions

이 문서는 다음 Claude 세션이 이 프로젝트의 맥락을 빠르게 이해할 수 있도록 작성됨. 사용자(주로 한국어)와 협업하며 기능 추가/수정 요청을 받게 됨.

## Project Summary

**Meta Ads Reporting Dashboard** — Flask 기반 단일 페이지 웹 앱.
British Council Korea Meta(Facebook/Instagram) 광고 운영자가 주간 전략 리포트를 빠르게 생성하기 위한 도구.

- 백엔드: `app.py` (Flask, Python)
- 프론트: `templates/index.html` (Vanilla JS + CSS, 단일 페이지)
- 외부 의존성: Meta Marketing API (Graph API v25.0)
- 인증: User Access Token (장기 60일) → `.env`의 `META_ACCESS_TOKEN`
- 광고 계정은 `accounts.json`에 영속, UI에서 추가/삭제 가능

## File Structure

```
JAY/
├── app.py                  # Flask backend, all endpoints, Excel generation
├── templates/index.html    # SPA: all UI, CSS, JS in one file
├── requirements.txt        # Flask, requests, python-dotenv, openpyxl
├── .env                    # META_ACCESS_TOKEN, META_AD_ACCOUNT_IDS, META_API_VERSION
├── .env.example            # Template (no real values)
├── accounts.json           # Persisted account allowlist (auto-managed)
├── .gitignore
├── README.md               # User-facing setup guide (Korean, non-programmer)
└── CLAUDE.md               # This file
```

## Architecture & Key Conventions

### Data Flow
1. Frontend `/조회` → backend `/api/insights?account_id=...&since=...&until=...&level=ad`
2. Backend `fetch_insights()` calls Meta `/{act_id}/insights?level=ad&time_increment=1` with pagination
3. Each row is enriched server-side via `_enrich_row()`:
   - `video_views` from `actions[action_type="video_view"]` (null if absent → image creative)
   - `vtr` = video_views / impressions × 100 (null if no video)
   - `conversions` from action types in `_CONVERSION_ACTION_TYPES` priority order
   - `cvr` = conversions / impressions × 100
4. Frontend annotates rows via `annotateRows()` → adds `_target`, `_targeting`, `_targeting_main`, `_creative` (BC naming heuristics)
5. All cards re-aggregate from the same row set client-side; filter changes trigger client re-render without re-fetch

### Critical Design Decision: `level=ad + time_increment=1` always
The UI's "리포트 수준" dropdown was REMOVED (hidden `<select id="level">` with value "ad"). All fetches use ad-level daily data. Higher-level views (campaign/adset) aggregate in JS. This unblocks the daily breakdown which needs date_start per row.

### BC Naming Heuristics (hard-coded in both JS and Python — must stay in sync)

**Target** parsed from `campaign_name`:
- `kinder` or `parents` → "Kinder"
- `adults` → "Adults"
- `younglearners` / `youngleaners` → "Young Learners"
- else → "기타"

**Targeting** parsed from `adset_name + campaign_name`:
- Priority: A+ > LAL > PT > Ilsan
- Modifiers appended: Primary / summer / winter / BAU
- Examples: "A+", "LAL - Primary", "A+ - summer"

**Creative** parsed from `ad_name` + the targeting main tag:
- Find main tag (A+/LAL/PT/Ilsan) in ad_name, take substring after, strip leading `-_`
- Example: `english-ea-ad-kr-conversion-A+-video-testimonial-jeon` → `video-testimonial-jeon`
- Tooltip shows full ad_name

JS implementations: `parseTarget`, `parseTargeting`, `parseTargetingMain`, `parseCreative` (in index.html, search for these names).
Python implementations: `_parse_target`, `_parse_targeting`, `_parse_targeting_main`, `_parse_creative` (app.py, same logic).

### LocalStorage Keys
- `metaAccountId` — currently selected account
- `comparisonKpi` — selected main KPI for insights ("impressions" / "clicks" / "video_views" / "conversions")
- `progressOverrides` — per-account, per-campaign progress period overrides for the Target × Asset card

### State Objects (frontend)
- `summaryState`: `{allRows, level, filters}` — drives SUMMARY + daily + detail
- `tableState`: `{allRows, level, filters}` — drives detail table (per-column filters)
- `summaryExpanded`: Set of `target|||targeting` keys for SUMMARY drilldown
- `comparisonState`: `{mode, data}` for insight comparison card
- `comparisonExpanded`: `{campaigns: Set, adsets: Set}` for comparison drilldown
- `dailyState`: `{context, subContext}` — Target/Targeting context for daily wide-format table
- `insightKpi`: string, the chosen main KPI

## Features Built (Chronological)

1. **Initial scaffold** — Flask + index.html with date range, level dropdown, basic table + CSV
2. **Weekly insights card** — original "이번 주 vs 지난 주 같은 길이" comparison
3. **Multi-account management** — `accounts.json` storage, POST/DELETE endpoints, UI for add/remove
4. **Target × Asset breakdown card** — campaign hierarchy + 진행률(time%) / 소진률(spend%) progress bars + manual period override
5. **Excel template export (deprecated)** — BC's 25MB template with `_xlfn.LET` formulas. openpyxl can't preserve dynamic-array functions → all formulas became `#NAME?`. Endpoints kept (`/api/report.xlsx/*`) but UI button removed.
6. **VIEW/VTR + hatched empty cells** for image creatives
7. **SUMMARY card** — Target × Targeting grouped table with Sub Totals
8. **Multi-select filters** — Target/Targeting/Campaign/AdSet/Ad with checkbox popups
9. **Daily card (wide format)** — Date × Target/Targeting columns with context selector
10. **Comparison card refactor** — replaced single weekly with tabbed (주간/월간) + custom date ranges + per-campaign breakdown
11. **Comparison drilldown** — campaign → adset → ad expand
12. **KPI selector lens** — 4 buttons (노출/클릭/뷰/전환) drive auto-insight bullets + tile highlight + table sort
13. **Creative-level drilldown** — SUMMARY Targeting rows expand to show Creative rows + per-Targeting Sub Total
14. **Daily sub-context** — Target selection reveals Targeting buttons; picking one shows Date × [Targeting Total + Creatives]
15. **Excel rewrite** — 5 sheets (Weekly Insights / SUMMARY / 일자별_[Target] × 3) with try/except+traceback

## Endpoints

| Method | Path | Notes |
|---|---|---|
| GET | `/` | Renders SPA |
| GET | `/api/accounts` | List with Meta API names cached in-memory |
| POST | `/api/accounts` | Add (validates against Meta API) |
| DELETE | `/api/accounts/<id>` | Remove from allowlist |
| GET | `/api/insights` | Ad-level daily rows (level=ad + time_increment=1) |
| GET | `/api/insights.csv` | Same as above but CSV |
| GET | `/api/breakdown` | Campaign hierarchy + cumulative spend + progress |
| GET | `/api/comparison` | Two periods of raw ad rows for insight card |
| GET | `/api/view.xlsx` | 5-sheet Excel report. Wrapped in try/except + traceback in response |
| GET | `/api/weekly-insights` | (legacy, unused but kept for backward compat) |
| POST/GET | `/api/report.xlsx/*` | (deprecated BC template flow, retained for URL stability) |

## Filter System

Two filter layers, AND-combined:
1. **SUMMARY filter bar** (`summaryState.filters`) — top of SUMMARY card, drives EVERYTHING
2. **Detail table column header filters** (`tableState.filters`) — additional narrowing within detail table only

When SUMMARY filters change, detail table filters reset and `tableState.allRows` becomes the filtered set.

## Deferred / Non-Goals

Things we discussed but explicitly deferred:
- **LLM-based natural language analysis** — User chose option C (KPI selector) over option A (LLM) for custom analysis requests
- **Planned budget / CPM / CPC / CPV / CPA columns** — requires user-managed target values (no UI yet)
- **Achievement / Pacing / E.Conversion** — derived from Planned (above)
- **Email / Slack automation** — out of scope
- **PDF export** — out of scope
- **4-week trend chart** — out of scope
- **Per-creative budget data** — Meta API doesn't expose per-ad budget
- **BC template (`_xlfn.LET`) restoration** — openpyxl limitation, abandoned

## Common Gotchas

1. **Token expires every 60 days** — `META_ACCESS_TOKEN` must be refreshed. README has instructions.
2. **Adding non-allowlisted account ID** — backend rejects with "허용되지 않은 광고 계정 ID". Must go through `/api/accounts` POST or directly edit `accounts.json`
3. **`act_` prefix required** — Meta API requires `act_` prefix on account IDs. `_normalize_account_id()` auto-adds if missing.
4. **video_views null vs 0** — null means image-only creative (hatched cell). 0 means video but no plays. Don't conflate.
5. **`A+` is a regex meta-char** — `parseCreative` and Python `_parse_creative` use `re.escape()` on the tag.
6. **Sheet names ≤ 31 chars + no `/\?*:[]`** — `_safe_sheet_name()` handles this.
7. **Excel `delete_rows` is O(n²)** — Don't use in tight loops. The deprecated BC template export hit this; new view.xlsx writes directly so it's safe.

## Adding a New Feature — Where to Look

| Feature type | Backend | Frontend |
|---|---|---|
| New metric column | `_enrich_row()` (action extraction) + `_agg_metrics_full()` | `SUMMARY_METRIC_COLS`, `DAILY_METRIC_COLS`, `METRIC_COLS` |
| New aggregation level | Already covered — JS aggregates ad-level rows | Add new state + render fn (see `renderSummaryTable`, `renderDailyTable`) |
| New filter dimension | `applySummaryFilters()` — pass through key | `SUMMARY_FILTERS` constant, `renderSummaryFilterBar()` |
| New API endpoint | `app.py` new `@app.route` | Add fetch in JS, render result |
| New Excel sheet | `_build_view_xlsx_response()` after existing sheets | (none — backend only) |

## Setup / Run

```bash
cd C:\projects\JAY
pip install -r requirements.txt
python app.py
# Open http://127.0.0.1:5000
```

`.env` required keys:
```
META_API_VERSION=v25.0
META_ACCESS_TOKEN=EAAB...
META_AD_ACCOUNT_IDS=act_1234567890,act_9876543210
```

## How to Continue Work in a New Session

When user opens a new Claude session in this project:
1. **Read this file first.**
2. Check `accounts.json` to see if account allowlist exists.
3. Check `.env` to see if token is fresh (no way to programmatically verify without API call).
4. Ask user what they want to add/change. Cross-reference with the "Deferred / Non-Goals" section above — those were intentionally postponed.
5. For Meta API documentation: https://developers.facebook.com/docs/marketing-api/reference/ads-insights/

## Project Owner Notes

- User is a marketing operator for British Council Korea
- Reports go out **weekly**, primary use case is Monday morning strategy review
- KPI varies per campaign (impressions / clicks / views / conversions)
- Excel is the eventual deliverable — Excel output structure matters more than fancy in-browser charts
- User is non-programmer, prefers concrete instructions over flexibility
- Speaks Korean primarily
