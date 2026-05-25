import os
import csv
import io
import json
from datetime import date, datetime, timedelta

import requests
from dotenv import load_dotenv
from flask import Flask, Response, jsonify, render_template, request

load_dotenv()

app = Flask(__name__)

META_API_VERSION = os.getenv("META_API_VERSION", "v21.0")
META_ACCESS_TOKEN = os.getenv("META_ACCESS_TOKEN", "").strip()

GRAPH_URL = f"https://graph.facebook.com/{META_API_VERSION}"


ACCOUNTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "accounts.json")


def _normalize_account_id(account_id: str) -> str:
    if not account_id:
        return ""
    account_id = account_id.strip()
    return account_id if account_id.startswith("act_") else f"act_{account_id}"


def _is_valid_account_format(account_id: str) -> bool:
    return account_id.startswith("act_") and account_id[4:].isdigit() and len(account_id) > 4


def _load_accounts_from_file():
    if not os.path.exists(ACCOUNTS_FILE):
        return None
    try:
        with open(ACCOUNTS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return [_normalize_account_id(x) for x in data.get("accounts", []) if x]
    except (json.JSONDecodeError, OSError):
        return []


def _save_accounts_to_file(ids):
    with open(ACCOUNTS_FILE, "w", encoding="utf-8") as f:
        json.dump({"accounts": ids}, f, ensure_ascii=False, indent=2)


def _seed_from_env():
    raw = os.getenv("META_AD_ACCOUNT_IDS", "").strip() or os.getenv("META_AD_ACCOUNT_ID", "").strip()
    return [_normalize_account_id(x) for x in raw.split(",") if x.strip()]


def _initial_accounts():
    """accounts.json이 있으면 그걸 사용, 없으면 .env에서 시드하고 파일 생성."""
    from_file = _load_accounts_from_file()
    if from_file is not None:
        return from_file
    seeded = _seed_from_env()
    if seeded:
        _save_accounts_to_file(seeded)
    return seeded


ACCOUNT_IDS = _initial_accounts()
_account_meta_cache = {}

FIELDS = [
    "campaign_name",
    "adset_name",
    "ad_name",
    "impressions",
    "clicks",
    "spend",
    "cpc",
    "cpm",
    "ctr",
    "reach",
    "frequency",
]

CSV_COLUMNS = [
    "date_start",
    "date_stop",
    "campaign_name",
    "adset_name",
    "ad_name",
    "impressions",
    "clicks",
    "spend",
    "cpc",
    "cpm",
    "ctr",
    "reach",
    "frequency",
]


def _resolve_account_id(requested):
    """요청 파라미터로 받은 account_id를 검증/정규화. 없으면 첫 번째 계정 사용."""
    if requested:
        normalized = _normalize_account_id(requested)
        if normalized in ACCOUNT_IDS:
            return normalized
        return None  # 허용되지 않은 ID
    return ACCOUNT_IDS[0] if ACCOUNT_IDS else None


def fetch_account_meta(account_id: str):
    """광고 계정의 이름/통화/타임존 등 메타정보 조회. 메모리 캐시 사용."""
    if account_id in _account_meta_cache:
        return _account_meta_cache[account_id]

    fallback = {"id": account_id, "name": account_id}
    if not META_ACCESS_TOKEN:
        return fallback

    try:
        r = requests.get(
            f"{GRAPH_URL}/{account_id}",
            params={
                "fields": "name,currency,timezone_name,account_status",
                "access_token": META_ACCESS_TOKEN,
            },
            timeout=15,
        )
        data = r.json()
    except requests.RequestException:
        return fallback

    if "error" in data:
        meta = {**fallback, "error": data["error"].get("message")}
    else:
        meta = {
            "id": account_id,
            "name": data.get("name", account_id),
            "currency": data.get("currency"),
            "timezone": data.get("timezone_name"),
            "status": data.get("account_status"),
        }
    _account_meta_cache[account_id] = meta
    return meta


def fetch_insights(date_preset=None, since=None, until=None, level="campaign", account_id=None):
    if not META_ACCESS_TOKEN or not ACCOUNT_IDS:
        return {
            "error": "META_ACCESS_TOKEN 또는 META_AD_ACCOUNT_IDS가 비어 있습니다. .env 파일을 채워주세요.",
            "code": "missing_credentials",
        }

    resolved = _resolve_account_id(account_id)
    if not resolved:
        return {
            "error": "허용되지 않은 광고 계정 ID입니다. .env의 META_AD_ACCOUNT_IDS에 추가해주세요.",
            "code": "invalid_account",
        }
    url = f"{GRAPH_URL}/{resolved}/insights"

    params = {
        "fields": ",".join(FIELDS),
        "level": level,
        "access_token": META_ACCESS_TOKEN,
        "limit": 200,
    }

    if since and until:
        params["time_range"] = json.dumps({"since": since, "until": until})
    else:
        params["date_preset"] = date_preset or "last_7d"

    try:
        r = requests.get(url, params=params, timeout=30)
        data = r.json()
    except requests.RequestException as e:
        return {"error": f"네트워크 오류: {e}"}

    if "error" in data:
        err = data["error"]
        return {
            "error": err.get("message", "Meta API 오류"),
            "code": err.get("code"),
            "type": err.get("type"),
        }

    return {"data": data.get("data", [])}


def _get_week_ranges(today=None):
    """이번 주(월요일~오늘) vs 같은 길이의 지난 주 (지난 주 월요일~동일 요일).

    오늘이 월요일이면 1일치 vs 1일치 비교가 되어 공정한 비교가 됩니다.
    """
    today = today or date.today()
    days_since_monday = today.weekday()  # Mon=0, Sun=6
    this_start = today - timedelta(days=days_since_monday)
    this_end = today
    last_start = this_start - timedelta(days=7)
    last_end = last_start + timedelta(days=days_since_monday)
    return this_start, this_end, last_start, last_end


def _to_float(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _aggregate(rows):
    impressions = sum(_to_float(r.get("impressions")) for r in rows)
    clicks = sum(_to_float(r.get("clicks")) for r in rows)
    spend = sum(_to_float(r.get("spend")) for r in rows)
    return {
        "impressions": impressions,
        "clicks": clicks,
        "spend": spend,
        "ctr": (clicks / impressions * 100) if impressions else 0.0,
        "cpc": (spend / clicks) if clicks else 0.0,
        "cpm": (spend / impressions * 1000) if impressions else 0.0,
    }


def _pct_change(now, prev):
    if not prev:
        return None
    return (now - prev) / prev * 100


def _build_deltas(this_total, last_total):
    out = {}
    for k in ("spend", "impressions", "clicks", "ctr", "cpc", "cpm"):
        out[k] = {
            "this": this_total[k],
            "last": last_total[k],
            "abs": this_total[k] - last_total[k],
            "pct": _pct_change(this_total[k], last_total[k]),
        }
    return out


def _build_movers(this_rows, last_rows, top_n=3):
    this_by = {r.get("campaign_name", ""): r for r in this_rows}
    last_by = {r.get("campaign_name", ""): r for r in last_rows}
    names = set(this_by) | set(last_by)
    movers = []
    for name in names:
        if not name:
            continue
        this_spend = _to_float(this_by.get(name, {}).get("spend"))
        last_spend = _to_float(last_by.get(name, {}).get("spend"))
        movers.append({
            "campaign_name": name,
            "this_spend": this_spend,
            "last_spend": last_spend,
            "delta": this_spend - last_spend,
            "delta_pct": _pct_change(this_spend, last_spend),
        })
    movers.sort(key=lambda x: x["delta"], reverse=True)
    up = [m for m in movers if m["delta"] > 0][:top_n]
    down = [m for m in reversed(movers) if m["delta"] < 0][:top_n]
    return up, down


def _fmt_won(v):
    return f"{int(round(_to_float(v))):,}원"


def _fmt_int(v):
    return f"{int(round(_to_float(v))):,}"


def _fmt_pct_delta(pct):
    if pct is None:
        return "비교 불가"
    arrow = "▲" if pct > 0 else ("▼" if pct < 0 else "—")
    return f"{arrow} {abs(pct):.1f}%"


def _build_bullets(this_total, last_total, deltas, up, down, period):
    bullets = []

    sp = deltas["spend"]
    direction = "증가" if sp["abs"] > 0 else ("감소" if sp["abs"] < 0 else "유지")
    bullets.append(
        f"이번 주({period['this']['start']}~{period['this']['end']}) 총 지출은 "
        f"{_fmt_won(sp['this'])}로 지난 주({_fmt_won(sp['last'])}) 대비 "
        f"{_fmt_pct_delta(sp['pct'])} {direction}했습니다."
    )

    cl = deltas["clicks"]
    imp = deltas["impressions"]
    bullets.append(
        f"노출 {_fmt_int(imp['this'])}회 ({_fmt_pct_delta(imp['pct'])}), "
        f"클릭 {_fmt_int(cl['this'])}회 ({_fmt_pct_delta(cl['pct'])})"
    )

    ctr = deltas["ctr"]
    cpc = deltas["cpc"]
    bullets.append(
        f"효율: CTR {ctr['this']:.2f}% ({_fmt_pct_delta(ctr['pct'])}), "
        f"CPC {_fmt_won(cpc['this'])} ({_fmt_pct_delta(cpc['pct'])})"
    )

    if up:
        m = up[0]
        bullets.append(
            f"지출 증가 1위 캠페인: \"{m['campaign_name']}\" — "
            f"{_fmt_won(m['last_spend'])} → {_fmt_won(m['this_spend'])} "
            f"({_fmt_pct_delta(m['delta_pct'])})"
        )

    if down:
        m = down[0]
        bullets.append(
            f"지출 감소 1위 캠페인: \"{m['campaign_name']}\" — "
            f"{_fmt_won(m['last_spend'])} → {_fmt_won(m['this_spend'])} "
            f"({_fmt_pct_delta(m['delta_pct'])})"
        )

    return bullets


@app.route("/")
def index():
    token_ready = bool(META_ACCESS_TOKEN)
    return render_template(
        "index.html",
        credentials_ready=token_ready,
        has_accounts=bool(ACCOUNT_IDS),
    )


@app.route("/api/accounts", methods=["GET"])
def api_accounts_list():
    accounts = [fetch_account_meta(aid) for aid in ACCOUNT_IDS]
    return jsonify({
        "accounts": accounts,
        "default": ACCOUNT_IDS[0] if ACCOUNT_IDS else None,
    })


@app.route("/api/accounts", methods=["POST"])
def api_accounts_add():
    if not META_ACCESS_TOKEN:
        return jsonify({"error": "META_ACCESS_TOKEN이 비어 있습니다. .env를 확인하세요."}), 400

    body = request.get_json(silent=True) or {}
    raw = (body.get("id") or "").strip()
    if not raw:
        return jsonify({"error": "광고 계정 ID를 입력해주세요."}), 400

    normalized = _normalize_account_id(raw)
    if not _is_valid_account_format(normalized):
        return jsonify({"error": f"올바른 형식이 아닙니다: '{normalized}'. 숫자만 또는 act_숫자 형태여야 합니다."}), 400

    if normalized in ACCOUNT_IDS:
        return jsonify({"error": "이미 등록된 계정입니다.", "id": normalized}), 409

    # Meta API 호출로 토큰 권한 + 계정 존재 검증
    try:
        r = requests.get(
            f"{GRAPH_URL}/{normalized}",
            params={
                "fields": "name,currency,timezone_name,account_status",
                "access_token": META_ACCESS_TOKEN,
            },
            timeout=15,
        )
        data = r.json()
    except requests.RequestException as e:
        return jsonify({"error": f"네트워크 오류: {e}"}), 400

    if "error" in data:
        msg = data["error"].get("message", "Meta API 오류")
        return jsonify({"error": f"검증 실패: {msg}"}), 400

    ACCOUNT_IDS.append(normalized)
    _save_accounts_to_file(ACCOUNT_IDS)
    _account_meta_cache.pop(normalized, None)
    meta = fetch_account_meta(normalized)
    return jsonify({"ok": True, "account": meta}), 201


@app.route("/api/accounts/<account_id>", methods=["DELETE"])
def api_accounts_delete(account_id):
    normalized = _normalize_account_id(account_id)
    if normalized not in ACCOUNT_IDS:
        return jsonify({"error": "등록되지 않은 계정입니다."}), 404
    ACCOUNT_IDS.remove(normalized)
    _save_accounts_to_file(ACCOUNT_IDS)
    _account_meta_cache.pop(normalized, None)
    return jsonify({"ok": True, "id": normalized})


@app.route("/api/weekly-insights")
def api_weekly_insights():
    account_id = request.args.get("account_id")
    this_start, this_end, last_start, last_end = _get_week_ranges()

    this_result = fetch_insights(
        since=this_start.isoformat(),
        until=this_end.isoformat(),
        level="campaign",
        account_id=account_id,
    )
    if "error" in this_result:
        return jsonify(this_result), 400

    last_result = fetch_insights(
        since=last_start.isoformat(),
        until=last_end.isoformat(),
        level="campaign",
        account_id=account_id,
    )
    if "error" in last_result:
        return jsonify(last_result), 400

    this_rows = this_result["data"]
    last_rows = last_result["data"]

    this_total = _aggregate(this_rows)
    last_total = _aggregate(last_rows)
    deltas = _build_deltas(this_total, last_total)
    up, down = _build_movers(this_rows, last_rows)

    period = {
        "this": {"start": this_start.isoformat(), "end": this_end.isoformat()},
        "last": {"start": last_start.isoformat(), "end": last_end.isoformat()},
        "days": (this_end - this_start).days + 1,
    }

    bullets = _build_bullets(this_total, last_total, deltas, up, down, period)

    return jsonify({
        "period": period,
        "totals": {"this": this_total, "last": last_total},
        "deltas": deltas,
        "top_spend_up": up,
        "top_spend_down": down,
        "bullets": bullets,
    })


@app.route("/api/insights")
def api_insights():
    result = fetch_insights(
        date_preset=request.args.get("date_preset"),
        since=request.args.get("since"),
        until=request.args.get("until"),
        level=request.args.get("level", "campaign"),
        account_id=request.args.get("account_id"),
    )
    status = 400 if "error" in result else 200
    return jsonify(result), status


@app.route("/api/insights.csv")
def api_insights_csv():
    result = fetch_insights(
        date_preset=request.args.get("date_preset"),
        since=request.args.get("since"),
        until=request.args.get("until"),
        level=request.args.get("level", "campaign"),
        account_id=request.args.get("account_id"),
    )
    if "error" in result:
        return result["error"], 400

    rows = result["data"]
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=CSV_COLUMNS, extrasaction="ignore")
    writer.writeheader()
    for row in rows:
        writer.writerow({k: row.get(k, "") for k in CSV_COLUMNS})

    csv_data = "﻿" + buf.getvalue()
    filename = f"meta_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response(
        csv_data,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
