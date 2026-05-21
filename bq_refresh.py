#!/usr/bin/env python3
"""
BQ → data.json refresh.

Local usage (requires `gcloud auth application-default login`):
    python bq_refresh.py

GitHub Actions usage (SA key in GOOGLE_APPLICATION_CREDENTIALS):
    python bq_refresh.py --push

Options:
    --push          git commit + push after writing data.json
    --days N        rolling window size in days (default: 7)
    --lag N         exclude the last N days (default: 1, waits for data completeness)
"""

import argparse
import datetime
import json
import subprocess
import sys
from pathlib import Path

REPO_DIR = Path(__file__).parent
DATA_JSON = REPO_DIR / "data.json"

PROJECT = "trip-ibu-bi-dw-etl"

LOCALES = [
    "zh-hk", "en-hk", "ko-kr", "zh-tw", "th-th", "en-th", "ja-jp",
    "en-sg", "zh-sg", "en-my", "zh-my", "ms-my", "en-id", "id-id",
]

# ── SQL templates ─────────────────────────────────────────────────────────────

SQL_PAGE = """\
SELECT
    FORMAT_DATE('%Y-%m-%d', d) AS date,
    region,
    locale,
    COUNT(DISTINCT context.vid) AS uv,
    SAFE_DIVIDE(SUM(context.duration), COUNT(DISTINCT context.vid)) AS avg_duration
FROM `trip-ibu-bi-dw-etl.ibu_bi_dw_cdw.edw_usr_ubt_ibu_pageview`
WHERE d >= '{start}' AND d < '{end}'
  AND locale IN ({locales})
  AND page.p_page = '10651196530'
  AND ua_channeltype = 'app'
  AND iscrawler = 0
GROUP BY 1, 2, 3
"""

SQL_FLOW = """\
SELECT
    FORMAT_DATE('%Y-%m-%d', d) AS date,
    region,
    locale,
    COUNT(DISTINCT context.vid) AS uv
FROM `trip-ibu-bi-dw-etl.ibu_bi_dw_cdw.edw_usr_ubt_ibu_pageview`
WHERE d >= '{start}' AND d < '{end}'
  AND locale IN ({locales})
  AND context.frompage = '10651196530'
  AND page.p_pageid <> '10651196530'
  AND prdtype IN ('F', 'H', 'P', 'C', 'A', 'X', 'CRU')
  AND ua_channeltype = 'app'
  AND iscrawler = 0
GROUP BY 1, 2, 3
"""

# NOTE: added region + locale to GROUP BY vs. the original SQL,
# so the dashboard can show per-region component CTR.
SQL_EVENTS = """\
SELECT
    FORMAT_DATE('%Y-%m-%d', d) AS date,
    region,
    locale,
    logkey,
    COUNT(DISTINCT vid) AS uv
FROM `trip-ibu-bi-dw-etl.ibu_bi_dw_cdw.edw_usr_ubt_foreend_click`
WHERE d >= '{start}' AND d < '{end}'
  AND logkey IN ({logkeys})
  AND locale IN ({locales})
  AND ua_channeltype = 'app'
GROUP BY 1, 2, 3, 4
"""


# ── BQ helpers ────────────────────────────────────────────────────────────────

def ensure_bq_client():
    try:
        from google.cloud import bigquery
    except ImportError:
        print("Installing google-cloud-bigquery…")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "google-cloud-bigquery"],
            check=True,
        )
        from google.cloud import bigquery
    return bigquery.Client(project=PROJECT)


def run_query(client, sql, description=""):
    print(f"  Querying {description}…")
    rows = list(client.query(sql).result())
    print(f"    → {len(rows)} rows")
    return [dict(row) for row in rows]


# ── Date helpers ──────────────────────────────────────────────────────────────

def date_serial(date_str):
    """'2026-05-14' → Excel serial int (used by dashboard JS)."""
    d = datetime.date.fromisoformat(date_str)
    return (d - datetime.date(1899, 12, 30)).days


# ── Data build ────────────────────────────────────────────────────────────────

def build_data(start: str, end: str, existing_data: dict) -> dict:
    """Query BQ and return the full DATA dict ready for data.json."""
    client = ensure_bq_client()

    locale_sql = ", ".join(f"'{l}'" for l in LOCALES)

    # ── componentMeta: always reuse from data.json (stable config) ──
    comp_meta = existing_data.get("componentMeta", [])
    if not comp_meta:
        raise RuntimeError(
            "componentMeta not found in data.json. "
            "Run update.py from the Excel file first to initialise it."
        )
    logkey_map = {}  # logkey → ('exposure'|'click', comp_name, rule)
    all_logkeys = set()
    for m in comp_meta:
        if m.get("exposureKey"):
            logkey_map[m["exposureKey"]] = ("exposure", m["component"], m["exposureRule"])
            all_logkeys.add(m["exposureKey"])
        if m.get("clickKey"):
            logkey_map[m["clickKey"]] = ("click", m["component"], m["exposureRule"])
            all_logkeys.add(m["clickKey"])

    logkey_sql = ", ".join(f"'{k}'" for k in all_logkeys)

    fmt = dict(start=start, end=end, locales=locale_sql, logkeys=logkey_sql)

    # ── Page traffic ──
    page_rows_bq = run_query(client, SQL_PAGE.format(**fmt), "page traffic")
    flow_rows_bq = run_query(client, SQL_FLOW.format(**fmt), "flow traffic")

    page_map = {}
    for r in page_rows_bq:
        d_str = str(r["date"])
        key = (d_str, str(r["region"]), str(r["locale"]))
        page_map[key] = {
            "d": date_serial(d_str),
            "date": d_str,
            "region": str(r["region"]),
            "locale": str(r["locale"]),
            "pageUV": int(r["uv"] or 0),
            "flowUV": 0,
        }
    for r in flow_rows_bq:
        d_str = str(r["date"])
        key = (d_str, str(r["region"]), str(r["locale"]))
        if key in page_map:
            page_map[key]["flowUV"] = int(r["uv"] or 0)

    pageRows = sorted(page_map.values(), key=lambda x: (x["date"], x["region"], x["locale"]))

    # ── Component events ──
    events_bq = run_query(client, SQL_EVENTS.format(**fmt), "component events")

    page_by_key = {(r["date"], r["region"], r["locale"]): r for r in pageRows}
    comp_data = {}

    for r in events_bq:
        d_str = str(r["date"])
        logkey = str(r["logkey"])
        if logkey not in logkey_map:
            continue
        kind, comp_name, rule = logkey_map[logkey]
        key = (d_str, str(r["region"]), str(r["locale"]), comp_name)
        if key not in comp_data:
            comp_data[key] = {
                "d": date_serial(d_str),
                "date": d_str,
                "region": str(r["region"]),
                "locale": str(r["locale"]),
                "component": comp_name,
                "exposureUV": 0,
                "clickUV": 0,
                "exposureRule": rule,
            }
        if kind == "exposure":
            comp_data[key]["exposureUV"] = int(r["uv"] or 0)
        else:
            comp_data[key]["clickUV"] = int(r["uv"] or 0)

    # Ensure every component × every (date, region, locale) has a row
    for m in comp_meta:
        is_page_uv = m["exposureRule"] == "页面UV"
        for pr in pageRows:
            key = (pr["date"], pr["region"], pr["locale"], m["component"])
            if key not in comp_data:
                comp_data[key] = {
                    "d": pr["d"],
                    "date": pr["date"],
                    "region": pr["region"],
                    "locale": pr["locale"],
                    "component": m["component"],
                    "exposureUV": pr["pageUV"] if is_page_uv else 0,
                    "clickUV": 0,
                    "exposureRule": m["exposureRule"],
                }
            elif is_page_uv:
                comp_data[key]["exposureUV"] = pr["pageUV"]

    componentRows = sorted(
        comp_data.values(),
        key=lambda x: (x["date"], x["region"], x["locale"], x["component"]),
    )

    # ── Summary ──
    from update import _compute_summary
    summary = _compute_summary(pageRows, componentRows, comp_meta)
    summary["updateTime"] = datetime.date.today().isoformat()

    return {
        "pageRows": pageRows,
        "componentRows": componentRows,
        "componentMeta": comp_meta,
        "summary": summary,
    }


# ── Git push ──────────────────────────────────────────────────────────────────

def git_push(start: str, end: str):
    def run(cmd):
        r = subprocess.run(cmd, cwd=REPO_DIR, capture_output=True, text=True)
        if r.returncode != 0:
            print(r.stderr.strip())
            return False
        return True

    msg = f"Auto-refresh data: {start} ~ {end} (BQ)"
    ok = (
        run(["git", "add", "data.json"])
        and run(["git", "commit", "-m", msg])
        and run(["git", "push"])
    )
    if ok:
        print(f"Pushed: {msg}")
    else:
        print("Push failed — check git config or network")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--push", action="store_true", help="git commit + push after writing")
    parser.add_argument("--days", type=int, default=7, help="rolling window size in days")
    parser.add_argument("--lag", type=int, default=1, help="exclude last N days (data lag)")
    args = parser.parse_args()

    end_date = datetime.date.today() - datetime.timedelta(days=args.lag)
    start_date = end_date - datetime.timedelta(days=args.days)
    start, end = start_date.isoformat(), end_date.isoformat()

    print(f"Date window: {start} ~ {end} ({args.days}d, lag={args.lag}d)")

    if not DATA_JSON.exists():
        print("ERROR: data.json not found. Run update.py from Excel first.")
        sys.exit(1)

    with open(DATA_JSON, encoding="utf-8") as f:
        existing = json.load(f)

    data = build_data(start, end, existing)

    print(f"pageRows: {len(data['pageRows'])}, componentRows: {len(data['componentRows'])}")
    print(f"Date range in data: {data['summary']['dateMin']} ~ {data['summary']['dateMax']}")

    with open(DATA_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    print(f"data.json updated ({DATA_JSON.stat().st_size // 1024} KB)")

    if args.push:
        git_push(start, end)


if __name__ == "__main__":
    main()
