#!/usr/bin/env python3
"""
Feishu Sheet → data.json refresh (no external libraries needed).

Required env vars (set as GitHub Secrets):
    FEISHU_APP_ID       Feishu app_id
    FEISHU_APP_SECRET   Feishu app_secret

Usage:
    python feishu_refresh.py          # write data.json locally
    python feishu_refresh.py --push   # also git commit + push
"""

import argparse
import datetime
import json
import os
import subprocess
import sys
import urllib.error
import urllib.request
from pathlib import Path

REPO_DIR = Path(__file__).parent
DATA_JSON = REPO_DIR / "data.json"

SPREADSHEET_TOKEN = "WWlhsS3qMhid3zt0DcLcKbiXnJD"
SHEET_PAGE   = "0UUcJp"  # 页面数据
SHEET_FLOW   = "1TBAmn"  # 导流页面数据
SHEET_EVENTS = "2wRaco"  # 组件数据
SHEET_META   = "3kTqZz"  # 组件埋点 (unused at runtime — componentMeta comes from data.json)

FEISHU_API = "https://open.feishu.cn/open-apis"


# ── Auth ──────────────────────────────────────────────────────────────────────

def get_access_token():
    app_id = os.environ.get("FEISHU_APP_ID", "")
    app_secret = os.environ.get("FEISHU_APP_SECRET", "")
    user_refresh_token = os.environ.get("FEISHU_USER_REFRESH_TOKEN", "")

    if user_refresh_token:
        return _refresh_user_token(app_id, app_secret, user_refresh_token)

    if not app_id or not app_secret:
        raise RuntimeError(
            "Environment variables FEISHU_APP_ID and FEISHU_APP_SECRET must be set."
        )
    url = f"{FEISHU_API}/auth/v3/tenant_access_token/internal"
    body = json.dumps({"app_id": app_id, "app_secret": app_secret}).encode()
    req = urllib.request.Request(
        url, data=body, method="POST",
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req) as r:
        resp = json.loads(r.read())
    if resp.get("code") != 0:
        raise RuntimeError(f"Feishu auth failed: {resp}")
    return resp["tenant_access_token"]


def _refresh_user_token(app_id, app_secret, refresh_token):
    url = f"{FEISHU_API}/authen/v2/oauth/token"
    payload = json.dumps({
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": app_id,
        "client_secret": app_secret,
    }).encode()
    req = urllib.request.Request(url, data=payload, method="POST",
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req) as r:
            resp = json.loads(r.read())
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"User token refresh failed: HTTP {e.code} — {err_body[:300]}")

    access_token = resp.get("access_token")
    new_refresh_token = resp.get("refresh_token")

    if not access_token:
        raise RuntimeError(f"Failed to get user_access_token: {resp}")

    print("  Using user_access_token (personal OAuth)")

    if new_refresh_token and new_refresh_token != refresh_token:
        _rotate_github_secret("FEISHU_USER_REFRESH_TOKEN", new_refresh_token)

    return access_token


def _rotate_github_secret(secret_name, new_value):
    import base64
    # Requires REPO_PAT secret (classic PAT with repo scope) — optional.
    # Without it, rotation is skipped and the refresh_token silently stays as-is.
    # Daily runs keep the sliding 30-day window alive so expiry is unlikely.
    github_token = os.environ.get("REPO_PAT", "") or os.environ.get("GITHUB_TOKEN", "")
    github_repo = os.environ.get("GITHUB_REPOSITORY", "")
    if not github_token or not github_repo:
        print(f"  [skip] Cannot auto-rotate {secret_name} (REPO_PAT not set)")
        return
    try:
        from nacl import public  # PyNaCl
    except ImportError:
        print(f"  [warn] PyNaCl not installed — {secret_name} not rotated")
        return

    api = f"https://api.github.com/repos/{github_repo}/actions/secrets"
    headers = {
        "Authorization": f"Bearer {github_token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    req = urllib.request.Request(f"{api}/public-key", headers=headers)
    with urllib.request.urlopen(req) as r:
        pk = json.loads(r.read())

    pk_bytes = base64.b64decode(pk["key"])
    sealed = public.SealedBox(public.PublicKey(pk_bytes)).encrypt(new_value.encode())
    encrypted = base64.b64encode(sealed).decode()

    body = json.dumps({"encrypted_value": encrypted, "key_id": pk["key_id"]}).encode()
    req = urllib.request.Request(
        f"{api}/{secret_name}", data=body, method="PUT",
        headers={**headers, "Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req) as r:
        pass
    print(f"  Rotated GitHub Secret: {secret_name}")


# ── Sheet reading ─────────────────────────────────────────────────────────────

def read_sheet(access_token, sheet_id, max_rows=5000):
    """Return data rows (header row excluded), trailing empty rows stripped."""
    import urllib.parse
    range_str = f"{sheet_id}!A1:Z{max_rows}"
    encoded_range = urllib.parse.quote(range_str, safe="")
    url = f"{FEISHU_API}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{encoded_range}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {access_token}"})
    try:
        with urllib.request.urlopen(req) as r:
            resp = json.loads(r.read())
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Sheet read failed ({sheet_id}): HTTP {e.code} {e.reason} — {body[:500]}")
    if resp.get("code") != 0:
        raise RuntimeError(f"Sheet read failed ({sheet_id}): {resp}")
    all_rows = (resp.get("data") or {}).get("valueRange", {}).get("values") or []
    rows = all_rows[1:]  # skip header
    while rows and all(v is None or v == "" for v in rows[-1]):
        rows.pop()
    print(f"    -> {len(rows)} rows")
    return rows


# ── Date parsing ──────────────────────────────────────────────────────────────

def parse_date(raw):
    """Feishu may return dates as an Excel serial float or a formatted string."""
    if raw is None:
        raise ValueError("date is None")
    if isinstance(raw, (int, float)):
        n = int(raw)
        d = datetime.date(1899, 12, 30) + datetime.timedelta(days=n)
        return n, d.strftime("%Y-%m-%d")
    s = str(raw).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y/%m/%d 0:00:00", "%m/%d/%Y"):
        try:
            d = datetime.datetime.strptime(s, fmt).date()
            n = (d - datetime.date(1899, 12, 30)).days
            return n, d.strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValueError(f"Cannot parse date: {raw!r}")


def _cell(row, idx, default=None):
    v = row[idx] if len(row) > idx else default
    return v if v is not None else default


# ── Data build ────────────────────────────────────────────────────────────────

def build_data(existing_data: dict) -> dict:
    print("Authenticating with Feishu...")
    token = get_access_token()

    # componentMeta always comes from data.json (stable config)
    comp_meta = existing_data.get("componentMeta", [])
    if not comp_meta:
        raise RuntimeError(
            "componentMeta not found in data.json. "
            "Run update.py from the Excel file first to initialise it."
        )
    logkey_map = {}
    for m in comp_meta:
        if m.get("exposureKey"):
            logkey_map[m["exposureKey"]] = ("exposure", m["component"], m["exposureRule"])
        if m.get("clickKey"):
            logkey_map[m["clickKey"]] = ("click", m["component"], m["exposureRule"])

    # ── 页面数据: (d, region, locale, uv, avg_duration) ──
    print("  Reading 页面数据...")
    page_map = {}
    for row in read_sheet(token, SHEET_PAGE):
        if not row or row[0] is None:
            continue
        try:
            d_num, date_str = parse_date(row[0])
        except ValueError:
            continue
        region = str(_cell(row, 1, ""))
        locale = str(_cell(row, 2, ""))
        uv = int(_cell(row, 3, 0) or 0)
        key = (d_num, region, locale)
        page_map[key] = {
            "d": d_num, "date": date_str,
            "region": region, "locale": locale,
            "pageUV": uv, "flowUV": 0,
        }

    # ── 导流页面数据: (d, region, locale, uv) ──
    print("  Reading 导流页面数据...")
    for row in read_sheet(token, SHEET_FLOW):
        if not row or row[0] is None:
            continue
        try:
            d_num, _ = parse_date(row[0])
        except ValueError:
            continue
        key = (d_num, str(_cell(row, 1, "")), str(_cell(row, 2, "")))
        if key in page_map:
            page_map[key]["flowUV"] = int(_cell(row, 3, 0) or 0)

    pageRows = sorted(page_map.values(), key=lambda x: (x["date"], x["region"], x["locale"]))

    # ── 组件数据: (d, region, locale, logkey, uv) ──
    print("  Reading 组件数据...")
    comp_data = {}
    for row in read_sheet(token, SHEET_EVENTS):
        if not row or row[0] is None:
            continue
        try:
            d_num, date_str = parse_date(row[0])
        except ValueError:
            continue
        region  = str(_cell(row, 1, ""))
        locale  = str(_cell(row, 2, ""))
        logkey  = str(_cell(row, 3, "")).strip()
        uv      = int(_cell(row, 4, 0) or 0)
        if logkey not in logkey_map:
            continue
        kind, comp_name, rule = logkey_map[logkey]
        key = (d_num, region, locale, comp_name)
        if key not in comp_data:
            comp_data[key] = {
                "d": d_num, "date": date_str,
                "region": region, "locale": locale,
                "component": comp_name,
                "exposureUV": 0, "clickUV": 0,
                "exposureRule": rule,
            }
        if kind == "exposure":
            comp_data[key]["exposureUV"] = uv
        else:
            comp_data[key]["clickUV"] = uv

    # Ensure every component × every (date, region, locale) has a row
    for m in comp_meta:
        is_page_uv = m["exposureRule"] == "页面UV"
        for pr in pageRows:
            key = (pr["d"], pr["region"], pr["locale"], m["component"])
            if key not in comp_data:
                comp_data[key] = {
                    "d": pr["d"], "date": pr["date"],
                    "region": pr["region"], "locale": pr["locale"],
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

def git_push():
    def run(cmd):
        r = subprocess.run(cmd, cwd=REPO_DIR, capture_output=True, text=True)
        if r.returncode != 0:
            print(r.stderr.strip())
            return False
        return True

    today = datetime.date.today().isoformat()
    msg = f"Auto-refresh data: {today} (Feishu)"
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
    args = parser.parse_args()

    if not DATA_JSON.exists():
        print("ERROR: data.json not found. Run update.py from Excel first.")
        sys.exit(1)

    with open(DATA_JSON, encoding="utf-8") as f:
        existing = json.load(f)

    data = build_data(existing)

    print(f"pageRows: {len(data['pageRows'])}, componentRows: {len(data['componentRows'])}")
    print(f"Date range: {data['summary']['dateMin']} ~ {data['summary']['dateMax']}")

    with open(DATA_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    print(f"data.json updated ({DATA_JSON.stat().st_size // 1024} KB)")

    if args.push:
        git_push()


if __name__ == "__main__":
    main()
