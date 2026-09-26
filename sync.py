"""
同步兼職資料 → jobs-data.js

規則：
1. 以原網站 /api/jobs 為底
2. 名稱帶【需要確認】的任務，改用試算表的資料
3. 班克爾區只在試算表有，直接從試算表加入（資料未完成，數量可能是空的）

用法：python sync.py
"""
import io
import json
import re
import sys
import urllib.request
from datetime import datetime
from pathlib import Path

import openpyxl

API_URL = "https://mabinogi-mobile-jobs.vtuberparrot2021.chatgpt.site/api/jobs"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1dY03kiMC4x3jBB4gA71jdo2YqsLAgHC7kp-WmbsNSnU/export?format=xlsx"
OUT = Path(__file__).with_name("jobs-data.js")

UNCONFIRMED = "【需要確認】"
SHEET_REGIONS = ["堤爾克那", "杜巴頓", "庫漢", "班克爾"]
SHEET_ONLY_REGIONS = ["班克爾"]


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read()


def plain_name(name):
    """去掉【跑腿】【大推】等前綴與標點，用來比對名稱"""
    name = re.sub(r"【[^】]*】", "", name)
    return re.sub(r"[？?！!\s]", "", name)


def parse_requirement(req, qty):
    """試算表的「要求」欄 → (交付物名稱, 是否跑腿, 備註)"""
    req = str(req or "").strip()
    if req.startswith("跑腿"):
        return req, True, ""
    note = ""
    if "+交付" in req:
        note, req = req.split("+交付", 1)
    m = re.match(r"^(.*?)\((.*)\)$", req)
    if m:
        req, extra = m.group(1), m.group(2)
        note = (note + " " + extra).strip()
    return req, False, note


def read_sheet(raw):
    wb = openpyxl.load_workbook(io.BytesIO(raw))
    rows = []
    for region in SHEET_REGIONS:
        if region not in wb.sheetnames:
            continue
        ws = wb[region]
        area = region
        for r in ws.iter_rows(min_row=1, values_only=True):
            a, shop, npc, name, req, qty = (list(r) + [None] * 6)[:6]
            if a and str(a).strip():
                area = str(a).strip()
            if not name or name == "任務名稱":
                continue
            qty = int(qty) if isinstance(qty, (int, float)) else None
            deliver, errand, note = parse_requirement(req, qty)
            rows.append({
                "region": region, "area": area,
                "shop": str(shop or "").strip(), "npc": str(npc or "").strip(),
                "name": str(name).strip(), "deliver": deliver,
                "qty": qty, "errand": errand, "note": note,
            })
    return rows


def sheet_row_to_job(row, incomplete):
    name = row["name"]
    if row["errand"] and not name.startswith("【"):
        name = "【跑腿】" + name
    item = {"name": row["deliver"], "quantity": 1 if row["errand"] else row["qty"]}
    job = {
        "city": row["region"], "area": row["area"],
        "shop": row["shop"], "npc": row["npc"], "name": name,
        "deliverable": item, "materials": [item],
        "afterAccept": False, "source": "sheet",
    }
    if row["note"]:
        job["note"] = row["note"]
    if incomplete:
        job["incomplete"] = True
    return job


def main():
    site = json.loads(fetch(API_URL))
    sheet = read_sheet(fetch(SHEET_URL))
    by_name = {plain_name(r["name"]): r for r in sheet}

    jobs, patched, missing = [], [], []
    for job in site["jobs"]:
        job = dict(job)
        job.pop("id", None)
        if UNCONFIRMED in job["name"]:
            row = by_name.get(plain_name(job["name"]))
            if row:
                fixed = sheet_row_to_job(row, incomplete=False)
                fixed["city"], fixed["area"] = job["city"], job["city"]
                patched.append(f'{job["name"]} → {row["deliver"]}×{row["qty"]}')
                job = fixed
            else:
                missing.append(job["name"])
        job.setdefault("area", job["city"])
        jobs.append(job)

    for row in sheet:
        if row["region"] in SHEET_ONLY_REGIONS:
            jobs.append(sheet_row_to_job(row, incomplete=True))

    # id 用名稱組成，重新同步後已勾選的任務才不會跑掉
    seen = {}
    for job in jobs:
        base = f'{job["city"]}|{job["npc"]}|{plain_name(job["name"])}'
        seen[base] = seen.get(base, 0) + 1
        job["id"] = base if seen[base] == 1 else f"{base}#{seen[base]}"

    data = {
        "jobs": jobs,
        "siteUpdatedAt": site.get("updatedAt"),
        "syncedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "incompleteRegions": SHEET_ONLY_REGIONS,
    }
    OUT.write_text(
        "// 由 sync.py 產生，請勿手動修改\nwindow.JOBS_DATA = "
        + json.dumps(data, ensure_ascii=False, indent=1) + ";\n",
        encoding="utf-8",
    )
    print(f"共 {len(jobs)} 筆任務 → {OUT.name}")
    for p in patched:
        print("  以試算表修正：", p)
    for m in missing:
        print("  ⚠ 試算表找不到：", m)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
