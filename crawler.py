#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fill brand-by-city monthly sales into a folder of brand workbooks and compute
the share of each brand-city market taken by a given set of placements.

Each workbook holds one brand, one row per (province, branch city, month):

  A 品牌名称   B 展业省份   C 展业城市（分公司）   D 投放月份   E 投放数（台）
  F 品牌城市月销量（台）   G 月销量市占比

F is filled from a sales source; G is written as the live Excel formula =E/F.

SALES SOURCES
-------------
  demo   deterministic synthetic figures, no network (default)
  http   POST to $SALES_API_URL, with $SALES_API_TOKEN sent as a bearer token
         if set.
           request   {"date": "YYYY-MM", "area": "广东省:广州", "type": 0}
           response  {"data": [{"name": "比亚迪", "sales": "1,234"}, ...]}
         type 0 returns a brand ranking, type 1 a model ranking (used for the
         HarmonyOS 鸿蒙智行 marques, which the source lists by model only).

One request for a (month, city) returns every brand at once, so responses are
cached and each (month, city) is fetched a single time per run.

The tool was built against a third-party sales-ranking API that requires a
signed request token. That client is not published; `http` stands in for it.

USAGE
-----
  python crawler.py                    # fill the workbooks next to this file
  python crawler.py --dir workbooks    # a different folder
  python crawler.py --dry-run          # fetch and report, write nothing
  python crawler.py --extend           # also add rows for new months, then fill
  python crawler.py --extend --through 26年12月
  python crawler.py --source http      # use $SALES_API_URL instead of demo data
"""

import argparse
import datetime
import glob
import hashlib
import json
import os
import sys
import time

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Missing openpyxl. Run:  pip install -r requirements.txt")


# ------------------------------------------------------------------ mappings

# Workbook brand name -> name used by the sales source. "SUB:<kw>" means the
# source lists this marque under 鸿蒙智行 by model, so every model whose name
# contains <kw> is summed instead.
BRAND_MAP = {
    "AITO 问界": "SUB:问界", "LUXEED 智界": "SUB:智界",
    "享界": "SUB:享界", "尚界": "SUB:尚界",
    "smart": "smart", "上汽奥迪": "上汽奥迪", "上汽通用别克": "别克",
    "东风奕派": "东风奕派", "东风日产": "东风日产", "东风风神": "东风风神",
    "东风风行": "东风风行", "北京奔驰": "奔驰", "吉利汽车": "吉利", "埃安": "埃安",
    "奇瑞新能源": "奇瑞新能源", "奇瑞汽车": "奇瑞", "小米汽车": "小米",
    "小鹏汽车": "小鹏", "岚图汽车": "岚图", "方程豹": "方程豹", "星途": "星途",
    "智己汽车": "智己", "极氪": "极氪", "极石": "ROX 极石", "比亚迪": "比亚迪",
    "沃尔沃": "沃尔沃", "深蓝汽车": "深蓝", "特斯拉": "特斯拉", "理想": "理想",
    "腾势汽车": "腾势", "蔚来汽车": "蔚来", "赛力斯蓝电": "蓝电",
    "长城汽车": "长城汽车", "长安启源": "长安启源", "长安汽车": "长安",
    "阿维塔科技": "阿维塔", "零跑汽车": "零跑", "领克": "领克",
}

# Branches named after a region rather than a city are read as that region's
# capital.
CAPITAL_OF = {"内蒙古": "呼和浩特", "四川": "成都", "贵州": "贵阳"}

# Suffixes that number several branches in one city.
CITY_SUFFIXES = ("第一", "第二", "第三", "第四")


def city_to_area(province, city):
    """Excel (province, branch city) -> source area string '省:市'."""
    base = city
    for suf in CITY_SUFFIXES:
        if base.endswith(suf):
            base = base[: -len(suf)]
            break
    base = CAPITAL_OF.get(base, base)
    return f"{province}:{base}"


def month_to_date(m):
    """'25年1月' -> '2025-01'."""
    yy, rest = m.split("年")
    return f"20{int(yy):02d}-{int(rest.replace('月', '')):02d}"


def month_to_ym(m):
    """'25年1月' -> (2025, 1)."""
    yy, rest = m.split("年")
    return 2000 + int(yy), int(rest.replace("月", ""))


def ym_to_label(year, month):
    """(2025, 1) -> '25年1月' (the sheet's own style, no zero-padding)."""
    return f"{year % 100}年{month}月"


def months_range(start, end):
    """Yield (year, month) from start to end inclusive."""
    y, mo = start
    while (y, mo) <= end:
        yield (y, mo)
        mo += 1
        if mo > 12:
            mo, y = 1, y + 1


def _num(sales):
    """'13,693' -> 13693; '' or '-' -> 0."""
    if sales is None:
        return 0
    s = str(sales).replace(",", "").strip()
    if s in ("", "-"):
        return 0
    try:
        return int(float(s))
    except ValueError:
        return 0


def brand_sales_from(records, api_name):
    """One brand's sales from a list of {"name", "sales"} records."""
    if api_name.startswith("SUB:"):
        kw = api_name[4:]
        return sum(_num(r["sales"]) for r in records if kw in r["name"])
    for r in records:
        if r["name"] == api_name:
            return _num(r["sales"])
    return 0                                   # no sales in that city


# ------------------------------------------------------------- sales sources

class DemoSource:
    """Synthetic sales, stable across runs. Months from the current calendar
    month on are treated as unpublished and return nothing, as the real source
    does."""

    name = "demo (synthetic data)"
    MODELS = {
        "问界": ["问界M5", "问界M7", "问界M8", "问界M9"],
        "智界": ["智界S7", "智界R7"],
        "享界": ["享界S9"],
        "尚界": ["尚界H5"],
    }

    def __init__(self):
        today = datetime.date.today()
        self.first_unpublished = f"{today.year:04d}-{today.month:02d}"
        self.brands = sorted({v for v in BRAND_MAP.values()
                              if not v.startswith("SUB:")})
        self.models = [m for ms in self.MODELS.values() for m in ms]

    @staticmethod
    def _sales(date, area, name):
        h = int(hashlib.sha256(f"{date}|{area}|{name}".encode()).hexdigest(), 16)
        scale = 200 + h % 1300
        month = int(date[-2:])
        season = 1.0 + 0.25 * (month in (3, 6, 9, 12))   # quarter-end push
        return int(scale * season)

    def records(self, date, area, dtype):
        if date >= self.first_unpublished:
            return []
        names = self.brands if dtype == 0 else self.models
        return [{"name": n, "sales": f"{self._sales(date, area, n):,}"}
                for n in names]


class HttpSource:
    """Any JSON endpoint that follows the contract in the module docstring."""

    def __init__(self, url, token=None, retries=3):
        try:
            import requests
        except ImportError:
            sys.exit("Missing requests. Run:  pip install -r requirements.txt")
        self.url, self.retries = url, retries
        self.session = requests.Session()
        self.headers = {"Content-Type": "application/json;charset=utf-8"}
        if token:
            self.headers["Authorization"] = f"Bearer {token}"
        self.name = f"http ({url})"

    def records(self, date, area, dtype):
        body = {"date": date, "area": area, "type": dtype}
        last = None
        for attempt in range(self.retries):
            try:
                r = self.session.post(self.url, headers=self.headers,
                                      data=json.dumps(body), timeout=30)
                r.raise_for_status()
                return r.json().get("data") or []
            except Exception as e:  # noqa: BLE001
                last = e
                time.sleep(1.5 * (attempt + 1))
        raise RuntimeError(f"Source failed for {date} {area} type={dtype}: {last}")


def make_source(kind):
    if kind == "http":
        url = os.environ.get("SALES_API_URL")
        if not url:
            sys.exit("--source http needs SALES_API_URL set.")
        return HttpSource(url, os.environ.get("SALES_API_TOKEN"))
    return DemoSource()


# ------------------------------------------------------------------ the fill

HDR_BRAND = "品牌名称"
HDR_PROV = "展业省份"
HDR_CITY = "展业城市（分公司）"
HDR_MONTH = "投放月份"
HDR_PLACED = "投放数（台）"             # E: placements
HDR_TARGET = "品牌城市月销量（台）"      # F: brand sales in that city (filled)
HDR_SHARE = "月销量市占比"              # G: E / F (written as a formula)


def base_dir():
    """Folder holding this script, or the .exe when frozen by PyInstaller."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def find_workbooks(folder):
    files = sorted(glob.glob(os.path.join(folder, "*.xlsx")))
    return [f for f in files if not os.path.basename(f).startswith("~$")]


def extend_rows(ws, ci, end_ym):
    """Add a row for every missing month, per (province, city), from the
    file's earliest month through end_ym. Returns the number of rows added."""
    brand0 = None
    combos, present, min_ym = [], set(), None
    for row in ws.iter_rows(min_row=2):
        b, p, c, m = (str(row[ci[h]].value or "").strip()
                      for h in (HDR_BRAND, HDR_PROV, HDR_CITY, HDR_MONTH))
        if not (b and p and c and m):
            continue
        brand0 = brand0 or b
        if (p, c) not in combos:
            combos.append((p, c))
        present.add((p, c, m))
        ym = month_to_ym(m)
        min_ym = ym if min_ym is None else min(min_ym, ym)
    if not combos:
        return 0
    added, maxcol = 0, ws.max_column
    for (p, c) in combos:
        for (y, mo) in months_range(min_ym, end_ym):
            label = ym_to_label(y, mo)
            if (p, c, label) in present:
                continue
            r = ws.max_row + 1
            for h, v in ((HDR_BRAND, brand0), (HDR_PROV, p),
                         (HDR_CITY, c), (HDR_MONTH, label)):
                ws.cell(row=r, column=ci[h] + 1, value=v)
            for col in range(1, maxcol + 1):       # copy row-2 styling
                ws.cell(row=r, column=col)._style = ws.cell(row=2, column=col)._style
            present.add((p, c, label))
            added += 1
    return added


def safe_console():
    """Don't crash on consoles that can't print Chinese (e.g. CI logs)."""
    for s in (sys.stdout, sys.stderr):
        try:
            "中".encode(s.encoding or "ascii")
        except (UnicodeEncodeError, LookupError):
            s.reconfigure(errors="replace")


def main():
    safe_console()
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=None)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--source", choices=("demo", "http"), default=None,
                    help="default: http if SALES_API_URL is set, else demo")
    ap.add_argument("--sleep", type=float, default=0.2,
                    help="seconds between requests to the http source")
    ap.add_argument("--extend", action="store_true",
                    help="add month rows up to --through before filling")
    ap.add_argument("--through", default=None,
                    help="last month for --extend, e.g. 26年12月 "
                         "(default: the current month)")
    args = ap.parse_args()

    # Double-clicked (no arguments): ask the one question that matters.
    if len(sys.argv) == 1:
        print("=== 品牌城市月销量 抓取工具 ===")
        ans = input("是否同时补充新月份的行? Add rows for new months too? [y/N]: ")
        args.extend = ans.strip().lower() in ("y", "yes")

    if args.dir is None:
        args.dir = base_dir()
        if not find_workbooks(args.dir) and os.path.isdir(
                os.path.join(args.dir, "workbooks")):
            args.dir = os.path.join(args.dir, "workbooks")

    if args.source is None:
        args.source = "http" if os.environ.get("SALES_API_URL") else "demo"
    source = make_source(args.source)
    sleep = args.sleep if args.source == "http" else 0

    if args.through:
        end_ym = month_to_ym(args.through)
    else:
        today = datetime.date.today()
        end_ym = (today.year, today.month)

    files = find_workbooks(args.dir)
    if not files:
        sys.exit(f"No .xlsx files in {args.dir}")
    print(f"Source: {source.name}")
    print(f"Found {len(files)} workbooks in {args.dir}")

    cache = {}

    def records_for(date, area, dtype):
        k = (date, area, dtype)
        if k not in cache:
            cache[k] = source.records(date, area, dtype)
            time.sleep(sleep)
        return cache[k]

    unknown_brands = set()
    total_rows = filled = 0

    for path in files:
        wb = load_workbook(path)
        ws = wb.active
        cols = {str(c.value or "").strip(): i for i, c in enumerate(ws[1])}
        need = [HDR_BRAND, HDR_PROV, HDR_CITY, HDR_MONTH, HDR_TARGET]
        if any(h not in cols for h in need):
            print(f"  ! skipping {os.path.basename(path)} (missing columns)")
            continue
        ci = {h: cols[h] for h in need}
        placed_col, share_col = cols.get(HDR_PLACED), cols.get(HDR_SHARE)
        f_letter = ws.cell(row=1, column=ci[HDR_TARGET] + 1).column_letter
        e_letter = (ws.cell(row=1, column=placed_col + 1).column_letter
                    if placed_col is not None else None)

        added = extend_rows(ws, ci, end_ym) if args.extend else 0

        rows_filled = 0
        for row in ws.iter_rows(min_row=2):
            brand, prov, city, month = (str(row[ci[h]].value or "").strip()
                                        for h in (HDR_BRAND, HDR_PROV,
                                                  HDR_CITY, HDR_MONTH))
            if not (brand and prov and city and month):
                continue
            total_rows += 1
            api_name = BRAND_MAP.get(brand)
            if not api_name:
                unknown_brands.add(brand)
                continue

            dtype = 1 if api_name.startswith("SUB:") else 0
            recs = records_for(month_to_date(month), city_to_area(prov, city), dtype)
            if not recs:
                continue                 # month not published yet: leave blank
            val = brand_sales_from(recs, api_name)
            row[ci[HDR_TARGET]].value = val
            rows_filled += 1
            filled += 1

            # G = E/F as a live formula; blank when F is 0 so nothing divides by zero.
            if share_col is not None and e_letter is not None:
                g = row[share_col]
                if val:
                    g.value = f"={e_letter}{g.row}/{f_letter}{g.row}"
                    g.number_format = "0.00%"
                else:
                    g.value = None

        if not args.dry_run:
            wb.save(path)
        print(f"  {os.path.basename(path)}: filled {rows_filled} rows"
              + (f", added {added} new-month rows" if added else "")
              + ("  (dry run, not saved)" if args.dry_run else ""))

    print(f"\nDone. {filled}/{total_rows} rows filled, {len(cache)} unique requests.")
    if unknown_brands:
        print("Unmapped brands (add them to BRAND_MAP):")
        for b in sorted(unknown_brands):
            print("   ", b)


if __name__ == "__main__":
    double_clicked = len(sys.argv) == 1
    code = 0
    try:
        main()
    except Exception as e:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        print(f"\nERROR: {e}")
        code = 1
    finally:
        if double_clicked:          # keep the console window open
            try:
                input("\n完成，按回车键关闭。 Finished - press Enter to close.")
            except EOFError:
                pass
    sys.exit(code)
