#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
品牌城市月销量 抓取工具 — Mac 零依赖版
The same fill as crawler.py, using only the Python 3 that ships with macOS.

openpyxl is replaced by editing the worksheet XML inside each .xlsx directly
(zipfile + ElementTree): shared strings are resolved, F and G are set cell by
cell, and every other part of the file is copied through untouched, so styles,
widths and the percent format survive. Brand mapping, area rules and sales
sources come from crawler.py, so both versions always give the same numbers.

Double-click 运行_RUN_mac.command, or:
  python3 crawler_mac.py                 # the workbooks next to this file
  python3 crawler_mac.py --dir DIR
  python3 crawler_mac.py --dry-run

Differences from crawler.py: no --extend (add new month rows with the Windows
tool or by hand).
"""

import argparse
import io
import os
import re
import sys
import time
import zipfile
from xml.etree import ElementTree as ET

import crawler as core

M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"


def q(tag):
    return f"{{{M}}}{tag}"


def col_letters(ref):
    return re.match(r"[A-Z]+", ref).group(0)


def col_number(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def register_namespaces(xml_bytes):
    """Keep the file's own namespace prefixes when the XML is written back."""
    for _, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
        ET.register_namespace(prefix, uri)


class Workbook:
    """The first worksheet of an .xlsx, read and written without openpyxl."""

    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}
        self.sheet_part = self._first_sheet_part()
        self.strings = self._shared_strings()
        raw = self.parts[self.sheet_part]
        register_namespaces(raw)
        self.root = ET.fromstring(raw)

    def _first_sheet_part(self):
        wb = ET.fromstring(self.parts["xl/workbook.xml"])
        rid = wb.find(f"{q('sheets')}/{q('sheet')}").get(f"{{{R}}}id")
        rels = ET.fromstring(self.parts["xl/_rels/workbook.xml.rels"])
        for rel in rels.iter(f"{{{PR}}}Relationship"):
            if rel.get("Id") == rid:
                target = rel.get("Target")
                return target.lstrip("/") if target.startswith("/") else "xl/" + target
        raise ValueError("workbook has no first sheet")

    def _shared_strings(self):
        raw = self.parts.get("xl/sharedStrings.xml")
        if raw is None:
            return []
        root = ET.fromstring(raw)
        return ["".join(t.text or "" for t in si.iter(q("t"))) for si in root.iter(q("si"))]

    def value(self, cell):
        if cell is None:
            return ""
        t = cell.get("t")
        if t == "inlineStr":
            return "".join(x.text or "" for x in cell.iter(q("t")))
        v = cell.find(q("v"))
        if v is None or v.text is None:
            return ""
        return self.strings[int(v.text)] if t == "s" else v.text

    def rows(self):
        """Yield (row number, row element, {column letters: cell})."""
        for row in self.root.find(q("sheetData")).iter(q("row")):
            cells = {col_letters(c.get("r")): c for c in row.findall(q("c"))}
            yield int(row.get("r")), row, cells

    @staticmethod
    def cell(row, cells, letters, rn):
        """The cell at <letters><rn>, created in column order if missing."""
        if letters in cells:
            return cells[letters]
        c = ET.Element(q("c"), {"r": f"{letters}{rn}"})
        pos = sum(1 for x in row.findall(q("c"))
                  if col_number(col_letters(x.get("r"))) < col_number(letters))
        row.insert(pos, c)
        cells[letters] = c
        return c

    @staticmethod
    def _reset(c):
        for child in list(c):
            c.remove(child)
        c.attrib.pop("t", None)

    def set_number(self, c, n):
        self._reset(c)
        ET.SubElement(c, q("v")).text = str(n)

    def set_formula(self, c, formula):
        self._reset(c)
        ET.SubElement(c, q("f")).text = formula

    def clear(self, c):
        self._reset(c)

    def save(self):
        self.parts[self.sheet_part] = ET.tostring(self.root, encoding="UTF-8",
                                                  xml_declaration=True)
        # Formulas are written without cached results: have Excel/Numbers
        # recalculate on open.
        wb = self.parts["xl/workbook.xml"].decode("utf-8")
        if "fullCalcOnLoad" not in wb:
            if "<calcPr" in wb:
                wb = wb.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
            else:
                wb = re.sub(r"(</sheets>)", r'\1<calcPr fullCalcOnLoad="1"/>', wb, count=1)
            self.parts["xl/workbook.xml"] = wb.encode("utf-8")
        tmp = self.path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in self.parts.items():
                z.writestr(name, data)
        os.replace(tmp, self.path)


def fill(path, records_for, dry_run):
    wb = Workbook(path)
    rows = wb.rows()
    _, _, header = next(rows)
    cols = {wb.value(c).strip(): letters for letters, c in header.items()}
    need = [core.HDR_BRAND, core.HDR_PROV, core.HDR_CITY, core.HDR_MONTH, core.HDR_TARGET]
    if any(h not in cols for h in need):
        return None
    e_col, f_col, g_col = cols.get(core.HDR_PLACED), cols[core.HDR_TARGET], cols.get(core.HDR_SHARE)

    filled, unknown = 0, set()
    for rn, row, cells in rows:
        brand, prov, city, month = (wb.value(cells.get(cols[h])).strip()
                                    for h in need[:4])
        if not (brand and prov and city and month):
            continue
        api_name = core.BRAND_MAP.get(brand)
        if not api_name:
            unknown.add(brand)
            continue
        dtype = 1 if api_name.startswith("SUB:") else 0
        recs = records_for(core.month_to_date(month), core.city_to_area(prov, city), dtype)
        if not recs:
            continue                              # not published yet: leave blank
        val = core.brand_sales_from(recs, api_name)
        wb.set_number(wb.cell(row, cells, f_col, rn), val)
        filled += 1
        if g_col and e_col:
            g = wb.cell(row, cells, g_col, rn)
            if val:
                wb.set_formula(g, f"{e_col}{rn}/{f_col}{rn}")
            else:
                wb.clear(g)
    if not dry_run:
        wb.save()
    return filled, unknown


def main():
    core.safe_console()
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=None)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--source", choices=("demo", "http"), default=None)
    args = ap.parse_args()

    if args.dir is None:
        args.dir = os.path.dirname(os.path.abspath(__file__))
        if not core.find_workbooks(args.dir) and os.path.isdir(
                os.path.join(args.dir, "workbooks")):
            args.dir = os.path.join(args.dir, "workbooks")
    if args.source is None:
        args.source = "http" if os.environ.get("SALES_API_URL") else "demo"
    source = core.make_source(args.source)

    files = core.find_workbooks(args.dir)
    if not files:
        sys.exit(f"No .xlsx files in {args.dir}")
    print(f"Source: {source.name}")
    print(f"Found {len(files)} workbooks in {args.dir}")

    cache = {}

    def records_for(date, area, dtype):
        k = (date, area, dtype)
        if k not in cache:
            cache[k] = source.records(date, area, dtype)
            if args.source == "http":
                time.sleep(0.2)
        return cache[k]

    total, unknown = 0, set()
    for f in files:
        res = fill(f, records_for, args.dry_run)
        name = os.path.basename(f)
        if res is None:
            print(f"  ! skipping {name} (missing columns)")
            continue
        n, unk = res
        total += n
        unknown |= unk
        print(f"  {name}: filled {n} rows" + ("  (dry run, not saved)" if args.dry_run else ""))

    print(f"\nDone. {total} rows filled, {len(cache)} unique requests.")
    if unknown:
        print("Unmapped brands (add them to BRAND_MAP in crawler.py):")
        for b in sorted(unknown):
            print("   ", b)


if __name__ == "__main__":
    main()
