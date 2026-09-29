# Brand-city sales tool

Fills monthly brand-by-city car sales into a set of Excel workbooks and computes the share of each brand-city market taken by a given set of placements.

The live sales client and real data are not included (see [What's not here](#whats-not-here)). The demo files are synthetic.

## How it works

1. **`build_panel.py`** turns a contract book (one row per contract: brand, branch, funding date) into one workbook per brand, with one row per branch city and month and the contract count in column E. Branch names are cleaned first: the company prefix and the 分公司 suffix are stripped, and each branch is matched to its province through `branches.csv`.
2. **`crawler.py`** fills column F with the brand's sales in that city and month, and writes column G as the live formula `=E/F`.
   - Numbered branches (广州第一) are collapsed to the city, and branches named after a region (四川) are read as its capital.
   - One request per (month, city) returns every brand, so responses are cached and each is fetched once per run.
   - The source lists HarmonyOS marques (问界, 智界, 享界, 尚界) by model only, so their models are summed.
   - Months the source has not published yet stay blank rather than 0, and a later run fills them in.
   - `--extend` adds rows for every month up to the current one before filling.
3. **Distribution.** Users never install anything.
   - **Windows:** a `.exe` that GitHub Actions builds on every push, runs on the demo workbooks, and publishes as a download.
   - **Mac:** `crawler_mac.py` does the same fill with only the Python that ships with macOS. It swaps openpyxl for direct edits to the worksheet XML inside each `.xlsx` (zipfile + ElementTree), resolving shared strings and leaving styles untouched. `运行_RUN_mac.command` runs it on a double-click. It shares the mapping and sources with `crawler.py`, and both give identical output on the demo workbooks.
   - `使用说明_HOWTO.txt` is the bilingual user guide.

## Try it

```bash
pip install -r requirements.txt
python crawler.py --dir workbooks
```

This fills the three demo workbooks in place. On a Mac with nothing installed, run `python3 crawler_mac.py` instead, or double-click `运行_RUN_mac.command`. For the Windows build, open **Actions → Build Windows EXE → latest run → Artifacts**; the download holds the `.exe` and the demo workbooks side by side.

To regenerate the demo data from scratch:

```bash
python make_demo_data.py
python build_panel.py --contracts sample_data/contracts_demo.xlsx --out workbooks \
    --strip-prefix 示例汽车服务有限公司
```

Sample rows after a run (synthetic figures):

| 品牌名称 | 展业省份 | 展业城市（分公司） | 投放月份 | 投放数（台） | 品牌城市月销量（台） | 月销量市占比 |
|---|---|---|---|---|---|---|
| 比亚迪 | 广东省 | 广州第一 | 25年1月 | 13 | 736 | 1.77% |
| 比亚迪 | 广东省 | 广州第一 | 25年2月 | 9 | 307 | 2.93% |
| AITO 问界 | 四川省 | 四川 | 25年1月 | 7 | 2,470 | 0.28% |

## Sales sources

| `--source` | What it does |
|---|---|
| `demo` (default) | Deterministic synthetic figures, no network. The current month counts as unpublished. |
| `http` | `POST $SALES_API_URL` with `{"date": "YYYY-MM", "area": "省:市", "type": 0}`, where `type` 0 is a brand ranking and 1 a model ranking. Expects `{"data": [{"name": ..., "sales": "1,234"}]}` back. `$SALES_API_TOKEN`, if set, is sent as a bearer token. |

`http` is used automatically whenever `SALES_API_URL` is set.

## What's not here

- **Real data.** `sample_data/` and `workbooks/` are random numbers in the real layout.
- **The live client.** The tool was built against a third-party sales-ranking API that signs every request with an encrypted token. That client is withheld; `--source http` takes its place for any endpoint that follows the contract above.
- **The branch list.** `branches.csv` holds the six demo branches only.

## Files

| File | Purpose |
|---|---|
| `crawler.py` | Fills F and G in every workbook in a folder |
| `crawler_mac.py` | The same fill with no dependencies, for Macs |
| `运行_RUN_mac.command` | Double-click launcher for the Mac version |
| `build_panel.py` | Contract book → one workbook per brand |
| `branches.csv` | Branch → province lookup |
| `make_demo_data.py` | Writes the synthetic contract book |
| `sample_data/contracts_demo.xlsx` | Synthetic contract book |
| `workbooks/` | Demo brand workbooks, unfilled |
| `build_windows.bat` | Local Windows build with PyInstaller |
| `.github/workflows/build-windows.yml` | Cloud build, smoke test and download |
| `icon.ico` | Program icon |
| `使用说明_HOWTO.txt` | User guide (Chinese and English) |
