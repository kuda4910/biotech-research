import openpyxl, json
from helpers import (set_cell, BLUE, BLACK, GREEN, BOLD_BLACK, NOTE, RED, FILL_INPUT, FILL_LIGHT,
                      CUR, CUR2, PCT, NUM, NUM1)

row_map = json.load(open("/home/claude/cntb/row_map.json"))
wb = openpyxl.load_workbook("/home/claude/cntb/model.xlsx")

bs = wb["Balance Sheet & Cash Runway"]
ca = wb["Cover & Assumptions"]

CASH_EQ_ROW = 7
ST_INV_ROW = 8
bs[f"C{CASH_EQ_ROW}"] = 46.034
bs[f"G{CASH_EQ_ROW}"] = "CORRECTED from actual 10-Q balance sheet (period 31-Mar-2026): cash and cash equivalents = $46.034mm. Prior version of this file wrongly split a secondary source's headline '$46.0mm' figure between this line and the one below - the $46.0mm figure in press summaries is actually just this cash-and-equivalents line, not the total."
bs[f"C{ST_INV_ROW}"] = 5.997
bs[f"B{ST_INV_ROW}"] = "Short-term investments"
bs[f"G{ST_INV_ROW}"] = "CORRECTED from actual 10-Q balance sheet: short-term investments = $5.997mm, additive to cash and equivalents above (total liquidity $52.03mm, not $46.0mm as press summaries implied)."

TOTAL_CASH_ROW = row_map["BS_TOTAL_CASH_ROW"]
bs[f"G{TOTAL_CASH_ROW}"] = "Now ties to the actual 10-Q balance sheet total: $46.034mm + $5.997mm = $52.03mm as of 31-Mar-2026 (source: cntb-20260331.html, Condensed Consolidated Balance Sheets)."

bs[f"G{row_map['BS_RUNWAY_Q_ROW']+1}"] = ("UPDATED: the 10-Q's formal Liquidity & Going Concern note (Note 2) only commits to cash being sufficient for 'at least one year from the date this Quarterly Report is filed' - filed 12-May-2026, i.e. through at least ~12-May-2027. "
    "The looser 'into 2H 2027' framing comes from press-release/8-K language, not this more conservative formal disclosure. Treat ~May-2027 as the company's actual legal commitment and '2H 2027' as aspirational guidance on top of it.")

SHARES_ROW = row_map["SHARES_ROW"]
ca[f"C{SHARES_ROW}"] = 62.7117
ca[f"G{SHARES_ROW}"] = "CORRECTED to the exact figure on the 10-Q balance sheet: 62,711,690 ordinary shares issued and outstanding as of 31-Mar-2026 (vs. 56,442,308 at 31-Dec-2025). Source: cntb-20260331.html."

FD_ROW = row_map["FD_ROW"]
ca[f"C{FD_ROW}"] = 78.48
ca[f"G{FD_ROW}"] = ("RESOLVED from the actual 10-Q (Note on Net Loss per Share, anti-dilutive securities table): 15,167,637 stock options outstanding (weighted-avg exercise price $2.08 - meaningfully in-the-money at the current ~$2.40 share price) "
    "plus 600,000 employee stock purchase plan rights, on top of 62.7117mm basic shares = ~78.48mm on a simple if-exercised gross basis (62.7117 + 15.1676 + 0.60). "
    "Note this is a gross, not treasury-method-reduced, fully-diluted count - the treasury-stock-method diluted EPS count would be lower because option-exercise proceeds would notionally repurchase some shares, but for valuation (not EPS) purposes the gross if-exercised count is the more standard convention. Source: cntb-20260331.html, Note on Net Loss per Share.")
ca[f"C{FD_ROW}"].font = BLUE

sc_label_row = None
for rr in range(20, 40):
    v = bs[f"B{rr}"].value
    if v and "options / RSUs outstanding" in str(v):
        sc_label_row = rr
        break

if sc_label_row:
    bs[f"B{sc_label_row}"] = "Stock options outstanding (mm), wtd-avg strike $2.08"
    bs[f"C{sc_label_row}"] = 15.1676
    bs[f"G{sc_label_row}"] = ("SOURCED (was previously an unknown placeholder). Per 10-Q Note 10 (Equity Incentive Plans): 15,167,637 options outstanding at 31-Mar-2026, weighted-average exercise price $2.08 - i.e. in-the-money at the current share price, real near-term dilution risk, "
        "not a distant hypothetical. Separately, 600,000 employee stock purchase plan rights are also outstanding (anti-dilutive table). Source: cntb-20260331.html.")
    bs[f"C{sc_label_row}"].font = BLUE

wb.save("/home/claude/cntb/model.xlsx")
print("patched share count row:", sc_label_row)
