import openpyxl, json
from helpers import (set_cell, BLUE, BLACK, GREEN, BOLD_BLACK, NOTE, RED, FILL_INPUT, FILL_LIGHT,
                      CUR, CUR2, PCT, NUM, NUM1, TITLE, SUBTITLE, section_bar)

row_map = json.load(open("/home/claude/cntb/row_map.json"))
wb = openpyxl.load_workbook("/home/claude/cntb/model.xlsx")
bs = wb["Balance Sheet & Cash Runway"]

# Precision updates from the actual 10-Q
bs["C22"] = 62.7117
bs["G22"] = "CORRECTED to exact 10-Q figure: 62,711,690 shares as of 31-Mar-2026."
bs["C23"] = 56.545
bs["G23"] = "CORRECTED to exact 10-Q figure: weighted-avg basic & diluted shares, three months ended 31-Mar-2026 = 56,545,000."
bs["C28"] = 15.030
bs["C29"] = 4.746
bs["C30"] = 0.169
bs["C31"] = 19.398
bs["G31"] = "Exact 10-Q figure: net loss $19,398 thousand = $19.398mm ($0.34/share), vs. $10,272 thousand ($0.19/share) in Q1 2025. Includes $48K income tax expense."

# --- Related-party correction: James Huang / Panacea Venture ---
bs["B18"] = "Panacea Venture (largest shareholder) PIPE lead, 31-Mar-2026 ($mm)"
bs["C18"] = 4.0
bs["G18"] = ("CORRECTED / IMPORTANT: per 10-Q Note 9 (Shareholders' Equity), this $4.0mm was NOT a generic director PIPE participation - it was purchased by Panacea Venture, described in the filing as 'the Company's largest current investor,' which LED the entire March 2026 private placement. "
    "Board member James Huang is the sole owner of Panacea Innovation Limited, which is the sole owner of Panacea Venture. This is therefore a related-party transaction: the company's largest shareholder (controlled by a sitting director) anchored its own financing round. "
    "Read this as a genuinely stronger commitment-of-capital signal than ordinary PIPE participation (the largest holder underwriting the raise), but also flag the governance question a related-party-led financing always raises (were the $3.25/share terms arm's-length; how was the transaction approved). "
    "We do not have Panacea Venture's TOTAL stake size (beyond this incremental $4.0mm) from this 10-Q - that would need a Schedule 13D/G filing, which we recommend pulling next, since Panacea Venture is arguably the single most important ownership fact in the file and was not captured in our earlier institutional-holder research.")

wb.save("/home/claude/cntb/model.xlsx")
print("done")
