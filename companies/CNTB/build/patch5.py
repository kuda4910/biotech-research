import openpyxl, json
from helpers import (set_cell, BLUE, BLACK, GREEN, BOLD_BLACK, NOTE, RED, FILL_INPUT, FILL_LIGHT,
                      SECTION, CUR, CUR2, PCT, NUM, NUM1, section_bar)

row_map = json.load(open("/home/claude/cntb/row_map.json"))
wb = openpyxl.load_workbook("/home/claude/cntb/model.xlsx")
ca = wb["Cover & Assumptions"]

# Update the China-related program status note (row range 19-25ish) with the fuller, sourced deal economics
# Find the "China (Simcere license):" row
china_row = None
for r in range(19, 30):
    v = ca[f"B{r}"].value
    if v and "China" in str(v):
        china_row = r
        break

if china_row:
    ca[f"C{china_row}"] = ("Simcere holds exclusive Greater-China rights to rademikibart, all indications (signed 21-Nov-2023). Deal economics per the actual Q1'26 10-Q: $21mm upfront (2023); "
        "total milestones originally up to $123mm; $5mm received in 2024; ~$8mm LAPSED in 2025 (a time-based milestone Simcere missed the deadline for - a real, already-materialized execution-risk event, not a hypothetical); "
        "$110mm remaining eligible milestones; tiered low-double-digit royalties on Greater China net sales for ~12 years post-commercialization. China Ph3 AD trial met all endpoints; NDA submitted to China's regulator.")
    ca[f"G{china_row}"] = "UPDATED with exact figures from cntb-20260331.html (10-Q Note 8, License and Collaboration Agreement) - materially more precise and complete than our original secondary-source summary, and importantly reveals the $8mm lapsed milestone as a real precedent for Simcere-side execution risk."
    ca.row_dimensions[china_row].height = 70

# Revise China milestone realization assumption downward slightly given the now-documented lapsed milestone
A = row_map["ASSUMP_START"]
CHINA_MS_PCT = A + 14  # row 56
ca[f"C{CHINA_MS_PCT}"] = 0.15
ca[f"D{CHINA_MS_PCT}"] = 0.30
ca[f"E{CHINA_MS_PCT}"] = 0.45
ca[f"G{CHINA_MS_PCT}"] = ("REVISED DOWN slightly from our original 20/35/50% placeholder now that the 10-Q confirms one $8mm time-based milestone already lapsed in 2025 (missed by Simcere) out of the original $123mm pool - real evidence that milestone timing/achievement is not automatic. "
    "ESTIMATE of probability/timing-weighted realization of the REMAINING ~$110mm, given AD NDA is filed but commercial/regulatory milestones remain contingent on Simcere's execution, which CNTB does not control.")

wb.save("/home/claude/cntb/model.xlsx")
print("china_row:", china_row)
