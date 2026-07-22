import openpyxl, json
wb = openpyxl.load_workbook("/home/claude/cntb/model.xlsx")
pts = wb["Price Target Summary"]

# Sensitivity rows were built at rows 24-28 (Driver, Downside flex, Upside flex, Notes)
new_vals = {
 24: (-1.13, 1.13),   # Net price
 25: (-1.13, 1.13),   # Peak penetration
 26: (-0.75, 0.94),   # WACC
 27: (-0.65, 0.65),   # PoS COPD
 28: (-0.35, 0.35),   # PoS Asthma Ph2
}
for row, (dn, up) in new_vals.items():
    pts[f"C{row}"] = dn
    pts[f"D{row}"] = up

# Update the closing caveat paragraph with the corrected Low/High spread
for row in range(29, 33):
    v = pts[f"B{row}"].value
    if v and "compounding uncertainty" in str(v):
        pts[f"B{row}"] = ("Read this as: holding everything else at Base, flexing ONE input by the stated amount moves rNPV/share by the amount shown (recomputed after the Q1'26 10-Q correction to fully-diluted shares and cash). Because the drivers interact "
                          "(the Low/High columns on the Cover tab move several at once), the full Low-to-High spread on the rNPV tab is much wider than the sum of these one-at-a-time moves - "
                          "that gap is itself the honest message: this stock's value is dominated by compounding uncertainty, not any single number.")
        break

wb.save("/home/claude/cntb/model.xlsx")
print("done")
