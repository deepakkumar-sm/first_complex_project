import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# load excel files
df_UBR = pd.read_excel("Excel_Audit_Project.xlsx", sheet_name="Sheet1")
df_CVS = pd.read_excel("Excel_Audit_Project.xlsx", sheet_name="Sheet2")

print(df_UBR)
print() 
print(df_CVS)


# Initialize audit columns

df_UBR["Check_Currency_UBR"] = ""
df_UBR["Check_Amount_UBR"] = ""
df_UBR["Check_Course_Term_UBR"] = ""
df_UBR["Check_vendor_name_UBR"] = ""
# df_UBR["Check_deliverydate_conflict_UBR"] = ""
df_UBR["MI_Eligibility"] = ""


# merge UBR with CVS on course name

df_merged = df_UBR.merge(
    df_CVS[["Course_Name","Currency","Amount","Course_Term","vendor_name"]],
    on="Course_Name",
    how="left",
    suffixes=("","_CVS")
)

# helper function for comparison

def compare_with_cvs(row, col, col_cvs):
    if pd.isna(row[col_cvs]):
        return "Course Not Found"
    return "Match" if row[col] == row[col_cvs] else "Mismatch"

# vectorized comparison
df_merged["Check_Currency_UBR"] = df_merged.apply(lambda r: compare_with_cvs(r, "Currency", "Currency_CVS"), axis=1)
df_merged["Check_Amount_UBR"] = df_merged.apply(lambda r: compare_with_cvs(r, "Amount", "Amount_CVS"), axis=1)
df_merged["Check_Course_Term_UBR"] = df_merged.apply(lambda r: compare_with_cvs(r, "Course_Term", "Course_Term_CVS"), axis=1)
df_merged["Check_vendor_name_UBR"] = df_merged.apply(lambda r: compare_with_cvs(r, "vendor_name", "vendor_name_CVS"), axis=1)

df_merged.loc[df_merged["POD_Recd"] == "No", "MI_Eligibility"] = "Not Eligible"
df_merged.loc[df_merged["POD_Recd"] == "Yes", "MI_Eligibility"] = "Eligible"



# convert true/false to match/mismatch
# df_merged["Check_Currency_UBR"] = df_merged["Check_Currency_UBR"].map({True: "Match", False: "Mismatch"})
# df_merged["Check_Amount_UBR"] = df_merged["Check_Amount_UBR"].map({True: "Match", False: "Mismatch"})
# df_merged["Check_Course_Term_UBR"] = df_merged["Check_Course_Term_UBR"].map({True: "Match", False: "Mismatch"})
# df_merged["Check_vendor_name_UBR"] = df_merged["Check_vendor_name_UBR"].map({True: "Match", False: "Mismatch"})

# save to new excel file
df_merged.to_excel("MI_Audit_UBR_audited.xlsx", index=False)

# Summary detailed report

wb = load_workbook(filename='MI_Audit_UBR_audited.xlsx')
sheet = wb.active

red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

for row in sheet.iter_rows(min_row=2, min_col=12, max_col=16):

    for cell in row:
        if cell.value == "Mismatch":
            cell.fill = red_fill
        elif cell.value == "Course Not Found":
            cell.fill = yellow_fill


wb.save("MI_Audit_UBR_audited.xlsx")
wb.close()



