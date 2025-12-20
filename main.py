import pandas as pd

# load excel files

df_UBR = pd.read_csv("MI_Audit_UBR.csv")
df_CVS = pd.read_csv("MI_Audit_CVS.csv")

# Initialize audit columns

df_UBR["Check_Course_Currency"] = ""
df_UBR["Check_Course_Amount"] = ""
df_UBR["Check_Course_Term"] = ""

# merge UBR with CVS on course name

df_merged = df_UBR.merge(
    df_CVS[["Course_Name", "Currency", "Amount", "Term"]],
    on="Course_Name",
    how="left",
    suffixes=("","_CVS")
)

# vectorized comparison

df_merged["Check_Course_Currency"] = df_merged["Currency"] == df_merged["Currency_CVS"]
df_merged["Check_Course_Amount"] = df_merged["Amount"] == df_merged["Amount_CVS"]
df_merged["Check_Course_Term"] = df_merged["Term"] == df_merged["Term_CVS"]

# convert true/false to match/mismatch

df_merged["Check_Course_Currency"] = df_merged["Check_Course_Currency"].map({True: "Match", False: "Mismatch"})
df_merged["Check_Course_Amount"] = df_merged["Check_Course_Amount"].map({True: "Match", False: "Mismatch"})
df_merged["Check_Course_Term"] = df_merged["Check_Course_Term"].map({True: "Match", False: "Mismatch"})

# Save updated results

df_merged.to_csv("MI_Audit_UBR_audited.csv", index=False)

# Summary detailed report

summary_currency = df_merged["Check_Course_Currency"].value_counts()
summary_amount = df_merged["Check_Course_Amount"].value_counts()
summary_term = df_merged["Check_Course_Term"].value_counts()

print("Audits Completed. Results saved as MI_Audit_UBR_audited.csv")
print("\n---- Summary Report ----")
print("Currency Check: ")
print(summary_amount)
print("\nAmount Check: ")
print(summary_amount)
print("Term Check")
print(summary_term)