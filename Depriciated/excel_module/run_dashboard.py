from commands import Session

s = Session()
s.load("data/Aun_Excel_Final.xlsx")
print("COLUMNS:", list(s.df.columns))
s.plan_and_create_dashboard(
    template_path="templates/dashboard.xlsx",
    output_path="out/candidate_dashboard.xlsx",
    metric="Expected Salary (â‚¹)",
    date_col="Interview Date",
    category_col="Degree",
    visible=True
)

