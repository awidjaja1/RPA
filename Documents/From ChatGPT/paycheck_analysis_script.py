
import pandas as pd

# Step 1: Load Excel file
df = pd.read_excel("paycheck_report.xlsx", engine='openpyxl', parse_dates=["Paycheck Issue Date"])

# Step 2: Detect missing paycheck dates (assuming biweekly frequency)
full_date_range = pd.date_range(start=df["Paycheck Issue Date"].min(),
                                end=df["Paycheck Issue Date"].max(),
                                freq='14D')

existing_dates = pd.to_datetime(df["Paycheck Issue Date"].unique())
missing_dates = full_date_range.difference(existing_dates)

# Step 3: Create placeholder rows for missing dates
missing_rows = pd.DataFrame({
    "Paycheck Issue Date": missing_dates,
    "Is_Missing": True
})

# Add flag to original data
df["Is_Missing"] = False

# Step 4: Combine and sort data
combined_df = pd.concat([df, missing_rows], ignore_index=True).sort_values("Paycheck Issue Date")

# Step 5: Export to Excel with background color formatting for missing rows
output_path = "processed_paycheck_report.xlsx"

with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    combined_df.to_excel(writer, sheet_name='Data', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['Data']

    # Define formatting
    highlight_format = workbook.add_format({'bg_color': '#FFCCCC'})  # Light red
    
    for row_num, is_missing in enumerate(combined_df['Is_Missing'], start=1):
        if is_missing:
            worksheet.set_row(row_num, cell_format=highlight_format)
    
    # Optional: create a pivot-style summary
    if "Employee" in df.columns and "Plan Type" in df.columns and "Deduction Amount" in df.columns:
        pivot_df = combined_df[~combined_df["Is_Missing"]].pivot_table(
            index="Employee",
            columns="Plan Type",
            values="Deduction Amount",
            aggfunc="sum",
            fill_value=0
        )
        pivot_df.to_excel(writer, sheet_name="Pivot Summary")
