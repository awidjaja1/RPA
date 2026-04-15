import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles.numbers import BUILTIN_FORMATS
from datetime import datetime


def ocfTRA(params):
    try:
        emplid = params[0]
        file_path = params[2]
        #======================================================
        # Step 1. Read the CSV or Excel input
        #======================================================
        #file_path = r"C:\Users\awidjaja\Documents\RFL- POC\Input_Files\Sample 1_Transaction Register Analysis - 2025-10-21T094422.255.csv"
        df = pd.read_csv(file_path)

        #======================================================
        # Step 2. Select & rename columns (mapping VBA intent)
        #======================================================
        # Adjust mapping to match your real headers
        df = df[[
            "Transaction Line Description",
            "Transaction Date",
            "Transaction Number",
            "Bill-to Customer Name",
            "Bill-to Customer Account Number",
            "Coverage Month",
            "Line Amount",
            "Previous Transaction Number"
        ]]

        df.columns = [
            "Description", "Invoice Date", "Invoice #", "Customer Name",
            "Customer Account #", "Coverage Month", "Amount", "Prev Transaction #"
        ]

        #======================================================
        # Step 3. Remove zero Amount rows (VBA V2.1)
        #======================================================
        df = df[df["Amount"] != 0]

        #======================================================
        # Step 4. Sort by Amount (then by Prev Transaction #)
        #======================================================
        df.sort_values(by=["Amount", "Prev Transaction #"], ascending=[True, True], inplace=True)

        #======================================================
        # Step 5. Create Subtotals grouped by Prev Transaction #
        #======================================================
        subtotals = (
            df.groupby("Prev Transaction #", dropna=False)["Amount"]
            .sum()
            .reset_index()
            .rename(columns={"Amount": "Subtotal"})
        )

        # Insert subtotal rows at the bottom of each group
        subtotal_rows = []
        for _, row in subtotals.iterrows():
            subtotal_rows.append({
                "Description": f"Total for {row['Prev Transaction #']}",
                "Invoice Date": "",
                "Invoice #": "",
                "Customer Name": "",
                "Customer Account #": "",
                "Coverage Month": "",
                "Amount": row["Subtotal"],
                "Prev Transaction #": row["Prev Transaction #"]
            })
        df_out = pd.concat([df, pd.DataFrame(subtotal_rows)], ignore_index=True)

        #======================================================
        # Step 6. Write to Excel
        #======================================================
        todaydate = datetime.now().strftime("%m-%d-%Y_%H-%M-%S")
        output_path = file_path
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")

        #======================================================
        # Step 7. Formatting in openpyxl
        #======================================================
        wb = load_workbook(output_path)
        ws = wb.active

        header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        header_font = Font(bold=True)
        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, bottom=thin)

        # Header formatting
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border

        # Bold subtotal rows (those beginning with “Total”)
        for row in ws.iter_rows(min_row=2):
            cell = row[0]
            if isinstance(cell.value, str) and cell.value.startswith("Total"):
                for c in row:
                    c.font = Font(bold=True)
                    c.border = border

        # Auto column width
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2

        #======================================================
        # Step 8. Append “Pmts / Applied / Balance” block
        #======================================================
        last_row = ws.max_row + 2
        ws[f"F{last_row}"] = "Pmts"
        ws[f"F{last_row+1}"] = "Applied"
        ws[f"F{last_row+2}"] = "Balance"
        ws[f"G{last_row}"] = 0
        ws[f"G{last_row+1}"] = f"=G{last_row-3}"  # mimic VBA =R[-3]C
        ws[f"G{last_row+2}"] = f"=G{last_row}-G{last_row+1}"

        for r in ws[f"F{last_row}:G{last_row+2}"]:
            for c in r:
                c.fill = header_fill
                c.font = header_font
                c.border = border

        ws.auto_filter.ref = ws.dimensions
        wb.save(output_path)
        return "Success"
    except Exception as E:
        return f"{type(E).__name__}: {E}"

#=======================================
# Function to Process OCF Payments
#=======================================
def ocfPmts(params):
    try:
        #===================================
        # Step 1. Load the workbook and sheet
        #===================================
        file_path = params[3]
        wb = load_workbook(file_path)
        ws = wb['Exported']
        #===================================
        # Step 2. Load data into pandas
        #===================================
        data = pd.read_excel(file_path, sheet_name='Exported')

        #===================================
        # Step 3. Delete column A (index 0)
        #===================================
        data.drop(data.columns[0], axis=1, inplace=True)

        #===================================
        # Step 4. Delete rows where column B has "Credit Memo" or "Invoice"
        #===================================
        filtered_data = data[~data.iloc[:, 1].isin(["Credit Memo", "Invoice"])].copy()

        #===================================
        # Step 5. Sort by column C (index 2)
        #===================================
        filtered_data.sort_values(by=filtered_data.columns[2], inplace=True)

        #===================================
        # Step 6. Column E (index 4) = -1 * column D (index 3)
        #===================================
        filtered_data.iloc[:, 4] = filtered_data.iloc[:, 3] * -1

        #===================================
        # Step 7. Format columns D to F (as comma style - done in Excel styling below)
        #===================================
        # Replace data in sheet
        for row in ws['A2': f'{chr(65 + filtered_data.shape[1] - 1)}{ws.max_row}']:
            for cell in row:
                cell.value = None

        for r_idx, row in enumerate(dataframe_to_rows(filtered_data, index=False, header=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        #===================================
        # Step 8. Set column widths
        #===================================
        ws.column_dimensions['F'].width = 6.89
        ws.column_dimensions['G'].width = 7.89

        #===================================
        # Step 9. Apply conditional formatting to column H for "Reversed"
        #===================================

        highlight_fill = PatternFill(start_color='808080', end_color='808080', fill_type='solid')
        ws.conditional_formatting.add(
            'H2:H1000',
            FormulaRule(formula=['ISNUMBER(SEARCH("Reversed",H2))'], fill=highlight_fill)
        )

        #===================================
        # Step 10. Auto-fit column F (limited support)
        # OpenPyXL can't truly auto-fit. You could approximate if needed.
        #===================================
        # Save result
        output_path = f"C:\ProgramData\AutomationAnywhere\Bots\Logs\RFL-UCPath\ProcessLogs\OCF-Payments\{emplid}_{todaydate}_OCF_Payments.xlsx"
        wb.save(output_path)
        return "Success"
    except Exception as E:
        return f"{type(E).__name__}: {E}"

