import pandas as pd
import numpy as np
from datetime import datetime
import xlsxwriter
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Color

def insertMissingDates(group,Calendardf):
    Calendar = pd.Series(Calendardf[(Calendardf >=group['Paycheck Issue Date'].min())&(Calendardf<=group['Paycheck Issue Date'].max())])
    
    missingDates = sorted(list(set(Calendar)-set(group['Paycheck Issue Date'])))
    newRows =[]
    #insert the missing dates:
    if len(missingDates)>0:
        for missedDate in missingDates:
            #print(group[['Business Unit','Plan Type']].iloc[0])
            #print(group['Business Unit'].loc[0])
            #======================================================
            # added 10/27/2025 at 8:45 am
            missedDateMonth = missedDate.month
            missedDateYear = missedDate.year
            #check the year of the known value where paycheck issue date year in front and back
            frontKnownExistedRowYear = group['Paycheck Issue Date'][(group['Paycheck Issue Date'])< missedDate].iloc[-1].year

            #if December (or 12), see the row after i.e. move forward
            if (missedDateMonth == 12):
                 # Get the known date after the missed date. the known value is at the top of the array, so iloc[0] should get the value on top of array  
                existedRowDate = group['Paycheck Issue Date'][(group['Paycheck Issue Date'] >missedDate)].iloc[0] 
            elif(missedDateYear == frontKnownExistedRowYear):
                # Get the known date after the missed date. the known value is at the bottom of the array, so iloc[-1] should get the value on bottom of array  
                existedRowDate = group['Paycheck Issue Date'][(group['Paycheck Issue Date'] <missedDate)].iloc[-1]
            else:
                # Get the known date after the missed date. the known value is at the top of the array, so iloc[0] should get the value on top of array  
                existedRowDate = group['Paycheck Issue Date'][(group['Paycheck Issue Date'] >missedDate)].iloc[0] 
            
            prior_row = group[(group["Paycheck Issue Date"] == existedRowDate)]
            #print(prior_row)
            if not prior_row.empty:
                if len(prior_row)>1:
                    for indexNum in range(len(prior_row)):
                        pr = prior_row.iloc[indexNum]
                        new_row ={
                            'Business Unit':pr['Business Unit'],
                            'Employee ID':pr['Employee ID'],
                            'Employee Record':pr['Employee Record'],
                            'Employee Name':pr['Employee Name'],
                            'Pay Group':pr['Pay Group'],
                            'Pay Period End Date':pr['Pay Period End Date'],
                            'Paycheck Issue Date':missedDate,
                            'Employee Type':pr['Employee Type'],
                            'Plan Type':pr['Plan Type'],
                            'Benefit Plan':pr['Benefit Plan'],
                            'Deduction Code':pr['Deduction Code'],
                            'Deduction Code Descr':pr['Deduction Code Descr'],
                            'Deduction Classification':pr['Deduction Classification'],
                            'Current Deduction':pr['Current Deduction'],
                            'Payback Amount':pr['Payback Amount'],
                            'Refund Amount':pr['Refund Amount'],
                            'Not Taken':pr['Not Taken'],
                            'Reason':pr['Reason'],
                            'Calculated Base':pr['Calculated Base'],
                            'Paycheck Status':pr['Paycheck Status'],
                            'Paycheck Option':pr['Paycheck Option'],
                            'Off Cycle':pr['Off Cycle'],
                            'Is New Row':True
                        }
                        newRows.append(new_row)
                else:
                    prIloc0 = prior_row.iloc[0]
                    new_row2 ={
                        'Business Unit':prIloc0['Business Unit'],
                        'Employee ID':prIloc0['Employee ID'],
                        'Employee Record':prIloc0['Employee Record'],
                        'Employee Name':prIloc0['Employee Name'],
                        'Pay Group':prIloc0['Pay Group'],
                        'Pay Period End Date':prIloc0['Pay Period End Date'],
                        'Paycheck Issue Date':missedDate,
                        'Employee Type':prIloc0['Employee Type'],
                        'Plan Type':prIloc0['Plan Type'],
                        'Benefit Plan':prIloc0['Benefit Plan'],
                        'Deduction Code':prIloc0['Deduction Code'],
                        'Deduction Code Descr':prIloc0['Deduction Code Descr'],
                        'Deduction Classification':prIloc0['Deduction Classification'],
                        'Current Deduction':prIloc0['Current Deduction'],
                        'Payback Amount':prIloc0['Payback Amount'],
                        'Refund Amount':prIloc0['Refund Amount'],
                        'Not Taken':prIloc0['Not Taken'],
                        'Reason':prIloc0['Reason'],
                        'Calculated Base':prIloc0['Calculated Base'],
                        'Paycheck Status':prIloc0['Paycheck Status'],
                        'Paycheck Option':prIloc0['Paycheck Option'],
                        'Off Cycle':prIloc0['Off Cycle'],
                        'Is New Row':True
                    }
                    newRows.append(new_row2)
    return newRows

def highlightAnomalies(val, col_name):
    if col_name == 'Payback Amount' and (val >0 or val <0):
        return 'background-color: pink'
    elif col_name =='Refund Amount' and (val > 0 or val<0):
        return 'background-color:pink'
    elif col_name =='Not Taken' and (val > 0 or val<0):
        return 'background-color:pink'
    elif col_name == 'Reason' and val != "Null":
        return 'background-color:pink'
    elif col_name =='Paycheck Status' and val !="Confirmed":
        return 'background-color:pink'
    elif col_name =='Paycheck Option' and val !="Advice":
        return 'background-color:pink'
    elif col_name =='Off Cycle' and val !="N":
        return 'background-color:pink'
    else:
        return ''

def highlightAll(row,rules,selectedPlans):
    # step 1. create a list of styles( one for each column) starting empty
    styles = [''] * len(row)

    # step 2. create a dictionary to find the index of each column name
    col_idx = {col:idx for idx,col in enumerate(row.index)}

    # step 3. apply the rule
    # rule example:
    #       'Payback Amount': (lambda v: v!=0, ['Payback Amount'],'blue')
    #           ^                   ^               ^               ^
    #           col                 condition       target_cols     color
    #--------------------------------------------------------------------
    for col, (condition, target_cols, color) in rules.items():
        try:
            #check if the value in the column meets the condition
            if condition(row[col]):
                for column in target_cols:
                    styles[col_idx[column]]=f"background-color:{color}"

        except KeyError:
            # if column doesnt exist, skip it
            continue
    if (
            pd.notna(row.get('Deduction Variance')) and 
            row['Deduction Variance'] != 0.0 and 
            row.get('Plan Type') in selectedPlans
        ):
            if 'Current Deduction' in col_idx:
                styles[col_idx['Current Deduction']] = 'background-color: pink'

    return styles

def AddLegends(fin001outputfile):
    try: 
        filepath= fin001outputfile
        wb = openpyxl.load_workbook(filepath)
        SummaryWB= wb['Summary']
        legendRow = (SummaryWB.max_row)+5

        LightGreen = Color(rgb='C6EFCE')
        DarkGreen = Color(rgb='006100')
        LightRed = Color(rgb='FFC7CE')
        DarkRed = Color(rgb='9C0006')
        LightYellow = Color(rgb='EBEB9C')
        DarkYellow = Color(rgb='9C6500')
        DarkTeal = Color(rgb='156082')
        Orange = Color(rgb='E97132')
        DarkGreen = Color(rgb='196B24')
        Plum  = Color(rgb='A02B93')
        DarkGrey = Color(rgb='808080')
        White = Color(rgb='FFFFFF')
        
        Good = PatternFill(patternType='solid', fgColor=LightGreen)
        Bad = PatternFill(patternType='solid', fgColor=LightRed)
        Neutral = PatternFill(patternType='solid', fgColor=LightYellow)
        Accent1 = PatternFill(patternType='solid', fgColor=DarkTeal)
        Accent2 = PatternFill(patternType='solid', fgColor=Orange)
        Accent3 = PatternFill(patternType='solid', fgColor=DarkGreen)
        Accent5 = PatternFill(patternType='solid', fgColor=Plum)
        WaivedDG = PatternFill(patternType='solid', fgColor=DarkGrey)
        
        legends ={
        "Waived":WaivedDG,
        "Arrears":Accent1,
        "EE owes":Bad,
        "EE paid":Neutral,
        "UC owes":Good,
        "Refund":Accent3,
        "Union Change":Accent5,
        "Missed Paycheck":Accent2
            
        }


        legendCounter= 1
        SummaryWB[f"B{legendRow}"] = "Legends:"
        for legend in legends.keys():
            row = legendRow + legendCounter
            SummaryWB[f"C{row}"] =str(legend)
            SummaryWB[f"B{row}"].fill = legends[legend]
            legendCounter +=1
        
        # Insert EE and RFL rows 1 and 2
        SummaryWB.insert_rows(idx=0, amount=5)
        SummaryWB['A1']="EEID"
        SummaryWB['B1']= 123456
        SummaryWB['A2']="RFL"
    except Exception as e:
        return(e)
    finally:
        wb.save(filepath)
        wb.close()
        return("success")

def main(FIN001filepath,paycheckfilepath):
    #calendarpath = r'C:\ProgramData\AutomationAnywhere\Bots\Logs\RFL-UCPath\Asset\'
    try:
        #===========================================
        #Step 1. Read FINOO1 sample report and paycheck table
        #=========================================== 
        FIN001filepath = FIN001filepath
        #FIN001filepath =r"C:\Users\awidjaja\Documents\RFL- POC\FIN001\UC_FIN001_PAYCHECK_DEDUCTIONS_10473193 - to be read by python.xls"
        df = pd.read_excel(FIN001filepath, header=1)
        paycheckfilepath=paycheckfilepath
        #paycheckfilepath=r"C:\Users\awidjaja\Documents\RFL- POC\Input_Files\BiWeekly Paycheck Dates.xlsx"
        dfPaycheck = pd.read_excel(paycheckfilepath)
        payGroup = df['Pay Group'].loc[0]
        payGroupType = 'Bi-Weekly'
        if payGroupType =='Bi-Weekly':
            paycheckfilepath=r"C:\Users\awidjaja\Documents\RFL- POC\Input_Files\BiWeekly Paycheck Dates.xlsx"
        else:
            paycheckfilepath=r"C:\Users\awidjaja\Documents\RFL- POC\Input_Files\Monthly Paycheck Dates.xlsx"
        dfPaycheck = pd.read_excel(paycheckfilepath)
        #===========================================
        #Step 2. Identify and Insert Missed Paychecks
        #===========================================
        selectedPlans =[
            "Accident",
            "Basic Dependent Life",
            "Critical Illness - EE (+Ch)",
            "Critical Illness - SP/DP",
            "Employee & Dependent AD&D",
            "Exp Dependent Life - Child",
            "Exp Dependent Life - Spouse/DP",
            "Hospital Indemnity",
            "Legal Insurance",
            "Medical",
            "Supplemental Life",
            "Voluntary LongTerm Disability",
            "Voluntary ShortTerm Disability",
            "Dental",
            "Life",
            "Basic Disability",
            "Vision"
        ]
        # create 2 dataframes: Selected vs Other
        dfSelected = df[(df['Plan Type'].isin(selectedPlans))]
        dfOther =df[~(df['Plan Type'].isin(selectedPlans))]
        # iterate each selected plan and insert missed paycheck dates
        for eachPlan in dfSelected['Plan Type'].unique():
            if eachPlan in(['Dental','Life','Basic Disability','Vision']):
                lInsertedMissingDates= insertMissingDates(dfSelected[(dfSelected['Plan Type']==eachPlan)],dfPaycheck['Bi-Weekly Paycheck Dates for Dental/Vision/Basic Disability'])
            else:
                lInsertedMissingDates= insertMissingDates(dfSelected[(dfSelected['Plan Type']==eachPlan)],dfPaycheck['Bi-Weekly Paycheck Dates'])
            newPD = pd.DataFrame(lInsertedMissingDates)
            dfSelected= pd.concat([dfSelected,newPD], ignore_index=True)

        # dfMaster is the original dataframe
        dfMaster = df
        # df is a concatenation between dfSelected and dfOther dataframes
        df = pd.concat([dfSelected,dfOther],ignore_index=True)
        df.sort_values(by='Paycheck Issue Date',axis=0,inplace=True)
        df.reset_index(drop=True,inplace=True)

        #===========================================
        #Step 3. Identify Variance
        #===========================================
        # create 'Paycheck Year','Paycheck Month','Deduction Variance' columns.
        # Paycheck Year column - contains year of the Paycheck Issue Date 
        # Paycheck Month column - contains the month of the Paycheck Issue Date
        # Deduction Variance column - contains the difference of 
        df['Paycheck Year'] = df['Paycheck Issue Date'].dt.year
        df['Paycheck Month'] = df['Paycheck Issue Date'].dt.month
        df = df.sort_values(by=['Plan Type', 'Paycheck Year','Paycheck Issue Date'])
        df['Deduction Variance'] = (df.groupby(['Paycheck Year','Plan Type','Deduction Classification'])['Current Deduction'].diff())

        #### THIS ONLY APPLIES TO BIWEEKLY
        df['Deduction Variance'] = df.apply(
            lambda row: None if row['Paycheck Month'] == 12 else row['Deduction Variance'],
            axis=1
        )
        #===========================================
        #Step 4. Highlight Anomalies and Create (001 worksheet)
        #===========================================
        rules = {
            'Payback Amount':      (lambda v: v != 0, ['Payback Amount'],"green"),
            'Refund Amount':       (lambda v: v != 0, ['Refund Amount'],"blue"),
            'Not Taken':           (lambda v: v != 0, ['Not Taken'],"pink"),
            'Reason':              (lambda v: not(pd.isna(v)), ['Reason'],"red"),
            'Paycheck Status':     (lambda v: v != "Confirmed", ['Paycheck Status'],"yellow"),
            'Paycheck Option':     (lambda v: v != "Advice", ['Paycheck Option'],"purple"),
            'Off Cycle':           (lambda v: v != "N", ['Off Cycle'],"brown"),
        }

        df001 = df.style.apply(lambda row: highlightAll(row,rules,selectedPlans), axis=1)
        
        #===========================================
        #Step 5. Pivot Table and Create Pivot DataFrame (Summary worksheet)
        #===========================================
        # 1. Filter to selected plan types
        filtered_df = df[df['Plan Type'].isin(selectedPlans)]
        # 2. Create a 'Highlight' flag where 'Is New Row' is True
        filtered_df['Highlight'] = filtered_df['Is New Row'] == True
        # 2.a. Find inserted Paycheck Issue Dates for the header highlighting
        inserted_dates = df.loc[df['Is New Row'] == True, 'Paycheck Issue Date'].unique()
        # 3. Create the pivot table
        pivotedSummary = pd.pivot_table(
            filtered_df,
            index=['Deduction Classification', 'Plan Type', 'Benefit Plan'],
            columns='Paycheck Issue Date',
            values='Current Deduction',
            aggfunc='sum',
            fill_value=0
        )
        # 4. Create the highlight mask (same shape as pivotedSummary)
        highlight_rows = filtered_df[filtered_df['Highlight'] == True]
        highlight_mask = pd.pivot_table(
            highlight_rows,
            index=['Deduction Classification', 'Plan Type', 'Benefit Plan'],
            columns='Paycheck Issue Date',
            values='Current Deduction',
            aggfunc=lambda x: True
        ).reindex_like(pivotedSummary).fillna(False)
        # 5. Highlight function using the mask
        def highlight_inserted(val, mask_val):
            return 'background-color: pink' if mask_val else ''
        # 6. Apply the styles using apply + mask
        def highlight_with_mask(row, mask_df):
            return [
                'background-color: pink' if mask_df.loc[row.name, col] else ''
                for col in row.index
            ]
        # 6.a. highlight column headers
        def highlight_headers(col_names):
            return ['background-color: pink' if col in inserted_dates else '' for col in col_names]

        styled = pivotedSummary.style.apply(
            lambda row: highlight_with_mask(row, highlight_mask),
            axis=1
        )
        dfSummary = styled.apply_index(
            highlight_headers,
            axis=1  # 1 means apply on columns (the header row)
        )

        #============================================
        #Step 6. Put all dataframes into one workbook
        #============================================
        todaydate = datetime.now().strftime("%m-%d-%Y_%H-%M-%S")
        excel_datetime_format = "m/d/yyyy"
        emplid=df['Employee ID'].loc[0]
        outputFile= r"C:\ProgramData\AutomationAnywhere\Bots\Logs\RFL-UCPath\ProcessLogs\RFLReport_"+emplid+"_"+todaydate+".xlsx"
        with pd.ExcelWriter(outputFile,datetime_format=excel_datetime_format, engine='xlsxwriter') as writer:
            dfSummary.to_excel(writer, engine='openpyxl',sheet_name='Summary')
            df001.to_excel(writer,index=False,sheet_name='001')
            print("done writing to excel")
        #============================================
        #Step 7. Add Legends to FIN001 output
        #============================================
        AddLegends(outputFile)

        return("Success")
    except Exception as E:
        return(E)