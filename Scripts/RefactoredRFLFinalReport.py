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
                            'Current Deduction':0,
                            'Payback Amount':0,
                            'Refund Amount':0,
                            'Not Taken':0,
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
                        'Current Deduction':0,
                        'Payback Amount':0,
                        'Refund Amount':0,
                        'Not Taken':0,
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

def AddLegends(fin001outputfile,emplid):
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
        SummaryWB[f'E{legendRow}']="EEID"
        SummaryWB[f'F{legendRow}']= emplid
        SummaryWB[f'E{legendRow+1}']="RFL"
    except Exception as E:
        return f"Inside AddLegends function.{type(E).__name__}: {E}"
    finally:
        wb.save(filepath)
        wb.close()
        return("success")


def highlight_inserted(val, mask_val):
    return 'background-color: lightgray' if mask_val else ''

def highlight_with_mask(row, mask_df):
    return [
        'background-color: lightgray' if mask_df.loc[row.name, col] else ''
        for col in row.index
    ]



def main(params):
    
    try:
        #===========================================
        #Step 1. Read FINOO1 sample report and paycheck table
        #=========================================== 
        emplid = params[0]
        FIN001filepath = params[2]
        payGroupType = params[1]
        df = pd.read_excel(FIN001filepath, header=1)
        paycheckfilepath=""
        if payGroupType =='BIWEEKLY':
            paycheckfilepath=r"C:\ProgramData\AutomationAnywhere\Bots\Asset\BiWeekly Paycheck Dates.xlsx"
        else:
            paycheckfilepath=r"C:\ProgramData\AutomationAnywhere\Bots\Asset\Monthly Paycheck Dates.xlsx"
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
        # as of 02/3/2026, this line is moved up. dfOrigin is the original dataframe
        dfOrigin = df
        # as of 02/3/2026 add column 'is new row' to df and set to np.nan (blank)
        df['Is New Row']= np.nan
        # create 2 dataframes: Selected vs Other
        dfSelected = df[(df['Plan Type'].isin(selectedPlans))]
        dfOther =df[~(df['Plan Type'].isin(selectedPlans))]
        # as of 02/4/2026 check dfSelected to see if there is/are rows where plan type is in selectedPlans
        if len(dfSelected) ==0:
            return "No In-Scope Plan Type"
        # iterate each selected plan and insert missed paycheck dates
        if payGroupType == "BIWEEKLY":
            for eachPlan in dfSelected['Plan Type'].unique():
                if eachPlan in(['Dental','Life','Basic Disability','Vision']):
                    lInsertedMissingDates= insertMissingDates(dfSelected[(dfSelected['Plan Type']==eachPlan)],dfPaycheck['Bi-Weekly Paycheck Dates for Dental/Vision/Basic Disability'])
                else:
                    lInsertedMissingDates= insertMissingDates(dfSelected[(dfSelected['Plan Type']==eachPlan)],dfPaycheck['Bi-Weekly Paycheck Dates'])
                newPD = pd.DataFrame(lInsertedMissingDates)
                dfSelected= pd.concat([dfSelected,newPD], ignore_index=True)
        else:
            for eachPlan in dfSelected['Plan Type'].unique():
                linsertedMissingDates = insertMissingDates(dfSelected[(dfSelected['Plan Type']==eachPlan)],dfPaycheck['Monthly Paycheck Dates 2022 forward '])
                newPD = pd.DataFrame(linsertedMissingDates)
                dfSelected = pd.concat([dfSelected,newPD], ignore_index=True)

        # df is a concatenation between dfSelected and dfOther dataframes
        df = pd.concat([dfSelected,dfOther],ignore_index=True)
        df.sort_values(by='Paycheck Issue Date',axis=0,inplace=True)
        df.reset_index(drop=True,inplace=True)
        # as of 02/4/2026 add If ~df['Is New Row'].any() which will return True if there is no value in that column that is true. np.nan is ignored in this boolean
        if ~ df['Is New Row'].any():
            return"No Missed Paycheck Date"
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
        # as of 02/03/2026 moved up highlight_headers  
        def highlight_headers(col_names):
            return ['background-color: pink' if col in inserted_dates else '' for col in col_names]
    
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
        excel_datetime_format = "m-d-yyyy"
        outputfilename = FIN001filepath.split("\\")[-1]
        outputfilename = outputfilename.split(".")[0]
        outputFile= f"C:\ProgramData\AutomationAnywhere\Bots\Logs\RFL-UCPATH\ProcessLogs\PendingReconReport\{outputfilename}_RFL_RPA_PENDING_VALIDATION.xlsx"
        with pd.ExcelWriter(outputFile,datetime_format=excel_datetime_format, engine='xlsxwriter') as writer:
            dfSummary.to_excel(writer, engine='openpyxl',sheet_name='Summary')
            df001.to_excel(writer,index=False,sheet_name='001')
            dfOrigin.to_excel(writer,index=False,sheet_name='Original FIN001')
            print("done writing to excel")
        #============================================
        #Step 7. Add Legends to FIN001 output
        #============================================
        AddLegends(outputFile,emplid)

        return("Success")
    except Exception as E:
        return f"{type(E).__name__}: {E}"
