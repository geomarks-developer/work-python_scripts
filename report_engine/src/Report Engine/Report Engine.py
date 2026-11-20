import pandas as pd
import win32com.client as win32
from sqlalchemy import create_engine
from datetime import datetime
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
import sys
import os
from pathlib import Path

current_date = datetime.now()
print("Current Date and Time:", current_date.strftime("%Y-%m-%d %H:%M"))

folder = Path(__file__).resolve().parent.parent.parent
print("Folder path is: ",folder)


path = Path("Report Engine Report List.xlsx")
parm_file = folder / Path("Report Engine Report List.xlsx")  
x=parm_file
print ("Path is: ",parm_file)

# Ensure arguments were actually passed
if len(sys.argv) > 1:
    first_argument = sys.argv[1]
    print(f"Argument 1: {first_argument}")
    mode = sys.argv[1]
else:
    print("No arguments were provided.")
    mode = 'AdHoc'

print(f"Mode is: {mode}")

report_name = mode
print(f"Running {report_name}")

df = pd.read_excel(parm_file, sheet_name=mode)

from configparser import ConfigParser

config = ConfigParser()
config.read(folder /"config"/"dbconfig.cfg") 

user = config["oracle"]["user"]
password = config["oracle"]["password"]
host = config["oracle"]["host"]
port = config["oracle"]["port"]
service_name = config["oracle"]["service_name"]

try:

    engine = create_engine(
        f"oracle+oracledb://{user}:{password}@{host}:{port}/?service_name={service_name}"
    )

    # Process each report
    for _, row in df.iterrows():

        report_name = str(row["ReportName"]).strip()
        active = row["Active"]
        email_to = row["EmailTo"]
        sql_query = row["SQLQuery"]

        print("=" * 50)
        print(f"Active?: {active}")
        print(f"Report: {report_name}")
        print(f"Email: {email_to}")
        #print(f"SQL: {sql_query}")
        
        if not active:
            continue

        # Query
        
        df = pd.read_sql(sql_query, engine)
        #output_file = folder + "//Data//" + report_name + ".xlsx"
        output_file = Path(folder) / "Data" / f"{report_name}.xlsx"
        output_sheet = "Output"

        # Check if workbook exists
        from openpyxl import Workbook, load_workbook
        import os

        if os.path.exists(output_file):
            wb = load_workbook(output_file)
        else:
            wb = Workbook()
            wb.save(output_file)

        # Remove worksheet
        if output_sheet in wb.sheetnames:
            ws = wb[output_sheet]

            ws.delete_rows(2, ws.max_row)  # keep headers
        else:
            ws = wb.create_sheet(output_sheet)
        
        #wb.save(output_sheet)
        ws = wb[output_sheet]

        # Clear existing data (leave headers if desired)
        for row in ws.iter_rows(min_row=0, max_row=ws.max_row):
            for cell in row:
                cell.value = None

        # Update headers
        for col_num, header in enumerate(df.columns, start=1):
            ws.cell(row=1, column=col_num, value=header)
            
        # Format header row
        header_fill = PatternFill(
            fill_type="solid",
            fgColor="DDEBF7" # Light Blue
        )
        header_font = Font(bold=True)
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

        # Make header row taller
        ws.row_dimensions[1].height = 30

        # Write new data
        for row_num, row in enumerate(df.itertuples(index=False), start=2):
            for col_num, value in enumerate(row, start=1):
                ws.cell(row=row_num, column=col_num, value=value)

        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            # Skip header row
            for cell in column[1:]:
                if cell.value is not None and str(cell.value).strip() != "":
                    max_length = max(max_length, len(str(cell.value)))

            if max_length == 0:
                ws.column_dimensions[column_letter].width = 8.11
            else:
                header_length = len(str(column[0].value or ""))
                ws.column_dimensions[column_letter].width = max(max_length, header_length) + 2
                
        # Save workbook
        wb.save(output_file)

        print("Data successfully exported to excel")

        if not email_to or pd.isna(email_to):
            print(f"No email_to, skipping email")
        else:
            # Create Outlook instance
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            # Email details
            mail.To = email_to
            #"gmarks@hauto.net"
            #email_to
            mail.CC = ""

            mail.BCC = ""
            mail.Subject = f" '{report_name}' "
            mail.Body = "All,\n\nSee Attached Excel File...\n\nThank-you"

            # Attach Excel file
            mail.Attachments.Add(str(output_file))

            # Send email
            mail.Send()
            print("Email sent via Outlook")

except Exception as e:
    print(f"ERROR: {e}")
    if engine is not None:
        engine.dispose()
finally:
    if engine is not None:
        engine.dispose()
        print("Database connection disposed")
print("=" * 50)
print("Program Complete")
print("=" * 50)