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

report_name = "Forecast w Tool Export"
print(f"Running {report_name}")

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

    # Prompt user
    forecast_id = input("Enter Forecast Id: ")

    sql_query = """
select DISTINCT e.eqno, c.mfgno, b.itemno, b.descrip, b.descrip2, g.mfgtype
, case when descrip2 in ('DZ1','CH2','TR8','CH2/DZ1') then 'Current Model' else 'Service' end as part_type
, a.eplant_id, a.qty, a.cycles_req, c.cycletm, round(a.qty/a.cycles_req,0) as cav, g.mfgcell, to_char(trunc(a.req_date,'IW')
,'mm/dd/yy') as week_start, a.year, a.month, a.week as weekno, to_char(a.req_date + 7,'mm/dd/yy') as req_date
, a.fore_head_id as forecast_id, a.name as forecast_name
from iqms.v_forecast a 
left join iqms.arinvt b on b.id = a.arinvt_id 
left join iqms.standard c on c.id = b.standard_id 
left join iqms.standard_pmeqmt d on d.standard_id = c.id 
left join iqms.pmeqmt e on e.id = d.pmeqmt_id 
left join iqms.mfgcell g on g.id = c.mfgcell_id 
join iqms.fore_dtl x on a.FORE_DTL_ID = x.id
where 1=1 
and a.qty <> 0 
and a.fore_head_id = :forecast_id
and g.mfgtype = 'INJECTION' 
order by a.year, a.month, c.mfgno, b.itemno
"""           
    # Query    
    #df = pd.read_sql(sql_query, engine)
    #output_file = folder + "//Data//" + report_name + ".xlsx"

    df = pd.read_sql(   
        sql_query,
        engine,
        params={"forecast_id": int(forecast_id)}
    )   

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

    email_to = 'gmarks@hauto.net'

    if not email_to or pd.isna(email_to):
        print(f"No email_to, skipping email")
    else:
        # Create Outlook instance
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        # Email details
        mail.To = email_to
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
        print("Database connection disposed")
finally:
    if engine is not None:
        engine.dispose()
        print("Database connection disposed")
print("=" * 50)
print("Program Complete")
print("=" * 50)