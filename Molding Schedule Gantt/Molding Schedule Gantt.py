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
from configparser import ConfigParser

current_date = datetime.now()
sys.stdout.write(f"Current Date and Time: {current_date.strftime('%Y-%m-%d %H:%M')}\n")
folder = Path(__file__).resolve().parent.parent.parent
print("Folder path is: ",folder)
config_file = folder / "config" / "dbconfig.cfg"

config = ConfigParser()
config.read(config_file)

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

    report_name = "Molding Schedule Gantt"
    email_to = "gmarks@hauto.net"

    sql_query = """
        select distinct
        b.eqno as "EQ No",d.mfgno || ' / ' || ltrim(to_char(a.cycles_to_go,'9,999,999')) as "MfgNo-Cycles To Go" 
         , TO_CHAR(a.prod_start_time, 'YYYY-MM-DD HH24:MI:SS') as "Prod Start"
         , TO_CHAR(a.prod_end_time, 'YYYY-MM-DD HH24:MI:SS') as "Prod End"
         , (a.prod_end_time - a.prod_start_time) * 24 as "Duration Hours"
         ,    case when a.down_idle IN (1, 2) THEN 'Yellow'
                when c.start_time IS NULL OR c.start_time = TO_DATE('1899-12-30', 'YYYY-MM-DD') THEN 'Black'
                when a.prod_start_time <= c.start_time AND c.origin = 'FORECAST' THEN 'Blue'
                when a.prod_start_time <= c.start_time AND c.origin <> 'FORECAST' THEN 'Green'
                when a.prod_end_time >= h.must_ship_date and a.prod_end_time <= h.promise_date THEN 'Maroon'
                when (a.prod_end_time < h.must_ship_date or a.prod_end_time > h.promise_date) and c.origin = 'FORECAST' THEN 'Navy'
                when (a.prod_end_time < h.must_ship_date or a.prod_end_time > h.promise_date) AND c.origin <> 'FORECAST' THEN 'Red'
                else 'Red'
              end AS "Schedule Color"
        from   iqms.v_sched a
         left join iqms.work_center b on a.work_center_id=b.id
         left join iqms.workorder c on a.workorder_id=c.id
         left join iqms.standard d on c.standard_id=d.id
         left join iqms.ptorder e on c.id=e.workorder_id 
         left join iqms.ptorder_rel g on e.id=g.ptorder_id 
         left join iqms.ptallocate h on g.ptallocate_id=h.id
         where  (b.is_virtual is  null  or (b.is_virtual=' ' or b.is_virtual='N'))
         and a.prod_start_time <= trunc(sysdate) + 1
         and b.eqno like '20-MC-INJ%'
         order by 1
         """

    print("=" * 50)
    print(f"Report: {report_name}")
    print(f"Email: {email_to}")
    #print(f"SQL: {sql_query}")
        
    # Query
    df = pd.read_sql(sql_query, engine)

    df['Prod Start'] = pd.to_datetime(df['Prod Start'])
    df['Prod End'] = pd.to_datetime(df['Prod End'])

    output_file = folder / "Data" / f"{report_name}.xlsx"

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(output_file)
    ws = wb.Worksheets("Data")

    # Clear existing contents
    ws.Cells.ClearContents()

    # Headers
    for col_num, header in enumerate(df.columns, start=1):
        ws.Cells(1, col_num).Value = header

    # Data
    for row_num, row in enumerate(df.itertuples(index=False), start=2):
        for col_num, value in enumerate(row, start=1):

            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()

            ws.Cells(row_num, col_num).Value = value

    # Autofit columns
    ws.Columns.AutoFit()

    wb.Save()
    wb.Close(False)
    excel.Quit()
    print("Data successfully exported to excel")

    # Create Outlook instance
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    # Email details
    mail.To = "gmarks@hauto.net"
    #"gmarks@hauto.net"
    #email_to
    mail.CC = ""

    mail.BCC = ""
    mail.Subject = f" '{report_name}' "
    mail.Body = "All,\n\nSee Attached Excel File...\n\nThank-you"

    # Attach Excel file
    #mail.Attachments.Add(output_file)
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
        print("Database connection disposed.")

print("Program Complete")
