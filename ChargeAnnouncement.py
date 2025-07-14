import pandas as pd
import pyodbc
from datetime import datetime,timedelta
import logging
from sendEmail import Email  # Assuming an Email module handles email sending
from logging.handlers import RotatingFileHandler

# Set Excel file path
excel_file = r"\\jpdejstcfs01\STC_share\●物流&OBM共用\蓄電池相關\TPS蓄電池検査_FAE物流共用表單_2025.xlsx"

# Set log file path
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
log_file = "Battery/TPS_Battery_Rotating_Stock_log.txt"
handler = RotatingFileHandler(log_file, maxBytes=10*1024*1024, backupCount=5)  # 10MB per file, keep 5 backups
logging.basicConfig(handlers=[handler], level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# Log start
logging.info("Program started running")

# Specify the sheet to read
specific_sheet = "周轉品"

# Read the specified Excel sheet
try:
    df_all = pd.read_excel(excel_file, sheet_name=specific_sheet)
    logging.info(f"Successfully read Excel file: {excel_file}, Sheet: {specific_sheet}")
except Exception as e:
    logging.error(f"Error reading Excel file: {e}")
    raise


# Convert "SOC%" column
if "SOC%" in df_all.columns:
    def convert_soc(value):
        try:
            if pd.isna(value) or value is None:
                return None
            elif isinstance(value, (int, float)):
                return f"{value * 100:.1f}%" if 0 <= value <= 1 else str(value)
            else:
                return str(value)
        except Exception as e:
            logging.error(f"Error processing SOC% value {value}: {e}")
            return str(value)

    df_all["SOC%"] = df_all["SOC%"].apply(convert_soc)
    logging.info("SOC% column successfully processed")



date_columns = ["Date", "Charging warning date"]
for col in date_columns:
    if col in df_all.columns:
        df_all[col] = pd.to_datetime(df_all[col], errors="coerce")  # Convert to datetime
        df_all[col] = df_all[col].where(df_all[col] >= pd.Timestamp('1753-01-01'), None)  # Set invalid dates to None
        logging.info(f"{col} column successfully converted to datetime format")

if "No." in df_all.columns:
    df_all["No."] = pd.to_numeric(df_all["No."], errors="coerce").fillna(0).astype(int)
    logging.info("No. column successfully converted to integer format")

df_all = df_all.astype(str)
df_all = df_all.where(df_all != "nan", None)  # 轉換 NaN 為 None
logging.info("Missing values handled successfully")

# Connect to SQL Server (using Windows authentication)
conn = pyodbc.connect(
    "DRIVER={SQL Server};"
    "SERVER=jpdejitdev01;"
    "DATABASE=ITQAS2;"
    "Trusted_Connection=yes;"
)
cursor = conn.cursor()
logging.info("Successfully connected to the database")

table_name = "TPS_Bettery_Rotating_Stock"

# 檢查資料表是否存在，若不存在則建立
check_table_query = f"""
IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{table_name}')
BEGIN
    CREATE TABLE {table_name} (
        {", ".join(f"[{col}] NVARCHAR(255)" for col in df_all.columns)}
    )
END
"""
cursor.execute(check_table_query)
conn.commit()
logging.info(f"Checked and created table if not exists: {table_name}")

# 清空表格
cursor.execute(f"DELETE FROM {table_name}")
conn.commit()
logging.info(f"Cleared table: {table_name}")


# Insert data
columns = ", ".join([f"[{col}]" for col in df_all.columns])
placeholders = ", ".join(["?" for _ in df_all.columns])
sql = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"
cursor.executemany(sql, df_all.values.tolist())
conn.commit()
logging.info(f"Successfully inserted {len(df_all)} records into {table_name}")

# 設定提前14天通知
days_before_warning = 14
warning_threshold = datetime.today() + timedelta(days=days_before_warning)

# Query charging warning items
query = f"""
SELECT * FROM {table_name}
WHERE [Charging warning date] <= ? AND [Remark] = 'Inventory'
"""
df_warning = pd.read_sql(query, conn, params=[warning_threshold])
logging.info(f"Found {len(df_warning)} records for charging warning")

cursor.close()
conn.close()
logging.info("Database connection closed successfully")

# Prepare email content
sender_email = "SRV.ITREMIND.RBT@deltaww.com"
password = "Dej1tasd"
email = Email()
subject = "Charging Warning for Inventory Items"

#charging data formate 
if not df_warning.empty:
    for col in ["Date", "Charging warning date"]:
        if col in df_warning.columns:
            # Check if the column is datetime type; if not, convert it
            if not pd.api.types.is_datetime64_any_dtype(df_warning[col]):
                df_warning[col] = pd.to_datetime(df_warning[col], errors="coerce")
            df_warning[col] = df_warning[col].dt.strftime('%Y-%m-%d')
    
    html_table = df_warning.to_html(index=False, escape=False)
    body_content = """
    <p style="font-size: 18px; font-family: 'Arial', sans-serif; color: #333;">The following models need attention for charging and discharging.</p>
    """
else:
    html_table = "<table><tr>" + "".join(f"<th>{col}</th>" for col in df_all.columns) + "</tr></table>"
    body_content = """
    <p style="font-size: 18px; font-family: 'Arial', sans-serif; color: #333;">No charging or discharging needed for models this week.</p>
    """

charging_warning_date_index = 8

# Convert table and modify header background color for Charging warning date
html_table = df_warning.to_html(index=False, escape=False)

# Manually set red background color for Charging warning date column header
html_table = html_table.replace(
    f"<th>{df_warning.columns[charging_warning_date_index]}</th>", 
    f"<th style='background-color: #FF0000; color: white;'>{df_warning.columns[charging_warning_date_index]}</th>"
)

body = f"""
<html>
    <head>
        <style>
            table {{
                width: 100%;
                border-collapse: collapse;
                font-family: Arial, sans-serif;
            }}
            table, th, td {{
                border: 1px solid #ddd;
            }}
            th {{
                background-color: #4CAF50;
                color: white;
                text-align: center;
            }}
            td {{
                padding: 8px;
                text-align: center;
            }}
            tr:nth-child(even) {{background-color: #f2f2f2;}}
            tr:hover {{background-color: #ddd;}}
        </style>
    </head>
    <body>
        {body_content}
        {html_table}
    </body>
</html>
"""

# Send email with log file attached
# for u in ['boris.wang@deltaww.com']:
#     email.send_email(sender_email, password, u, subject, body, log_file)
for u in ['boris.wang@deltaww.com','JPSTC.LGS@deltaww.com','JPOBMFAE@deltaww.com']:
    email.send_email(sender_email, password, u, subject, body, log_file)

logging.info("Email sent successfully")
