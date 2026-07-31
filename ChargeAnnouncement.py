import argparse
import logging
import os
from pathlib import Path
from logging.handlers import RotatingFileHandler

import pandas as pd

from sendEmail import Email


BASE_EXCEL_DIR = r"\\jpdejstcfs01\STC_share\●物流&OBM共用\蓄電池相關"
SCRIPT_DIR = Path(__file__).resolve().parent
IT_REMINDER_SENDER = "SRV.ITREMIND.RBT@deltaww.com"

TPS_RECIPIENTS = [
    # Production TPS recipients:
    "LUKE.ZHANG@deltaww.com",
    "XIN.ZHONG@DELTAWW.com",
    "V-CHIAHSING.LIN@DELTAWW.com",
    "V-YAOTING.YANG@DELTAWW.com",
    "SHAOJIN.WU@DELTAWW.com",
    "SA.PAPANA@deltaww.com",
    "AA.ZHU@deltaww.com",
]

PVI_RECIPIENTS = [
    # Production PVI recipients:
    "XIN.ZHONG@deltaww.com",
    "COOPER.ZHAO@deltaww.com",
    "V-JUNTENG.MA@deltaww.com",
    "V-JING.ZHOU@deltaww.com",
    "LANG.LUAN@deltaww.com",
]

JOBS = {
    "TPS": {
        "excel_filename": "TPS蓄電池検査_FAE物流共用表單_2025.xlsx",
        "sender_email": os.getenv("TPS_SENDER_EMAIL", IT_REMINDER_SENDER),
        "recipients": TPS_RECIPIENTS,
        "sheet_name": "周轉品",
        "log_filename": "TPS_Battery_Rotating_Stock_log.txt",
        "subject": "Charging Warning for Inventory Items",
    },
    "PVI": {
        "excel_filename": "PVI蓄电池検査2026.xlsx",
        "sender_email": os.getenv("PVI_SENDER_EMAIL", IT_REMINDER_SENDER),
        "recipients": PVI_RECIPIENTS,
        # PVI workbook tab is named Sheet1, but the sheet content is 新品.
        "sheet_name": "Sheet1",
        "log_filename": "PVI_Battery_Rotating_Stock_log.txt",
        "subject": "Charging Warning for Inventory Items",
    },
}


def parse_args():
    parser = argparse.ArgumentParser(description="Send battery charging warning email.")
    parser.add_argument(
        "jobs",
        nargs="*",
        help="Battery inspection sources to process. Default: run all jobs.",
    )
    parser.add_argument(
        "--sender-email",
        help="Override sender email for this run.",
    )
    parser.add_argument(
        "--excel-file",
        help="Override Excel file path for this run.",
    )
    args = parser.parse_args()
    invalid_jobs = [job_name for job_name in args.jobs if job_name not in JOBS]
    if invalid_jobs:
        parser.error(
            f"Unknown job(s): {', '.join(invalid_jobs)}. "
            f"Choose from: {', '.join(JOBS)}."
        )
    return args


def setup_logging(log_file):
    handler = RotatingFileHandler(
        log_file,
        maxBytes=10 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8",
    )
    logging.basicConfig(
        handlers=[handler],
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        force=True,
    )


def build_config(args):
    config = JOBS[args.job_name].copy()
    config["job_name"] = args.job_name
    config["excel_file"] = args.excel_file or str(
        Path(BASE_EXCEL_DIR) / config["excel_filename"]
    )
    config["sender_email"] = args.sender_email or config["sender_email"]
    if not config["sender_email"]:
        raise ValueError(
            f"{args.job_name} sender email is not configured. "
            "Set the sender with --sender-email or the related *_SENDER_EMAIL environment variable."
        )
    config["log_file"] = SCRIPT_DIR / config["log_filename"]
    return config


def read_excel(config):
    try:
        excel_book = pd.ExcelFile(config["excel_file"])
        sheet_name = config["sheet_name"]
        if sheet_name not in excel_book.sheet_names:
            fallback_sheet = excel_book.sheet_names[0]
            logging.warning(
                "Worksheet '%s' not found in %s. Available sheets: %s. Using first sheet: %s",
                sheet_name,
                config["excel_file"],
                ", ".join(excel_book.sheet_names),
                fallback_sheet,
            )
            sheet_name = fallback_sheet

        df_all = pd.read_excel(excel_book, sheet_name=sheet_name)
        logging.info(
            "Successfully read Excel file: %s, Sheet: %s",
            config["excel_file"],
            sheet_name,
        )
        return df_all
    except Exception as e:
        logging.error("Error reading Excel file: %s", e)
        raise


def convert_soc(value):
    try:
        if pd.isna(value) or value is None:
            return None
        if isinstance(value, (int, float)):
            return f"{value * 100:.1f}%" if 0 <= value <= 1 else str(value)
        return str(value)
    except Exception as e:
        logging.error("Error processing SOC%% value %s: %s", value, e)
        return str(value)


def clean_data(df_all):
    if "SOC%" in df_all.columns:
        df_all["SOC%"] = df_all["SOC%"].apply(convert_soc)
        logging.info("SOC%% column successfully processed")

    date_columns = ["Date", "Charging warning date"]
    for col in date_columns:
        if col in df_all.columns:
            df_all[col] = pd.to_datetime(df_all[col], errors="coerce")
            df_all[col] = df_all[col].where(df_all[col] >= pd.Timestamp("1753-01-01"), None)
            logging.info("%s column successfully converted to datetime format", col)

    if "No." in df_all.columns:
        df_all["No."] = pd.to_numeric(df_all["No."], errors="coerce").fillna(0).astype(int)
        logging.info("No. column successfully converted to integer format")

    df_all = df_all.where(pd.notna(df_all), None)
    logging.info("Missing values handled successfully")
    return df_all


def query_warnings(df_all):
    required_columns = {"Charging warning date", "Remark"}
    missing_columns = required_columns - set(df_all.columns)
    if missing_columns:
        raise ValueError(f"Missing required columns in Excel: {', '.join(sorted(missing_columns))}")

    warning_date = pd.to_datetime(df_all["Charging warning date"], errors="coerce")
    remark = df_all["Remark"].fillna("").astype(str).str.strip()
    mask = (warning_date <= pd.Timestamp.now()) & (remark == "Inventory")
    df_warning = df_all.loc[mask].copy()
    logging.info("Found %s records for charging warning", len(df_warning))
    return df_warning


def build_email_body(df_warning):
    if not df_warning.empty:
        for col in ["Date", "Charging warning date"]:
            if col in df_warning.columns:
                if not pd.api.types.is_datetime64_any_dtype(df_warning[col]):
                    df_warning[col] = pd.to_datetime(df_warning[col], errors="coerce")
                df_warning[col] = df_warning[col].dt.strftime("%Y-%m-%d")

        body_content = """
    <p style="font-size: 18px; font-family: 'Arial', sans-serif; color: #333;">The following models need attention for charging and discharging.</p>
    """
    else:
        body_content = """
    <p style="font-size: 18px; font-family: 'Arial', sans-serif; color: #333;">No charging or discharging needed for models this week.</p>
    """

    html_table = df_warning.to_html(index=False, escape=False)
    if "Charging warning date" in df_warning.columns:
        html_table = html_table.replace(
            "<th>Charging warning date</th>",
            "<th style='background-color: #FF0000; color: white;'>Charging warning date</th>",
        )

    return f"""
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


def send_warning_email(config, body):
    # Local PC test: set DELTA_SMTP_PASSWORD before running to authenticate with deltarelay.
    # VM deployment: leave DELTA_SMTP_PASSWORD unset to use the relay anonymously.
    password = os.getenv("DELTA_SMTP_PASSWORD", "")
    email = Email()
    attachments = [str(config["log_file"])] if config["log_file"].exists() else []
    receiver_email = ",".join(config["recipients"])

    email.send_email(
        config["sender_email"],
        password,
        receiver_email,
        config["subject"],
        body,
        attachments,
    )
    logging.info("Email sent successfully to all recipients")


def run_job(job_name, args):
    args.job_name = job_name
    config = build_config(args)
    setup_logging(config["log_file"])
    logging.info("Program started running for %s", config["job_name"])

    df_all = clean_data(read_excel(config))
    df_warning = query_warnings(df_all)
    body = build_email_body(df_warning)
    send_warning_email(config, body)
    logging.info("Program finished successfully for %s", config["job_name"])


def main():
    args = parse_args()
    selected_jobs = args.jobs or list(JOBS.keys())
    if args.excel_file and len(selected_jobs) != 1:
        raise ValueError("--excel-file can only be used when running a single job.")

    failures = []
    for job_name in selected_jobs:
        try:
            run_job(job_name, args)
        except Exception as e:
            failures.append((job_name, e))
            logging.exception("Program failed for %s", job_name)

    if failures:
        failed_names = ", ".join(job_name for job_name, _ in failures)
        raise SystemExit(f"Failed jobs: {failed_names}")


if __name__ == "__main__":
    main()
