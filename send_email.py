import os
import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass

BASE_DIR = Path(__file__).parent.resolve()
SUMMARY_PATH = BASE_DIR / "run_summary.json"


def get_latest_excel_path():
    candidate_files = [
        BASE_DIR / "Outlet name and Link - Updated.xlsx",
        BASE_DIR / "Outlet name and Link.xlsx",
        BASE_DIR / "Bvgg GMB Rating Sheet (2) - Updated.xlsx",
        BASE_DIR / "Bvgg GMB Rating Sheet (2).xlsx"
    ]
    existing = [f for f in candidate_files if f.exists()]
    if existing:
        return max(existing, key=os.path.getmtime)
    return BASE_DIR / "Outlet name and Link.xlsx"


def build_category_table(title, cat_data):
    if not cat_data:
        return ""

    results = cat_data.get("results", [])
    total = cat_data.get("total", len(results))
    success = sum(1 for x in results if x.get("status") == "success")

    rows_html = ""
    for idx, item in enumerate(results, 1):
        sno = item.get("sno") or idx
        code = item.get("store_code", "")
        addr = item.get("address", "")
        rating = item.get("rating")
        reviews = item.get("reviews")
        status = item.get("status", "unknown")

        if status == "closed":
            r_disp = "<span style='color:#7f8c8d; font-weight:600;'>CLOSED</span>"
            rev_disp = "<span style='color:#7f8c8d;'>-</span>"
            badge = "<span style='color:#7f8c8d; font-weight:600;'>CLOSED</span>"
        elif status == "success":
            r_disp = f"{rating} ★" if isinstance(rating, (int, float)) else str(rating)
            rev_disp = f"{reviews:,}" if isinstance(reviews, int) else str(reviews)
            badge = "<span style='color:#38761d; font-weight:600;'>SUCCESS</span>"
        else:
            r_disp = f"{rating}" if rating is not None else "<span style='color:#cc0000;'>N/A</span>"
            rev_disp = f"{reviews}" if reviews is not None else "<span style='color:#cc0000;'>N/A</span>"
            badge = "<span style='color:#b45f06; font-weight:600;'>PARTIAL</span>"

        row_bg = "#ffffff" if idx % 2 != 0 else "#fbfbfb"
        rows_html += f"""
        <tr style="background-color: {row_bg}; border-bottom: 1px solid #e0e0e0;">
            <td style="padding: 7px 10px; text-align: center; color: #666; font-weight: 600;">{sno}</td>
            <td style="padding: 7px 10px; font-weight: bold; color: #202124;">{code}</td>
            <td style="padding: 7px 10px; color: #444;">{addr}</td>
            <td style="padding: 7px 10px; text-align: center; font-weight: bold; color: #b45f06; background-color: #fff2cc;">{r_disp}</td>
            <td style="padding: 7px 10px; text-align: center; font-weight: bold; color: #134f5c; background-color: #d0e0e3;">{rev_disp}</td>
            <td style="padding: 7px 10px; text-align: center;">{badge}</td>
        </tr>
        """

    return f"""
    <div style="margin-top: 22px;">
        <h3 style="margin: 0 0 6px 0; color: #1a73e8; font-size: 15px; border-bottom: 2px solid #1a73e8; padding-bottom: 4px;">
            📍 {title} ({success}/{total} updated)
        </h3>
        <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
            <thead>
                <tr style="background-color: #202124; color: white;">
                    <th style="padding: 7px 10px; text-align: center; width: 45px;">S.No</th>
                    <th style="padding: 7px 10px; text-align: left; width: 80px;">Code</th>
                    <th style="padding: 7px 10px; text-align: left;">Store / Outlet</th>
                    <th style="padding: 7px 10px; text-align: center; width: 90px;">Live Rating</th>
                    <th style="padding: 7px 10px; text-align: center; width: 100px;">Review Count</th>
                    <th style="padding: 7px 10px; text-align: center; width: 80px;">Status</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
    </div>
    """


def build_email_html(summary):
    date_str = summary.get("date", datetime.now().strftime("%d/%m/%Y"))
    total = summary.get("total_stores", 0)
    success = summary.get("successful", 0)
    closed = summary.get("closed", 0)
    failed = summary.get("failed", 0)
    cats = summary.get("categories", {})

    tables_html = ""
    for tab_title in [
        "Bikanervala GGN",
        "Competitors Comparison",
        "India BV",
        "International BV",
        "Haldiram India",
        "OM Sweets India"
    ]:
        if tab_title in cats:
            tables_html += build_category_table(tab_title, cats[tab_title])

    return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f5f7; margin: 0; padding: 20px; }}
        .box {{ max-width: 900px; margin: 0 auto; background: #ffffff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); overflow: hidden; }}
        .header {{ background: linear-gradient(135deg, #FF6F00 0%, #D84315 100%); color: white; padding: 22px 28px; }}
        .header h1 {{ margin: 0; font-size: 22px; }}
        .header p {{ margin: 5px 0 0 0; opacity: 0.9; font-size: 13px; }}
        .body-content {{ padding: 24px 28px; }}
    </style>
</head>
<body>
    <div class="box">
        <div class="header">
            <h1>📊 Daily GMB Rating & Review Report (All Outlets)</h1>
            <p>Date: <strong>{date_str}</strong> | Format: Bvgg GMB Master Sheet (Exact Serial Order)</p>
        </div>
        <div class="body-content">
            <table style="width: 100%; border: none; margin-bottom: 15px;">
                <tr>
                    <td style="width: 25%; padding: 12px; background: #e8f0fe; border-radius: 6px; text-align: center;">
                        <div style="font-size: 22px; font-weight: bold; color: #1a73e8;">{total}</div>
                        <div style="font-size: 11px; text-transform: uppercase; color: #5f6368;">Total Stores</div>
                    </td>
                    <td style="width: 25%; padding: 12px; background: #e6f4ea; border-radius: 6px; text-align: center;">
                        <div style="font-size: 22px; font-weight: bold; color: #137333;">{success}</div>
                        <div style="font-size: 11px; text-transform: uppercase; color: #5f6368;">Active Updated</div>
                    </td>
                    <td style="width: 25%; padding: 12px; background: #f1f3f4; border-radius: 6px; text-align: center;">
                        <div style="font-size: 22px; font-weight: bold; color: #5f6368;">{closed}</div>
                        <div style="font-size: 11px; text-transform: uppercase; color: #5f6368;">Closed Stores</div>
                    </td>
                    <td style="width: 25%; padding: 12px; background: #fce8e6; border-radius: 6px; text-align: center;">
                        <div style="font-size: 22px; font-weight: bold; color: #c5221f;">{failed}</div>
                        <div style="font-size: 11px; text-transform: uppercase; color: #5f6368;">Failed / Missed</div>
                    </td>
                </tr>
            </table>

            <p style="font-size: 13px; color: #333; line-height: 1.6;">
                📎 <strong>The updated Excel file is attached:</strong> <code>Outlet name and Link.xlsx</code> (all 6 sheets: Bikanervala GGN, Competitors, India BV, International BV, Haldiram India, and OM Sweets India).<br>
                The store serial numbers and row layout match your master sheets <strong>1-to-1</strong> so you can directly copy today's column (or use the attached workbook).
            </p>

            {tables_html}
        </div>
        <div style="padding: 15px 28px; background: #f8f9fa; border-top: 1px solid #e9ecef; font-size: 12px; color: #70757a; text-align: center;">
            Bikanervala Automated Rating & Review Tracking System
        </div>
    </div>
</body>
</html>
"""


def main():
    sender = os.getenv("EMAIL_SENDER")
    password = os.getenv("EMAIL_PASSWORD")
    receiver = os.getenv("EMAIL_RECEIVER")
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))

    if not sender or not password or not receiver:
        print("[Notice] Email credentials not set in .env.")
        return

    summary = {}
    if SUMMARY_PATH.exists():
        with open(SUMMARY_PATH, "r", encoding="utf-8") as f:
            summary = json.load(f)

    date_str = summary.get("date", datetime.now().strftime("%d/%m/%Y"))

    msg = MIMEMultipart("mixed")
    msg["From"] = sender
    msg["To"] = receiver
    msg["Subject"] = f"📊 Daily GMB Rating & Review Report - {date_str} (All Outlets)"

    html_content = build_email_html(summary)
    msg.attach(MIMEText(html_content, "html"))

    target_excel = get_latest_excel_path()
    if target_excel and target_excel.exists():
        with open(target_excel, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{target_excel.stem} - {date_str.replace("/", "-")}.xlsx"',
        )
        msg.attach(part)
        print(f"[Attachment] Attached {target_excel.name}")

    print(f"[Email] Connecting to {smtp_server}:{smtp_port}...")
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender, password)

    receivers = [r.strip() for r in receiver.split(",") if r.strip()]
    server.sendmail(sender, receivers, msg.as_string())
    server.quit()
    print(f"[Success] Email successfully sent to {receiver}!")


if __name__ == "__main__":
    main()
