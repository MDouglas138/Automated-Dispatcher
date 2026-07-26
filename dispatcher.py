import os
import sys
import sqlite3
import smtplib
from email.message import EmailMessage
import pandas as pd
from fpdf import FPDF
from dotenv import load_dotenv

# Load the environment variables from the .env file
load_dotenv()

# Define base directory globally so all functions can see it
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

MONTHS_IN_YTD = 6
RESTOCK_THRESHOLD = 2.0
OVERSTOCK_THRESHOLD = 5.0


def update_db_status(task_id, status):
    """Updates the status in the database so the UI can see it."""
    db_path = os.path.join(BASE_DIR, "automations.db")
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("UPDATE automations SET last_status = ? WHERE id = ?", (status, task_id))
        conn.commit()
        conn.close()
    except Exception as e:
        with open("error_log.txt", "a") as f:
            f.write(f"Database error: {e}\n")


def create_pdf_report(pacing_df, restock_df, trades_df, dead_stock_df, unmatched_overstock_df,
                      output_filename="Dealership_Intelligence_Report.pdf"):
    pdf = FPDF()
    pdf.add_page()

    # Title
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, txt="Automated Dealership Intelligence Report", ln=True, align='C')
    pdf.ln(10)

    # Section 1: Pacing Warnings
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt="Target Pacing Warnings (< 90% of Quota)", ln=True)
    pdf.set_font("Arial", size=12)
    if pacing_df.empty:
        pdf.cell(200, 10, txt="No locations are currently falling behind pace.", ln=True)
    else:
        for index, row in pacing_df.iterrows():
            text = f"Location: {row['Location']} | Model: {row['Model']} | Pace: {row['Pacing_Status'] * 100:.1f}%"
            pdf.cell(200, 10, txt=text, ln=True)
    pdf.ln(5)

    # Section 2: Critical Shortages
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt="Critical Inventory Shortages", ln=True)
    pdf.set_font("Arial", size=12)
    if restock_df.empty:
        pdf.cell(200, 10, txt="No critical shortages detected.", ln=True)
    else:
        for index, row in restock_df.iterrows():
            text = f"Location: {row['Location']} | Model: {row['Model']} | Supply: {row['Months_of_Supply']:.2f} mo | Order: {row['Suggested_Order']} units"
            pdf.cell(200, 10, txt=text, ln=True)
    pdf.ln(5)

    # Section 3: Dealer Trades
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt="Actionable Dealer Trades", ln=True)
    pdf.set_font("Arial", size=12)
    if trades_df.empty:
        pdf.cell(200, 10, txt="No optimal dealer trades found.", ln=True)
    else:
        for index, row in trades_df.iterrows():
            text = f"Move {row['Model']} from {row['Location_Over']} to {row['Location_Short']}"
            pdf.cell(200, 10, txt=text, ln=True)
    pdf.ln(5)

    # Section 4: Dead Stock
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt="Dead Stock Alerts (0 YTD Sales)", ln=True)
    pdf.set_font("Arial", size=12)
    if dead_stock_df.empty:
        pdf.cell(200, 10, txt="No dead stock detected.", ln=True)
    else:
        for index, row in dead_stock_df.iterrows():
            text = f"Location: {row['Location']} | Model: {row['Model']} | Stuck Units: {row['Current_Stock']}"
            pdf.cell(200, 10, txt=text, ln=True)
    pdf.ln(5)

    # Section 5: Unmatched Overstock
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, txt="Unmatched Overstock (Cannot be traded)", ln=True)
    pdf.set_font("Arial", size=12)
    if unmatched_overstock_df.empty:
        pdf.cell(200, 10, txt="No unmatched overstock detected.", ln=True)
    else:
        for index, row in unmatched_overstock_df.iterrows():
            text = f"Location: {row['Location']} | Model: {row['Model']} | Supply: {row['Months_of_Supply']:.2f} months"
            pdf.cell(200, 10, txt=text, ln=True)

    pdf.output(output_filename)
    return output_filename


def main():
    # Debug log start
    with open("debug_log.txt", "w") as f:
        f.write("Dispatcher started successfully.\n")

    if len(sys.argv) < 2:
        sys.exit()

    task_id = sys.argv[1]
    db_path = os.path.join(BASE_DIR, "automations.db")

    # 1. Connect to DB
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT file_path, sender_email, precheck_offset FROM automations WHERE id=?", (task_id,))
        auto_data = cursor.fetchone()
        cursor.execute("SELECT email_address FROM task_recipients WHERE task_id=?", (task_id,))
        recipients = [row[0] for row in cursor.fetchall()]
        conn.close()
    except sqlite3.Error as e:
        sys.exit()

    if not auto_data:
        sys.exit()

    target_file, sender_email, offset = auto_data
    app_password = os.environ.get('EMAIL_PASSWORD')

    if not app_password:
        update_db_status(task_id, "Error: No App Password")
        sys.exit()

    # 2. Ghost Drive Guardrail
    if not os.path.exists(target_file):
        update_db_status(task_id, "Error: File Missing")
        sys.exit()

    # 3. Data Extraction and Empty Guardrail
    try:
        ytd_df = pd.read_excel(target_file, sheet_name="2026_YTD_Summary")
        inv_df = pd.read_excel(target_file, sheet_name="Current_Inventory")
        monthly_df = pd.read_excel(target_file, sheet_name="Monthly_Sales")

        # Bulletproof Empty File Guardrail
        inv_cleaned = inv_df.dropna(how='all')
        if inv_cleaned.empty:
            update_db_status(task_id, "Error: Sheet contains no readable data")
            sys.exit()

    except Exception as e:
        with open("error_log.txt", "w") as f:
            f.write(f"CRASH: {str(e)}\nTarget File: {target_file}\nCWD: {os.getcwd()}")
        update_db_status(task_id, "Error: Data Read Failure")
        sys.exit()

    # 4. Calculations
    try:
        inv_melted = inv_df.melt(id_vars=["Location"], var_name="Model", value_name="Current_Stock")

        recent_sales = monthly_df[monthly_df['Month'].isin(['May', 'June'])]
        early_sales = monthly_df[monthly_df['Month'].isin(['Jan', 'Feb', 'Mar', 'Apr'])]

        recent_avg = recent_sales.groupby(['Location', 'Model'])['Units_Sold'].mean().reset_index(name='Recent_Avg')
        early_avg = early_sales.groupby(['Location', 'Model'])['Units_Sold'].mean().reset_index(name='Early_Avg')

        trend_df = pd.merge(recent_avg, early_avg, on=['Location', 'Model'], how='left')
        trend_df['Early_Avg'] = trend_df['Early_Avg'].replace(0, 0.1)
        trend_df['Trend_Weight'] = trend_df['Recent_Avg'] / trend_df['Early_Avg']
        trend_df['Trend_Weight'] = trend_df['Trend_Weight'].clip(upper=1.5, lower=0.5)

        merged_df = pd.merge(inv_melted, ytd_df, on=["Location", "Model"], how="left")
        merged_df = pd.merge(merged_df, trend_df[['Location', 'Model', 'Trend_Weight']], on=["Location", "Model"],
                             how="left")

        merged_df['Jan_Jun_Sales'] = merged_df['Jan_Jun_Sales'].fillna(0)
        merged_df['Jan_Jun_Sales'] = merged_df['Jan_Jun_Sales'].replace(0, 0.1)

        merged_df['Monthly_Run_Rate'] = (merged_df['Jan_Jun_Sales'] / MONTHS_IN_YTD) * merged_df['Trend_Weight'].fillna(
            1.0)
        merged_df['Months_of_Supply'] = merged_df['Current_Stock'] / merged_df['Monthly_Run_Rate']

        merged_df['Pacing_Target'] = merged_df['Projected_2026_Sales'] / 2
        merged_df['Pacing_Status'] = merged_df['Jan_Jun_Sales'] / merged_df['Pacing_Target']

        pacing_warnings_df = merged_df[merged_df['Pacing_Status'] < 0.90]

        shortages_df = merged_df[merged_df['Months_of_Supply'] < RESTOCK_THRESHOLD].copy()
        overstock_df = merged_df[merged_df['Months_of_Supply'] > OVERSTOCK_THRESHOLD].copy()

        dealer_trades_df = pd.merge(
            shortages_df[['Location', 'Model', 'Months_of_Supply']],
            overstock_df[['Location', 'Model', 'Months_of_Supply']],
            on='Model',
            suffixes=('_Short', '_Over')
        )

        dead_stock_df = merged_df[(merged_df['Jan_Jun_Sales'] == 0.1) & (merged_df['Current_Stock'] > 0)]

        traded_overstock = dealer_trades_df[['Location_Over', 'Model']].drop_duplicates()
        traded_overstock.rename(columns={'Location_Over': 'Location'}, inplace=True)

        unmatched_overstock_df = pd.merge(overstock_df, traded_overstock, on=['Location', 'Model'], how='outer',
                                          indicator=True)
        unmatched_overstock_df = unmatched_overstock_df[unmatched_overstock_df['_merge'] == 'left_only']

        dealer_trades_df = dealer_trades_df[dealer_trades_df['Location_Short'] != dealer_trades_df['Location_Over']]

        restock_df = merged_df[merged_df['Months_of_Supply'] < RESTOCK_THRESHOLD].copy()

        TARGET_SUPPLY = 3.0
        restock_df['Suggested_Order'] = (
                    (TARGET_SUPPLY * restock_df['Monthly_Run_Rate']) - restock_df['Current_Stock']).round().astype(int)

        # 5. Build PDF
        pdf_file = create_pdf_report(pacing_warnings_df, restock_df, dealer_trades_df, dead_stock_df,
                                     unmatched_overstock_df)

    except Exception as e:
        with open("error_log.txt", "w") as f:
            f.write(f"CRASH: {str(e)}\nTarget File: {target_file}\nCWD: {os.getcwd()}")
        update_db_status(task_id, "Error: Data Processing Failure")
        sys.exit()

    # 6. Normal Email Dispatch
    if not recipients:
        update_db_status(task_id, "Error: No Recipients")
        sys.exit()

    msg = EmailMessage()
    msg.set_content("Please find the attached automated dealership intelligence report.")
    msg["Subject"] = "Action Required: Dealership Intelligence Report"
    msg["From"] = sender_email
    msg["To"] = ", ".join(recipients)

    try:
        with open(pdf_file, "rb") as f:
            file_data = f.read()
            file_name = os.path.basename(pdf_file)
        msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=file_name)
    except Exception as e:
        update_db_status(task_id, "Error: PDF Attachment Failed")
        sys.exit()

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, app_password)
            server.send_message(msg)
        update_db_status(task_id, "Success")
    except Exception as e:
        update_db_status(task_id, "Error: SMTP Failure")
        sys.exit()

    with open(os.path.join(BASE_DIR, "debug_log.txt"), "a") as f:
        f.write("Dispatcher finished. Processing complete.\n")


if __name__ == "__main__":
    main()