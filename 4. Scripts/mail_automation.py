import pandas as pd
import win32com.client as win32
from datetime import datetime
import os
import glob

# =========================================================
# PROJECT ROOT
# =========================================================
# FINAL DIRECTORY STRUCTURE:
#
# Intimation Mails/
# │
# ├── 1. Master/
# │   ├── Master ULB List_v1.xlsx
# │   ├── Master ULB List_v2.xlsx
# │
# ├── 2. FW Plan/
# │   ├── FW_Plan.xlsx
# │
# ├── 3. Outputs/
# │   ├── Emails/
# │   └── Logs/
# │
# ├── 4. Scripts/
# │   └── mail_automation.py
# │
# ├── 5. Attachment/
# │   └── Authorization Letter.pdf
# =========================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)

# =========================================================
# FOLDER PATHS
# =========================================================

MASTER_FOLDER = os.path.join(
    PROJECT_ROOT,
    "1. Master"
)

FW_PLAN_FOLDER = os.path.join(
    PROJECT_ROOT,
    "2. FW Plan"
)

OUTPUTS_FOLDER = os.path.join(
    PROJECT_ROOT,
    "3. Outputs"
)

LOG_OUTPUT_BASE = os.path.join(
    OUTPUTS_FOLDER,
    "Logs"
)

ATTACHMENT_FOLDER = os.path.join(
    PROJECT_ROOT,
    "5. Attachment"
)

ATTACHMENT_FILE = os.path.join(
    ATTACHMENT_FOLDER,
    "Authorization Letter.pdf"
)

# =========================================================
# CHECK ATTACHMENT EXISTS
# =========================================================

if not os.path.exists(ATTACHMENT_FILE):

    raise Exception(
        f"Attachment file not found:\n{ATTACHMENT_FILE}"
    )

# =========================================================
# CREATE DATE-WISE LOG FOLDER
# =========================================================

run_date = datetime.now().strftime("%d-%b-%Y")

LOG_FOLDER = os.path.join(
    LOG_OUTPUT_BASE,
    run_date
)

os.makedirs(LOG_FOLDER, exist_ok=True)

print(f"\nLog Folder: {LOG_FOLDER}")

# =========================================================
# PICK LATEST MASTER FILE
# =========================================================

master_files = glob.glob(
    os.path.join(MASTER_FOLDER, "Master ULB List_v*.xlsx")
)

if not master_files:
    raise Exception("No master version files found.")

MASTER_FILE = max(
    master_files,
    key=os.path.getctime
)

print(f"\nUsing Master File:")
print(MASTER_FILE)

# =========================================================
# PICK FW PLAN FILE
# =========================================================

fw_plan_files = glob.glob(
    os.path.join(FW_PLAN_FOLDER, "FW_Plan*.xlsx")
)

if not fw_plan_files:
    raise Exception("No FW Plan file found.")

FW_PLAN_FILE = max(
    fw_plan_files,
    key=os.path.getctime
)

print(f"\nUsing FW Plan File:")
print(FW_PLAN_FILE)

# =========================================================
# READ FILES
# =========================================================

master_df = pd.read_excel(MASTER_FILE)

daily_df = pd.read_excel(FW_PLAN_FILE)

# =========================================================
# CLEAN DATA
# =========================================================

master_df['ULB_Code'] = (
    master_df['ULB_Code']
    .astype(str)
    .str.strip()
    .str.replace('.0', '', regex=False)
)

daily_df['ULB_Code'] = (
    daily_df['ULB_Code']
    .astype(str)
    .str.strip()
    .str.replace('.0', '', regex=False)
)

daily_df['FW Start Date'] = pd.to_datetime(
    daily_df['FW Start Date'],
    errors='coerce'
)

# =========================================================
# KEEP REQUIRED FW PLAN COLUMNS
# =========================================================

daily_df = daily_df[
    ['ULB_Code', 'FW Start Date']
]

# =========================================================
# MERGE
# =========================================================

merged = pd.merge(
    master_df,
    daily_df,
    on='ULB_Code',
    how='inner'
)

# =========================================================
# FILTER FRESH ULBs
# =========================================================

fresh_ulbs = merged[
    (
        merged['Intimation Email Sent?\n(Yes/No)']
        .fillna('')
        .astype(str)
        .str.strip()
        .str.lower()
        .isin(['', 'no'])
    )
    &
    (merged['FW Start Date'].notna())
]

print(f"\nFresh ULBs Found: {len(fresh_ulbs)}")

if len(fresh_ulbs) == 0:

    print("\nNo fresh ULBs found.")

    exit()

# =========================================================
# CREATE TIMESTAMP COLUMN IF NOT EXISTS
# =========================================================

timestamp_col = "Intimation Email Sent Timestamp"

if timestamp_col not in master_df.columns:

    master_df[timestamp_col] = ""

# =========================================================
# UPDATE MASTER FILE
# =========================================================

current_timestamp = datetime.now().strftime(
    "%d-%b-%Y %H:%M:%S"
)

for _, row in fresh_ulbs.iterrows():

    ulb_code = row['ULB_Code']

    fw_date = row['FW Start Date']

    # =====================================================
    # UPDATE EMAIL STATUS
    # =====================================================

    master_df.loc[
        master_df['ULB_Code'] == ulb_code,
        'Intimation Email Sent?\n(Yes/No)'
    ] = 'Yes'

    # =====================================================
    # UPDATE FW DATE
    # =====================================================

    master_df.loc[
        master_df['ULB_Code'] == ulb_code,
        'FW Start Date\n(Scheduled)'
    ] = fw_date.strftime('%d-%b-%Y')

    # =====================================================
    # UPDATE TIMESTAMP
    # =====================================================

    master_df.loc[
        master_df['ULB_Code'] == ulb_code,
        timestamp_col
    ] = current_timestamp

# =========================================================
# GROUP DATA
# =========================================================

grouped = fresh_ulbs.groupby(
    ['StateName', 'FW Start Date']
)

# =========================================================
# OUTLOOK SETUP
# =========================================================

outlook = win32.Dispatch('outlook.application')

EMAIL_SUBJECT = (
    "SS2025: Intimation Regarding Field Assessment"
)

# =========================================================
# GENERATE STATE MAIL DRAFTS
# =========================================================

for (state, fw_date), group in grouped:

    # =====================================================
    # TO & CC
    # =====================================================

    to_emails = (
        group['State_Mail to be in To']
        .dropna()
        .astype(str)
        .unique()
    )

    cc_emails = (
        group['State_Mail to be in CC']
        .dropna()
        .astype(str)
        .unique()
    )

    to_string = ";".join(to_emails)

    cc_string = ";".join(cc_emails)

    # =====================================================
    # CREATE HTML TABLE
    # =====================================================

    table_html = """
    <table border="1" cellspacing="0" cellpadding="5"
    style="border-collapse: collapse; font-family: Arial;">

    <tr>
        <th>S.No</th>
        <th>ULB Code</th>
        <th>ULB Name</th>
        <th>District</th>
    </tr>
    """

    for i, (_, row) in enumerate(group.iterrows(), start=1):

        table_html += f"""
        <tr>
            <td>{i}</td>
            <td>{row['ULB_Code']}</td>
            <td>{row['ULBName']}</td>
            <td>{row['DistrictName']}</td>
        </tr>
        """

    table_html += "</table>"

    # =====================================================
    # HTML BODY
    # =====================================================

    body = f"""
    <html>
    <body>

    <p>Dear Sir/Madam,</p>

    <p>
    Warm greetings from Ipsos Research Private Limited.
    </p>

    <p>
    The Ministry of Housing and Urban Affairs (MoHUA),
    Government of India has engaged Ipsos Research Pvt. Ltd.
    (Ipsos) to conduct assessments under Swachh Survekshan
    2025-2026 in the ULBs of the State of
    <b>{state}</b>.
    </p>

    <p>
    The On-Field Validation assumes a pivotal role in ensuring
    a fair and meticulous evaluation process.
    </p>

    <p>
    As part of the survey process, we would like to inform
    you about initiating the field assessment of the following
    ULBs on <b>{fw_date.strftime('%d-%b-%Y')}</b>.
    </p>

    {table_html}

    <br>

    <p>
    We kindly request you to ensure that all facilities and
    schools are accessible on all seven days of the week,
    as the assessment will be conducted throughout the week.
    </p>

    <p>
    Kindly refer to the attached authorisation letter for
    more information.
    </p>

    <p>
    The above-mentioned ULBs have already been informed
    by mail.
    </p>

    <p>
    Further, we will continue to inform you about the
    upcoming assessments in the ULBs of your state.
    </p>

    <p>
    We would also humbly request your good office for
    intimation to the nodal officer(s) of the corresponding
    ULBs.
    </p>

    <p>
    Thanking you,
    <br><br>
    SS2025 Team
    <br>
    Ipsos Research Private Limited
    </p>

    </body>
    </html>
    """

    # =====================================================
    # CREATE OUTLOOK DRAFT
    # =====================================================

    mail = outlook.CreateItem(0)

    mail.To = to_string

    mail.CC = cc_string

    mail.Subject = EMAIL_SUBJECT

    mail.HTMLBody = body

    # =====================================================
    # ADD ATTACHMENT
    # =====================================================

    mail.Attachments.Add(ATTACHMENT_FILE)

    # =====================================================
    # OPEN DRAFT
    # =====================================================

    mail.Display()

    print(
        f"\nState Mail Draft Created:"
        f" {state} | {fw_date.strftime('%d-%b-%Y')}"
    )

# =========================================================
# GENERATE ULB MAIL DRAFTS
# =========================================================

for (state, fw_date), group in grouped:

    # =====================================================
    # CC & BCC
    # =====================================================

    cc_emails = (
        group['ULB_Mail to be in CC']
        .dropna()
        .astype(str)
        .unique()
    )

    bcc_emails = (
        group['ULB_Mail to be in BCC']
        .dropna()
        .astype(str)
        .unique()
    )

    cc_string = ";".join(cc_emails)

    bcc_string = ";".join(bcc_emails)

    # =====================================================
    # HTML BODY
    # =====================================================

    body = f"""
    <html>
    <body>

    <p>Dear Sir/Madam,</p>

    <p>
    Warm greetings from Ipsos Research Private Limited.
    </p>

    <p>
    The Ministry of Housing and Urban Affairs (MoHUA),
    Government of India has engaged Ipsos Research Pvt. Ltd.
    (Ipsos) to conduct assessments under Swachh Survekshan
    2025-2026 in the ULB of <b>{state}</b>.
    </p>

    <p>
    As part of the survey process, we would like to inform
    you about initiating the field assessment of your ULB
    on <b>{fw_date.strftime('%d-%b-%Y')}</b>.
    </p>

    <p>
    We kindly request you to ensure that all facilities and
    schools are accessible on all seven days of the week,
    as the assessment will be conducted throughout the week.
    </p>

    <p>
    Sanitation worker interviews will be conducted at the
    vehicle shed location.
    </p>

    <p>
    Kindly refer to the attached authorisation letter for
    more information.
    </p>

    <p>
    We appreciate your cooperation in this regard.
    </p>

    <p>
    Thanking you,
    <br><br>
    SS2025 Team
    <br>
    Ipsos Research Private Limited
    </p>

    </body>
    </html>
    """

    # =====================================================
    # CREATE OUTLOOK DRAFT
    # =====================================================

    mail = outlook.CreateItem(0)

    mail.CC = cc_string

    mail.BCC = bcc_string

    mail.Subject = EMAIL_SUBJECT

    mail.HTMLBody = body

    # =====================================================
    # ADD ATTACHMENT
    # =====================================================

    mail.Attachments.Add(ATTACHMENT_FILE)

    # =====================================================
    # OPEN DRAFT
    # =====================================================

    mail.Display()

    print(
        f"\nULB Mail Draft Created:"
        f" {state} | {fw_date.strftime('%d-%b-%Y')}"
    )

# =========================================================
# MASTER FILE VERSIONING
# =========================================================

existing_versions = glob.glob(
    os.path.join(MASTER_FOLDER, "Master ULB List_v*.xlsx")
)

version_numbers = []

for file in existing_versions:

    filename = os.path.basename(file)

    version = (
        filename
        .split("_v")[-1]
        .replace(".xlsx", "")
    )

    version_numbers.append(int(version))

next_master_version = max(version_numbers) + 1

new_master_file = os.path.join(
    MASTER_FOLDER,
    f"Master ULB List_v{next_master_version}.xlsx"
)

# =========================================================
# SAVE UPDATED MASTER
# =========================================================

master_df.to_excel(
    new_master_file,
    index=False
)

print(f"\nUpdated Master Saved:")
print(new_master_file)

# =========================================================
# GENERATE RUN SUMMARY
# =========================================================

summary_file = os.path.join(
    LOG_FOLDER,
    f"Run_Summary_v{next_master_version}.txt"
)

with open(summary_file, "w") as f:

    f.write("SS2025 EMAIL AUTOMATION RUN SUMMARY\n")
    f.write("=" * 60 + "\n\n")

    f.write(
        f"Run Timestamp: {current_timestamp}\n\n"
    )

    f.write(f"Master File Used:\n{MASTER_FILE}\n\n")

    f.write(f"New Master File:\n{new_master_file}\n\n")

    f.write(f"FW Plan File:\n{FW_PLAN_FILE}\n\n")

    f.write(f"Fresh ULBs Processed: {len(fresh_ulbs)}\n\n")

    f.write(
        f"States Processed: "
        f"{fresh_ulbs['StateName'].nunique()}\n\n"
    )

print(f"\nRun Summary Saved:")
print(summary_file)

print("\nAutomation Completed Successfully.")