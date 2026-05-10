#Email Automation System

## Overview

This project automates the workflow of generating and managing field assessment communication emails for Swachh Survekshan 2025 operations.

The automation system:

* Reads master ULB tracking data
* Reads daily fieldwork plans
* Identifies fresh ULBs requiring intimation emails
* Updates the master tracking sheet automatically
* Creates Outlook draft emails for:

  * State authorities
  * ULB authorities
* Adds PDF attachments automatically
* Maintains version-controlled master files
* Generates execution logs for audit tracking

The system is designed to reduce manual effort, eliminate repetitive tasks, and improve operational efficiency during large-scale assessment operations.

---

# Key Features

## Data Processing

* Reads latest Master Excel automatically
* Reads latest FW Plan automatically
* Cleans and standardizes ULB codes
* Filters only fresh ULBs with scheduled fieldwork dates

---

## Workflow Automation

* Creates State-level Outlook draft emails
* Creates ULB-level Outlook draft emails
* Populates:

  * TO
  * CC
  * BCC
  * Subject
  * HTML Email Body
* Adds PDF authorization attachment automatically

---

## Master File Management

* Updates:

  * Intimation Email Sent? column
  * FW Start Date (Scheduled)
  * Timestamp column
* Saves updated master with automatic versioning

---

## Logging & Audit Trail

* Creates execution logs
* Maintains run summaries
* Tracks timestamps of operations
* Ensures reproducibility and operational transparency

---

# Technologies Used

| Technology             | Purpose                        |
| ---------------------- | ------------------------------ |
| Python                 | Core automation logic          |
| Pandas                 | Excel processing and filtering |
| pywin32                | Outlook draft automation       |
| OpenPyXL               | Excel handling                 |
| Outlook COM Automation | Draft email generation         |

---

# Project Structure

```text
Intimation Mails/
│
├── 1. Master/
│   ├── Master ULB List_v1.xlsx
│   ├── Master ULB List_v2.xlsx
│
├── 2. FW Plan/
│   ├── FW_Plan.xlsx
│
├── 3. Outputs/
│   ├── Emails/
│   └── Logs/
│
├── 4. Scripts/
│   └── mail_automation.py
│
├── 5. Attachment/
│   └── Authorization Letter.pdf
│
├── requirements.txt
├── README.md
└── .gitignore
```

---

# Installation

## Clone Repository

```bash
git clone https://github.com/iamshahbaz96/Email-Automation-System.git
```

---

## Install Dependencies

```bash
pip install -r requirements.txt
```

---

# requirements.txt

```text
pandas
pywin32
openpyxl
```

---

# Setup Instructions

## Step 1 — Add Master File

Place the latest master file inside:

```text
1. Master/
```

Example:

```text
Master ULB List_v1.xlsx
```

---

## Step 2 — Add FW Plan

Place FW plan inside:

```text
2. FW Plan/
```

Example:

```text
FW_Plan.xlsx
```

---

## Step 3 — Add Attachment

Place authorization PDF inside:

```text
5. Attachment/
```

Example:

```text
Authorization Letter.pdf
```

---

# Running the Automation

Run the script:

```bash
python mail_automation.py
```

---

# What Happens During Execution

The automation performs the following steps:

1. Reads latest master file
2. Reads latest FW Plan
3. Identifies fresh ULBs
4. Updates master tracking sheet
5. Creates Outlook draft emails
6. Adds recipients and attachments
7. Opens drafts for manual review
8. Saves updated master file with next version
9. Generates execution logs

---

# Important Note

The script DOES NOT automatically send emails.

It only:

* Creates Outlook drafts
* Opens them for review
* Waits for manual sending

This prevents accidental email dispatch.

---

# Sample Automation Workflow

```text
FW Plan Received
        ↓
Python Processing
        ↓
ULB Filtering
        ↓
Master Sheet Update
        ↓
Outlook Draft Creation
        ↓
Manual Review
        ↓
Send Emails
```

---

# Example Outputs

## Outlook Drafts

* State-level communication drafts
* ULB-level communication drafts

---

## Updated Master File

```text
Master ULB List_v2.xlsx
```

---

## Execution Logs

```text
3. Outputs/
└── Logs/
    └── 10-May-2026/
        └── Run_Summary_v2.txt
```

---

# Resume Highlights

* Built a Python-based workflow automation system for government field assessment communication operations.
* Automated Outlook draft generation for state and ULB-level stakeholders using Outlook COM automation.
* Developed version-controlled Excel tracking workflows with automated timestamps and audit logging.
* Reduced manual operational effort through automated recipient mapping, HTML email generation, and attachment handling.
* Implemented scalable folder-agnostic architecture enabling reusable deployment across teams.

---

# Future Improvements

Potential future enhancements:

* Auto-send capability
* GUI dashboard
* Scheduler integration
* Email delivery tracking
* Power BI integration
* Database-backed storage
* Cloud deployment
* Retry/error handling engine
* Dynamic HTML templates
* Multi-state scalability

---

# License

MIT License

---

# Author

Shahbaz Ahmed

Python Automation | Data Analytics | Workflow Automation | Operations Intelligence
