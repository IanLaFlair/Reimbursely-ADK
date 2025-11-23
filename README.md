# Reimbursely-ADK
AI-powered reimbursement email processing agent built using **IBM watsonx Orchestrate**, **Gmail API**, and **Google Cloud Vision OCR**.

Reimbursely automatically:
- Fetches reimbursement emails from Gmail  
- Parses reimbursement form PDFs  
- Extracts payment amounts from receipts (OCR)  
- Reconciles form items vs receipts  
- Summarizes weekly reimbursement status  
- Generates Excel summaries  

This project is designed to automate manual financial workflows and support finance teams in processing reimbursement requests faster and more accurately.

## 🚀 Features
### 1. Gmail Email Retrieval
- Automatically fetches reimbursement-related emails.
- Excludes *advance* (cash advance) submissions.
- Correctly identifies applicant submissions, ignoring reply chains.

### 2. PDF Form Parsing
- Reads Pituku-style reimbursement forms.
- Extracts submission date, item descriptions, quantities, price, subtotal, and bank info.

### 3. OCR Receipt Extraction
- Uses Google Cloud Vision to extract OCR text and detect payment totals.
- Supports images inside PDFs.

### 4. Smart Reconciliation Engine
Matches form items with receipt amounts:
- Detects mismatches  
- Flags missing receipts  
- Identifies unused receipts  
- Produces an overall status (`OK` / `MISMATCH`)

### 5. Weekly Summary Generation
- Generates human‑readable tables.
- Provides downloadable XLSX summary files.

---
## 🏗️ Architecture Overview
```
┌────────────────┐     ┌────────────────────────┐
│ User (Chat UI) │────▶│ watsonx Orchestrate AI │
└────────────────┘     └──────────▲─────────────┘
                                   │
                                   │ calls
                                   ▼
                       ┌────────────────────────┐
                       │   Python Tooling ADK   │
                       │ (tools/gmail_tool)     │
                       └────────────────────────┘
                          │     │          │
                          ▼     ▼          ▼
              ┌────────────┐  ┌──────────────┐  ┌────────────────────┐
              │ Gmail API   │  │ Vision OCR   │  │ Excel Generator     │
              └────────────┘  └──────────────┘  └────────────────────┘

```

## 📦 Repository Structure
```
Reimbursely-ADK/
│
├── agents/
│   └── reimbursely.yaml
│
├── tools/
│   └── gmail_tool/
│       ├── source/gmail_tools.py
│       ├── token.json
│       ├── credentials.json
│       ├── vision_api_key.txt
│       └── requirements.txt
│
├── gmail_quickstart.py
└── README.md
```

---

## 🔧 Setup Instructions

### 1. Clone the Repository
```
git clone https://github.com/IanLaFlair/Reimbursely-ADK.git
cd Reimbursely-ADK
```

### 2. Create Virtual Environment
```
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies
```
pip install -r tools/gmail_tool/requirements.txt
```

### 4. Configure Gmail API
Place the following files inside `tools/gmail_tool/`:
- `credentials.json`  
- `token.json`  
- `vision_api_key.txt`  

Authenticate with Gmail:
```
python gmail_quickstart.py
```

---

## 🧠 IBM watsonx Orchestrate Tools

| Tool Name | Description |
|----------|-------------|
| `list_reimburse_emails_this_week` | Fetch reimbursement emails for the week |
| `parse_reimburse_form_from_email` | Parse reimbursement form PDF |
| `extract_all_payment_amounts_from_email` | OCR payment receipts |
| `analyze_reimburse_email` | Full form + receipt reconciliation |
| `export_reimburse_summary_this_week` | Generate XLSX weekly summary |

---

## 📄 Example Output (Summary Table)
```
ID: 19aa5600c1196c80  
Subject: 211125 - Pembayaran Biaya Kebersihan  
Status: MISMATCH  
Notes: Some items do not have a matching payment receipt.
```

---

## 🛡️ Security Notes
- OAuth tokens stay local; nothing is uploaded to the cloud.
- No sensitive credentials are included in the repository.
- Only read‑only Gmail access is used (`gmail.readonly`).

---

## 🤝 Contributing
Pull requests are welcome!  
For issues, open a GitHub Issue.

---

## 📜 License
MIT License  
Copyright (c) 2025