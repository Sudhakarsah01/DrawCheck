# DrawCheck AI — QA Annotation Extractor
**SG Design Nepal | v3.0 Professional**

## What This Tool Does
Extracts all QA markup annotations from HVAC engineering drawing PDFs
and exports them to a professional Excel report.

## Folder Structure
```
QA_tool/
├── app.py                  ← Main app (all code inside)
├── requirements.txt        ← Python libraries
├── logo_transparent.png    ← SG Design Nepal logo
├── README.md               ← This file
└── .streamlit/
    └── config.toml         ← Port + theme settings
```

## Installation

### Step 1 — Install Python 3.11+
Download from python.org

### Step 2 — Install libraries
```bash
pip install -r requirements.txt
```

### Step 3 — Run
```bash
streamlit run app.py
```
Opens at: http://localhost:5000

## Deployment (Streamlit Cloud — Free)
1. Push this folder to GitHub
2. Go to share.streamlit.io
3. Connect your GitHub repo
4. Set main file: app.py
5. Deploy

## How to Update
Make changes to app.py, then:
```bash
git add .
git commit -m "your update description"
git push
```
Streamlit Cloud redeploys automatically.

## Excel Report Structure
- Tab 1: SUMMARY — totals, per-sheet breakdown, category breakdown
- Tab 2: ALL MARKUPS — every markup from all pages in one sheet
- Tab 3+: Sheet 1, Sheet 2... — one tab per drawing page

## Support
SG Design Nepal — Driving Innovation For a Better Tomorrow
