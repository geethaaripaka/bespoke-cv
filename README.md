# 📄 Bespoke CV

> Tailor your CV to any job in 30 seconds — not a rewrite, a surgical edit.

## What It Does
Bespoke CV reads a job description and your CV, 
identifies only the weak or misaligned bullet points, 
and rewrites just those. Everything strong stays untouched.

## Why It's Different
Most CV tools rewrite everything. Bespoke CV only 
touches what needs fixing — you stay in control.

| Feature | Bespoke CV | ResumeUp / Others |
|---|---|---|
| Surgical bullet rewrites | ✅ | ❌ |
| Side-by-side diff view | ✅ | ❌ |
| Accept/reject each change | ✅ | ❌ |
| Live match score | ✅ | ❌ |
| ATS compatibility score | ✅ | ❌ |
| Format preserved exactly | ✅ | ❌ |

## How It Works
1. Paste a LinkedIn, Naukri, or company career page URL
2. Upload your CV as Word Doc
3. Gemini AI identifies weak bullets and rewrites only those
4. Accept or reject each suggestion — score updates live
5. Download your updated CV as Word (.docx)

## Tech Stack
- Python + Streamlit
- Google Gemini 2.5 Flash
- pdfplumber — CV text extraction
- python-docx — Word document generation
- BeautifulSoup — JD fetching

## Run Locally
pip install -r requirements.txt
streamlit run app.py

## Get Your Free Gemini API Key
aistudio.google.com/app/apikey — free, no credit card needed

## Built By
Geetha Sushma — Product Manager
IIM Kozhikode MBA | Payments & Fintech
github.com/geethaaripaka
