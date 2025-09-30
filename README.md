# 🦟 Malaria Apps (Cleaning & Indicators)

Two Streamlit apps:
- `Malaria_Data_Cleaning.py` — validates, cleans, flags duplicates & consistency issues.
- `Malaria_Indicator.py` — computes program indicators from cleaned data.

---

## 1) Get the code locally
Option A: Download this folder as a zip (e.g., via ChatGPT or GitHub) and unzip it.  
Option B: `git clone` once you create the GitHub repo (see step 4).

Recommended: **Python 3.10+**.

## 2) Create & activate a virtual environment
**macOS / Linux**
```bash
python3 -m venv .venv
source .venv/bin/activate
```

**Windows (PowerShell)**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

## 3) Install dependencies
```bash
pip install -r requirements.txt
```

## 4) Run the apps locally
**Cleaning app**
```bash
streamlit run Malaria_Data_Cleaning.py
```
**Indicator app**
```bash
streamlit run Malaria_Indicator.py
```

## 5) Initialize a Git repository
```bash
git init
git add .
git commit -m "Initial commit: Malaria Data Cleaning & Indicator apps"
```

## 6) Create a new GitHub repository
- Visit https://github.com → **New repository**.
- Name it, e.g. `malaria-apps`. Choose Public or Private.
- **Do not** add an initial README (we already have one).

Then connect & push:
```bash
git branch -M main
git remote add origin https://github.com/<your-username>/malaria-apps.git
git push -u origin main
```

## 7) Share the code
Share your GitHub repository link with collaborators.

## 8) Optional: Deploy to Streamlit Community Cloud
1. Go to https://streamlit.io/cloud and sign in.
2. **New app** → connect your GitHub repo.
3. Choose Branch: `main` and App file:
   - `Malaria_Data_Cleaning.py` (deploy first app)
   - Repeat for `Malaria_Indicator.py` (second app)
4. Streamlit will detect `requirements.txt` and install packages.
5. Copy the public URLs and share them.

---

### Notes
- Excel support: `.xlsx` via `openpyxl`; legacy `.xls` via `xlrd==1.2.0`.
- The apps output downloadable Excel workbooks and include UI elements for selecting sheets.
