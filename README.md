# UAE Mall Search Dashboard

A Streamlit search and intelligence dashboard built from `2gis_all_malls_company_details.xlsx`.

## Features

- Search by mall, company, category, owner, operator or location
- Filter by emirate, size, company count and ownership confidence
- Mall profile with size, owner, operator, floors, parking and company directory
- Search across all company records
- Compare up to eight malls
- Download filtered results as CSV
- Data-quality coverage view
- Upload a replacement workbook from the sidebar

The workbook must contain sheets named `Malls` and `Companies`. Add an `Operator` column to `Malls` whenever operator research is available; the app detects it automatically.

## Run locally

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

macOS/Linux:

```bash
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Render

1. Create a new GitHub repository and upload every file in this folder, including the `data` folder.
   Confirm that GitHub shows `data/2gis_all_malls_company_details.xlsx` before deploying.
2. In Render, select **New → Blueprint**.
3. Connect the GitHub repository.
4. Render reads `render.yaml`; approve the web service and deploy.
5. After deployment, open the generated `onrender.com` URL.

No database or environment variable is required for this Excel-backed version.

## If Render says the workbook is missing

Render only receives files committed to GitHub. In the GitHub repository, select
**Add file → Upload files**, upload `data/2gis_all_malls_company_details.xlsx`,
commit the change, and redeploy. The app also accepts the workbook in the project
root and allows a temporary upload through its sidebar.
