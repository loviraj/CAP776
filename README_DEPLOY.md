# CAP776 Web Evaluator — Deployment Guide

## Architecture

Student Browser
    |
    v
Streamlit Web App
    |
    +--> Evaluator engine
    |      +--> XLSX validation and recalculation
    |      +--> DOCX report verification
    |      +--> optional PY AST inspection
    |      +--> marks + PDF feedback
    |
    +--> Supabase/PostgreSQL
           +--> one database row per submission
           +--> faculty-only Excel export

The student files are processed in memory and are NOT stored in the database.
The database stores evaluation results, metadata, and scores.

## 1. Create the database

Create a free Supabase project at https://supabase.com/

Open SQL Editor and run the contents of `schema.sql`.

Copy the PostgreSQL connection string from:
Project Settings -> Database -> Connection string.

Prefer the connection string with SSL enabled.

## 2. Put the project on GitHub

Repository contents:

    app.py
    requirements.txt
    schema.sql
    .streamlit/config.toml

Do NOT upload real student files or secrets.

## 3. Deploy to Streamlit Community Cloud

Go to https://share.streamlit.io/

Create a new app:
- Repository: your GitHub repository
- Branch: main
- Main file: app.py

Deploy.

## 4. Add secrets

In Streamlit Cloud:
App -> Settings -> Secrets

Paste:

    DATABASE_URL = "your-postgresql-connection-string"
    ADMIN_PASSWORD = "a-long-random-password"

Do not put these in app.py or GitHub.

## 5. Student workflow

Students open the web URL.

They upload:
- registration-number.xlsx
- registration-number.docx
- optional registration-number.py

Then press:

    SUBMIT FOR EVALUATION

The app:
1. Reads the XLSX.
2. Validates the data.
3. Restricts calculation to 17 Aug 2026–21 Sep 2026.
4. Recalculates all prescribed indices independently.
5. Checks report values.
6. Calculates the three prescribed correlations.
7. Inspects the optional Python library.
8. Calculates the rubric score.
9. Shows marks immediately.
10. Generates a student PDF.
11. Records the result in PostgreSQL.

## 6. Faculty workflow

Open the same URL.

At the bottom:
Faculty / Admin -> enter admin password -> Open Results Dashboard.

The dashboard shows all submissions and provides:

    Download All Results as Excel

## 7. Recommended operational rules

1. Ask students to use the registration number in BOTH filenames:
       12345.xlsx
       12345.docx
       12345.py

2. Keep the Python file optional exactly as requested.

3. The XLSX remains the numerical source of truth.

4. Do not store uploaded student files unless there is a specific requirement.

5. Before opening the submission URL to all students, test with:
       - one correct submission
       - one missing-date submission
       - one incorrect report value
       - one submission without .py
       - one malformed .py

6. Download the Excel results periodically as a backup.

## 8. Important scaling note

For a class of 500+ students, this architecture is much safer than writing to a local SQLite file on Streamlit Cloud.

Streamlit handles the user interface and evaluation.
PostgreSQL/Supabase handles persistent shared storage.

If the university later requires institutional authentication, replace the public submission page with university SSO or add an authentication layer without changing the evaluator engine.

## 9. Local testing

Install:

    pip install -r requirements.txt

Set:

    DATABASE_URL=...
    ADMIN_PASSWORD=...

Then:

    streamlit run app.py

For local testing, you can also put secrets in:

    .streamlit/secrets.toml

Never commit that file to GitHub.
