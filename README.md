# Dental CRM

A local Streamlit CRM for dental laboratory orders.

## Features

- Admin and doctor authentication
- Doctor signup and admin approval
- Admin or doctor order creation
- STL and patient photo uploads
- Case status tracking
- PDF and editable DOCX invoice generation
- Payment recording and outstanding balance dashboard
- SQLite database with automatic setup

## Project structure

```text
Dental CRM/
|-- app.py
|-- dentalcrm.db                 # created on first run
|-- requirements.txt
|-- uploads/
|   |-- STL/
|   |-- Photos/
|   `-- Invoices/
`-- Invoice Templates/
```

## Run

```powershell
cd "Dental CRM"
python -m pip install -r requirements.txt
streamlit run app.py
```

Open the URL shown by Streamlit, normally `http://localhost:8501`.

## Initial admin account

- Email: `admin@dentalcrm.local`
- Password: `Admin@123`

Change the seeded password before using this app with real patient data. The app
is intended as a local starter and should be deployed behind HTTPS with managed
secrets, backups, access logs, and applicable healthcare privacy controls before
production use.
