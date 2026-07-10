# Sand Healthcare FDE Assignment — Quarterly Health Bulletin

Django prototype that automates Rwanda MoH **Quarterly Health Bulletin** compilation from DHIS2-like facility data.

## What it does

- Reads sample DHIS2-style CSV exports (`data/dhis2_facility_indicators.csv`)
- Calculates bulletin metrics:
  1. **Top 10 facilities** by patient volume
  2. **Maternal health indicators** (ANC visits, deliveries, complications, complication rate)
  3. **Facility performance scores** (timeliness + completeness)
  4. **Quarter-over-quarter trend analysis**
- Serves results as:
  - HTML dashboard (`/`)
  - JSON API (`/api/bulletin/`)
  - Downloadable JSON export (`/export/`)

## Setup

```bash
cd sand-fde-assignment
python -m venv .venv

# Windows
.venv\Scripts\activate

pip install -r requirements.txt
python manage.py migrate
python manage.py test
python manage.py runserver
```

Open http://127.0.0.1:8000/ in your browser.

## Project structure

```
sand-fde-assignment/
├── data/                          # Sample DHIS2-like exports
├── bulletin/
│   ├── services.py                # Metric calculations
│   ├── views.py                   # Dashboard + API endpoints
│   └── templates/bulletin/        # HTML dashboard
├── health_bulletin/               # Django project settings
├── generate_submission_doc.py     # Builds Word submission document
└── requirements.txt
```

## Generate Word submission document

```bash
pip install python-docx
python generate_submission_doc.py
```

Output: `Sand_FDE_Assignment_Submission.docx` in the project root.

## Assumptions (rapid prototype)

- CSV files stand in for live DHIS2 API pulls (2–3 week lag still applies to source data).
- Completeness is proxied by mandatory numeric fields; production would use DHIS2 validation rules.
- Authentication, scheduling, and PDF layout are out of scope for this sprint.

## Next steps toward production

See `Sand_FDE_Assignment_Submission.docx` Deliverables 3 & 4 for hardening and handover plan.
