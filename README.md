# Pay-Fixation-Agent
This app has been developed by SC Rehman Gul for personal use only and to empower the troops to calculate their pay fixation on their own through this app without consulting to ministerial or clerkal staff   

## Overview

The **FC Pay Fixation Agent** is a Python-based application designed to assist with the preparation and review of Frontier Corps (FC) personnel pay fixation records. It aims to simplify historical pay calculations by organizing service events, pay revisions, annual increments, remustering, and promotions into a chronological pay fixation statement.

The application provides a user-friendly interface built with **Streamlit** and supports exporting pay fixation records into Excel, Word, and PDF formats.

## Key Features

- **Chronological Pay History:** Organizes pay fixation events from enrollment to the selected end date.
- **Pay Revision Support:** Applies configured pay scale revision rules.
- **Annual Increment Tracking:** Calculates increments according to the rules currently implemented in the application.
- **Remustering and Promotion:** Records service changes and performs configured pay fixation calculations.
- **Multiple Export Formats:** Generates downloadable Excel (`.xlsx`), Word (`.docx`), and PDF (`.pdf`) reports.
- **Interactive Interface:** Provides a simple interface for entering personnel details and service dates.
- **Review and Verification:** Supports the review of calculated pay records against applicable official orders and reference documents.

## Technology Stack
- Python
- Streamlit
- Pandas
- OpenPyXL
- Python-Docx
- ReportLab

## Installation and Setup

### 1. Clone the Repository

```bash
git clone <YOUR_GITHUB_REPOSITORY_URL>
cd "Pay Fixation Agent"
```
Replace `<YOUR_GITHUB_REPOSITORY_URL>` with the HTTPS URL of this GitHub repository.

### 2. Create a Virtual Environment

```bash
python -m venv .venv
```

Activate it on Windows:

```bat
.venv\Scripts\activate
```

### 3. Install Dependencies

```bash
python -m pip install --upgrade pip
python -m pip install streamlit pandas openpyxl python-docx reportlab
```

### 4. Run the Application

Ensure that the main Python file is named `app.py`, then run:

```bash
python -m streamlit run app.py
```

Open the local URL displayed in the terminal, usually `http://localhost:8501`.

## Deployment

The application can be deployed on Streamlit Community Cloud using this GitHub repository. The repository should contain the main application file (`app.py`) and a `requirements.txt` file listing the required Python packages.

## Important Notice

This project is intended as an administrative assistance and calculation-review tool. Its calculations depend on the pay scales, formulas, dates, and rules configured in the application. All results should be verified against applicable official Frontier Corps orders, pay fixation instructions, and individual service records before being used for official financial or pension purposes.

The application should not be considered officially certified or an authoritative source of pay fixation unless formally validated and approved by the competent authority.

## Future Improvements

- Import pay scales and historical rules directly from reference workbooks.
- Expand support for all applicable ranks and historical pay revisions.
- Improve validation of promotions, remustering, and annual increment eligibility.
- Add detailed audit trails and calculation explanations.
- Improve input validation and error reporting.
- Strengthen access controls and protection of personnel data.

## License

A license has not yet been specified. All rights and permissions remain subject to the repository owner's decision until a license is added.
