# Executive Report Builder

> Python-based automation tool that generates standardized executive PowerPoint project status reports from structured JSON data.

Eliminates repetitive manual editing, ensures consistent formatting, and dynamically generates **one slide per project** using a predefined PowerPoint template.

---

## Table of Contents

- [Overview](#overview)
- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Template Configuration](#template-configuration)
- [JSON Structure](#json-structure)
- [How To Use](#how-to-use)
- [How It Works](#how-it-works)
- [Configuration](#configuration)
- [Suggested Weekly Workflow](#suggested-weekly-workflow)
- [Future Improvements](#future-improvements)

---

## Overview

This tool reads structured project data from a JSON file and automatically populates a PowerPoint template with:

| Field | Description |
|---|---|
| Project title & scope | Name and description of the project |
| Assigned user & area | Owner and business unit |
| Progress & days | Completion percentage and elapsed days |
| Observations & risks | Current notes and blockers |
| Milestones | Name, date, progress note, and status per milestone |

Each project in the JSON file becomes one slide in the final presentation.

---

## Project Structure

```
executive-report-builder/
│
├── template/
│   └── report_template.pptx
│
├── data/
│   └── example.json
│
├── output/
│
├── generate.py
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Requirements

- Python 3.10+
- A properly configured PowerPoint template (see [Template Configuration](#template-configuration))
- `python-pptx` library

```bash
pip install -r requirements.txt
```

---

## Template Configuration

The PowerPoint template must contain **named shapes**. To rename them:

> PowerPoint → Home → Select → Selection Pane → rename each text box

### Required Main Fields

| Shape Name | Description |
|---|---|
| `Project_title` | Project name |
| `project_scope_details` | Scope description |
| `project_user` | Assigned user |
| `project_area` | Business area |
| `report_date` | Report date |
| `progress_percentage_` | Completion % |
| `days_total_` | Total days |
| `observations_` | Observations text |
| `risk_` | Risk text |

### Milestone Rows

For each milestone row (up to `MAX_MILESTONES`), add these four shapes:

| Shape Name | Description |
|---|---|
| `milestone_N` | Milestone name |
| `date_logN` | Milestone date |
| `status_logN` | Progress note |
| `statusN` | Status label |

Where `N` is the row number (1, 2, 3...).

> ⚠️ Shape names are **case-sensitive** and must match exactly.

---

## JSON Structure

Create a JSON file inside the `data/` folder and populate it with your project data.

```json
{
  "report_date": "2026-02-20",
  "projects": [
    {
      "title": "RPA Ventas Crédito/SAP",
      "scope": "Automatización mediante RPA del proceso manual en SAP.",
      "user": "Matías Palermo",
      "area": "Cobranzas",
      "progress_percent": 20,
      "days_total": 8,
      "observations": [
        "Pendiente validación final",
        "Ajustes menores en documento"
      ],
      "risks": [
        "Dependencia de aprobación del usuario"
      ],
      "milestones": [
        {
          "milestone": "Levantamiento inicial",
          "date_log": "FEB 11",
          "status_log": "Estado: 90%",
          "status": "En Proceso"
        }
      ]
    }
  ]
}
```

> Each object in the `projects` array generates one slide.

---

## How To Use

### 1. Clone the Repository

```bash
git clone https://github.com/SheenyxX/executive-report-builder.git
cd executive-report-builder
```

### 2. Create a Virtual Environment

**Windows (PowerShell)**
```powershell
python -m venv venv
venv\Scripts\activate
```

**Mac / Linux**
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Create Your JSON File

Add a new file inside `data/`:

```
data/my_report.json
```

Fill it with your project information following the [JSON structure](#json-structure) above.

### 5. Generate the Report

```bash
python generate.py data/my_report.json
```

The final `.pptx` file will be saved inside the `output/` folder.

---

## How It Works

1. Loads the PowerPoint template
2. Reads structured project data from the JSON file
3. Fills each template shape by name
4. Duplicates the template slide for each project
5. Applies font adjustments to prevent text overflow
6. Saves the final presentation to `output/`

---

## Configuration

Inside `generate.py`, adjust the maximum number of milestone rows to match your template:

```python
MAX_MILESTONES = 5
```

---

## Suggested Weekly Workflow

1. Update the JSON file with current project status
2. Run the generator
3. Review the output `.pptx`
4. Share with stakeholders
5. Commit the JSON update to GitHub to maintain version history

---

## Future Improvements

- [ ] Automatic timestamped output filenames
- [ ] Multi-user report consolidation
- [ ] PDF export
- [ ] CLI argument enhancements
- [ ] GitHub Actions automation

---

*Executive Report Builder transforms manual executive reporting into a structured, automated workflow.*
