```markdown
# Executive_Report_Builder

Executive_Report_Builder is a Python-based automation tool that generates standardized executive PowerPoint project status reports from structured JSON data.

It eliminates repetitive manual editing, ensures consistent formatting, and dynamically generates one slide per project using a predefined PowerPoint template.

---

## Overview

This tool reads structured project data from a JSON file and automatically populates a PowerPoint template with:

- Project title and scope  
- Assigned user and area  
- Progress percentage and total days  
- Observations and risks  
- Milestones (name, date, progress note, status)

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

````

---

## Requirements

- Python 3.10+ recommended  
- A properly configured PowerPoint template  
- python-pptx library  

Install dependencies:

```bash
pip install -r requirements.txt
````

---

## Template Configuration (Important)

The PowerPoint template must contain named shapes.

Open PowerPoint → Home → Select → Selection Pane
Rename the text boxes exactly as follows:

### Required Main Fields

* `Project_title`
* `project_scope_details`
* `project_user`
* `project_area`
* `report_date`
* `progress_percentage_`
* `days_total_`
* `observations_`
* `risk_`

### Milestone Rows

If `MAX_MILESTONES = 5` in `generate.py`, the template must include:

Row 1:

* `milestone_1`
* `date_log1`
* `status_log1`
* `status1`

Row 2:

* `milestone_2`
* `date_log2`
* `status_log2`
* `status2`

And so on, up to the configured `MAX_MILESTONES`.

Shape names are case-sensitive and must match exactly.

---

## JSON Structure

Create a JSON file inside the `data/` folder.

Example:

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

Each project in the `projects` array generates one slide.

---

## How To Use (New User Guide)

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

**Mac/Linux**

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Create Your JSON File

Inside the `data/` folder:

```
data/my_report.json
```

Edit it with your project information.

### 5. Generate the Report

```bash
python generate.py data/my_report.json
```

The final PowerPoint file will be generated inside:

```
output/
```

---

## How It Works Internally

1. Loads the PowerPoint template.
2. Reads structured project data from JSON.
3. Fills template shapes by name.
4. Duplicates the template slide for each project.
5. Applies font adjustments to prevent overflow.
6. Saves the final presentation in the output folder.

---

## Configuration

Inside `generate.py`, you can modify:

```python
MAX_MILESTONES = 5
```

Adjust this number if your template contains more milestone rows.

---

## Suggested Weekly Workflow

1. Update the JSON file with current project status.
2. Run the generator.
3. Review the output PPT.
4. Share with stakeholders.
5. Commit JSON updates to GitHub to maintain history.

---

## Future Improvements

* Automatic timestamped filenames
* Multi-user report consolidation
* PDF export
* CLI argument enhancements
* GitHub Actions automation

Executive_Report_Builder transforms manual executive reporting into a structured, automated workflow.

```
```

