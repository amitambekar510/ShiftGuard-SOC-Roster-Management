<div align="center">

# 🛡️ ShiftGuard
### SOC Roster Management • Coverage Assurance • Analyst Well-being

**Build a shift roster. Inspect real coverage gaps. Export a clean Excel workbook.**

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge)
![Excel](https://img.shields.io/badge/Output-XLSX-217346?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-66b3a1?style=for-the-badge)

</div>

> [!IMPORTANT]
> **Coverage is evaluated by assigned shift**, not by counting active staff. A team can have analysts on its roster and still have an unstaffed night shift. This tool reports gaps; it cannot guarantee minimum coverage or labor-law compliance.

## Why ShiftGuard?
Security Operations Centers operate around the clock. ShiftGuard is a Python-based, local-first utility that creates illustrative or user-configured monthly rosters and checks whether each shift has the required number of analysts.

## ✨ Features

| Capability | What it does |
|---|---|
| 🗓️ Monthly schedules | Choose any month and year |
| 🔁 Shift rotations | Four-day (3-on/1-off) and five-day (3-on/2-off) patterns |
| 🛡️ Coverage validation | Checks **Morning, Afternoon, and Night** independently for each required role |
| ⚠️ Gap register | Dedicated sheet listing uncovered shifts and required headcount |
| 📈 Analyst statistics | Counts assigned mornings, afternoons, nights and rest days |
| 🧾 Excel delivery | Four worksheets with readable formatting and frozen headings |
| 🔐 Privacy | Uses synthetic example staff by default; your custom JSON stays local |

## 🚀 Run locally

```bash
git clone https://github.com/amitambekar510/ShiftGuard-SOC-Roster-Management.git
cd ShiftGuard-SOC-Roster-Management
python -m pip install -r requirements.txt
python shiftguard.py --year 2026 --month 10
```

Output: `SOC_Roster_Validated.xlsx`.

To flag insufficient coverage in automated runs:

```bash
python shiftguard.py --year 2026 --month 10 --min-l1 1 --min-l2 1 --fail-on-gaps
```

Exit status **2** means one or more role/shift coverage gaps were found; inspect the **Coverage Gaps** sheet. The workbook is still written.

### Configure your own team
Create a local JSON file (do **not** commit identifiable staff details to a public repository):

```json
[
  {"name": "Analyst 01", "role": "SOC L1", "pattern": "rotating", "offset": 0},
  {"name": "Analyst 02", "role": "SOC L2", "pattern": "rotating", "offset": 1},
  {"name": "Shift Lead", "role": "Team Lead", "pattern": "weekday"}
]
```

Use `--team team.json`. Optional member keys: `start` and `end` in `YYYY-MM-DD` format. All assignments anchor their cycle to **January 1, 2026**, so transitions between months do not reset the rotation. The year is configurable; holiday calendars and rest constraints are **not** automatically applied in v2.

## 📦 Workbook layout
- **SOC Roster:** names, roles, patterns, daily shift allocations.
- **Coverage Gaps:** every unmet minimum for each date, role and shift.
- **Shift Statistics:** per-person shift and off-day counts.
- **Read Me:** limitations and interpretation.

## 🧪 Tests
```bash
python -m unittest discover -s tests -v
```
GitHub Actions also runs the unit tests and exercises workbook generation.

## ⚠️ Limitations and responsible use
- ShiftGuard v2 does **not** ingest production HR systems or track approved leave, statutory holidays, skill eligibility, fatigue constraints, or overtime. Review these before adopting a roster.
- The output is a planning aid, not a legal compliance certification or a substitute for managerial approval.
- `generate_roster.py` is retained as a **legacy 2026 example**, including its hard-coded dates and historical logic. For corrected per-shift coverage checks and configurable monthly exports, use `shiftguard.py`.
- Do not publish personnel records, schedules, names, contact details or confidential employment data in this public project.

## 🔭 Roadmap
Holiday calendars by jurisdiction, leave constraints, configurable staffing/skill matrices, maximum consecutive nights, minimum rest periods, and a private local web interface.

## 👨‍💻 Maintainer
[Amit Ambekar](https://github.com/amitambekar510) · [MIT license](LICENSE.txt)
