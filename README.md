# CEGA Dashboard — SRLS

**MaCSS Team | May 2026**

Data available: https://dataverse.harvard.edu/dataverse/SRLS

The Center for Effective Global Action (CEGA) is primarily a research center that informs economic policies in developing countries. This project focuses on Syrian Refugee Life, which consists of a longitudinal panel data set with 5 years of data for 2,600 refugees. This data is comprehensive, with outcomes on employment, fertility, marriage, consumption, intentions to return, health, mental health, and more. We aim to make a user-friendly dashboard for policymakers and academics to use for their own information. Completed in partnership with CEGA, the MaCSS program, Ruben Gomez, & Hanning Zhou.

**Owner / contact:** [mkalra@berkeley.edu](mailto:mkalra@berkeley.edu)

---

## Repository Structure

```
CEGA-Dashboard/
├── Data/
│   ├── 2024 value codebook.xlsx
│   ├── Combined_survey_03_17.xlsx
│   ├── example.txt
│   └── processSurvey.py
├── dashboard_downloads/         # Tableau .twbx files
│   ├── Economics(CEGA).twbx
│   ├── Education(CEGA).twbx
│   ├── Food_Security [CEGA_SRLS].twbx
│   ├── Mental Health (CEGA).twbx
│   ├── Physical_Health [CEGA_SRLS].twbx
│   └── Safety (CEGA).twbx
├── Final Presentation (CEGA).pptx
└── README.md
```

---

## Dashboards

Public Tableau links (may not be active post-May 2026):

| Topic | Author | Link |
|---|---|---|
| Education | Hanning | https://public.tableau.com/app/profile/hanning.zhou/viz/CEGA_education_1/dashboard1_education#1 |
| Economics | Hanning | https://public.tableau.com/app/profile/hanning.zhou/viz/CEGA_eco/Dashboard1#1 |
| Food Security | Juan Rubén | https://public.tableau.com/app/profile/ruben.gomez8874/viz/Food_SecurityCEGA_SRLS/Dashboard_Food_Security |
| Health | Juan Rubén | https://public.tableau.com/app/profile/ruben.gomez8874/viz/Physical_HealthCEGA_SRLS/Physical_Health_Dashboard |
| Mental Health | Fuguan | https://public.tableau.com/app/profile/fuguan.obrien/viz/MentalHealthCEGA/Dashboard1 |
| Safety | Fuguan | https://public.tableau.com/app/profile/fuguan.obrien/viz/SafetyCEGA/Dashboard1 |

The packaged `.twbx` source files live in [`dashboard_downloads/`](./dashboard_downloads).

### Mental Health
<img width="982" height="784" alt="Mental Health Dashboard" src="https://github.com/user-attachments/assets/cea76919-64a2-48ec-861c-e772c0bf0239" />

### Safety
<img width="987" height="784" alt="Safety Dashboard" src="https://github.com/user-attachments/assets/d41ca996-ed08-416c-a1e3-b6a7e6d70c86" />

---

## Data Pipeline

1. **Convert `.dta` → `.xlsx`.** Use `exporting_to_excel.do` (located in the original Drive under `Data/ORIGINAL DATA/`) to convert Stata files to Excel. **If your data is already in Excel, skip this step** — but make sure column names and values are *actual values, not labels*.
2. **Run `processSurvey.py`** ([`Data/processSurvey.py`](./Data/processSurvey.py)). Ask Claude (or another LLM) to adjust it to add only one additional year of data. Keep the `2024 value codebook.xlsx` in the same folder while running.

### General notes on `processSurvey.py`
- All currency units are in **Dinar**.
- All time units are in **years** or **24-hour format** (respectively).
- The following adjustments were made manually after the script's output:
  - "Why not registered with UNHCR" was manually coded.
  - "Other (specify)" was collapsed to `"Other"`.
  - `Highest_Education` — `2.5` manually decoded to `1st grade`.
  - `Length_Schooling_Total_Years` — changed to a proxy (see Education section below).
  - `Why_Not_Looking_Job` — manually coded for 2020 (`q10_2_16_other_c1`).

---

## Post-Script Transformations (in Excel)

### Migration tab — `Why_Left_Syria` → 9 categories

```excel
=IFS(
    AND(ISNUMBER(SEARCH("Violence", B2)), ISNUMBER(SEARCH(",", B2)), OR(ISNUMBER(SEARCH("Work", B2)), ISNUMBER(SEARCH("school", B2)))), "Violence + Work/Edu",
    AND(ISNUMBER(SEARCH("Violence", B2)), ISNUMBER(SEARCH(",", B2))), "Violence + Personal/Other",
    ISNUMBER(SEARCH("Violence", B2)), "Violence Only",
    AND(ISNUMBER(SEARCH("unsafe", B2)), ISNUMBER(SEARCH(",", B2))), "Unsafe Area + Mixed",
    ISNUMBER(SEARCH("unsafe", B2)), "Unsafe Area Only",
    OR(ISNUMBER(SEARCH("Work", B2)), ISNUMBER(SEARCH("school", B2))), "Work/Education",
    OR(ISNUMBER(SEARCH("Family", B2)), ISNUMBER(SEARCH("Marriage", B2))), "Family/Marriage",
    ISNUMBER(SEARCH("Health", B2)), "Health",
    ISNUMBER(SEARCH("Vacation", B2)), "Vacation",
    TRUE, "Other"
)
```

**Why this approach.** A nested `IFS` formula was applied to the `=UNIQUE()` list to consolidate responses into 9 categories based on a hierarchy of primary drivers. To prevent acute factors like "Violence" or "Unsafe area" from masking multi-factor responses, the formula detects comma delimiters to separate single-reason moves from mixed-motive relocations (e.g., violence combined with economic needs). If stricter statistical granularity is required later, alternative transformations include splitting the data into binary (1/0) indicator columns for each distinct reason, or using Power Query to unpivot the comma-separated strings into individual rows.

### Education tab

`Length_Schooling_Total_Years` is **estimated** (uses only the most recent years rather than all years):

```excel
=IFS(
    D2="Never been in school",0,
    D2="Literacy",1,
    D2="Bible / Qur`an school",2,
    D2="Primary school",6,
    D2="Secondary school",12,
    D2="Vocational training",13,
    D2="College",14,
    D2="University",16,
    D2="Institution",10
)
```

Column renamed: `Highest_Edu_Categorized`.

`Education_Level_Binned` (where `L2` is `Highest_Edu_Categorized`; pull formula down to all filled values):

```excel
=IFS(
    L2="No Formal Education","0 years",
    L2="Early Childhood","1–6 years",
    L2="Primary","1–6 years",
    L2="Lower Secondary","7–12 years",
    L2="Upper Secondary","7–12 years",
    L2="Post-Secondary","13–16 years",
    L2="Higher Education","13–16 years"
)
```

### Mental Health tab — `CESD Score`

New column scoring the 10-item CES-D (range 0–30; higher = more depressive symptoms). Columns `C2:L2` correspond to: Bothered, Unconcentrated, Depressed, Without Energy, **Hopeful (reverse scored)**, Afraid, Trouble Sleeping, **Happy (reverse scored)**, Lonely, Lacked Motivation.

```excel
=IF(COUNTA(C2:L2)<10,"",
    SWITCH(C2,"Rarely or none of the time",0,"Some or a little of the time",1,"Occasionally or a moderate amount of time",2,"All of the time",3,0)
   +SWITCH(D2,"Rarely or none of the time",0,"Some or a little of the time",1,"Occasionally or a moderate amount of time",2,"All of the time",3,0)
   +SWITCH(E2,"Rarely or none of the time",0,"Some or a little of the time",1,"Occasionally or a moderate amount of time",2,"All of the time",3,0)
   +SWITCH(F2,"Rarely or none of the time",0,"Some or a little of the time",1,"Occasionally or a moderate amount of time",2,"All of the time",3,0)
   +SWITCH(G2,"Rarely or none of the time",3,"Some or a little of the time",2,"Occasionally or a moderate amount of time",1,"All of the time",0,3)
   +SWITCH(H2,"Rarely or none of the time",0,"Some or a little of the time",1,"Occasionally or a moderate amount of time",2,"All of the time",3,0)
   +SWITCH(I2,"Rarely or none of the time",0,"Some or a little of the time",1,"Occasionally or a moderate amount of time",2,"All of the time",3,0)
   +SWITCH(J2,"Rarely or none of the time",3,"Some or a little of the time",2,"Occasionally or a moderate amount of time",1,"All of the time",0,3)
   +SWITCH(K2,"Rarely or none of the time",0,"Some or a little of the time",1,"Occasionally or a moderate amount of time",2,"All of the time",3,0)
   +SWITCH(L2,"Rarely or none of the time",0,"Some or a little of the time",1,"Occasionally or a moderate amount of time",2,"All of the time",3,0)
)
```
> **Note:** *Hopeful* and *Happy* are reverse-scored — more positive responses *lower* the total depression score rather than raise it.

---

## Dashboard Design Conventions

Header / banner asset lives in the original Drive at `CEGA Refugee Dashboard Project > Documentation > Design`.

**Canvas**
- Size: `600 × 1080` pixels (Desktop)
- Max **5 visualizations** per dashboard

**Typography**
- **Title** — Tableau Regular (bold), size 15, color `#1F86C8`
- **Subtitle** — Tableau Regular (italic), size 9, color default (black)
- **Statistics** (single-number callouts) — Tableau Regular (bold), size 12, color `#666666`

**Color**
- Background (dashboard): `#F1EFEF`
- Background (individual charts): `#FFFFFF`
- Chart palette: **Tableau 20**

**Layout**
- Round chart corners: `5`
- Outer padding: `10`
