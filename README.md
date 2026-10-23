# English SPM Dashboard — Multi-Cohort System

This GitHub Pages dashboard reads current marks and class data from the existing Google Sheet.

- [Open the live Google Sheet](https://docs.google.com/spreadsheets/d/1HHL92HTiXhSvXAE-1g3mRIVUdyM16agTJw1lPf7pEf0/edit)
- [Open the dashboard](https://fatinlea1001.github.io/sti-english-dashboard/)
- **Existing cohorts:** SPM 2026 (Form 5 during 2026), SPM 2027 (Form 4 during 2026)
- **Next Form 4 in 2027:** SPM cohort **2028**

## One-time setup for adding future cohorts

If the GitHub dashboard currently says "Apps Script connected", its existing Code.gs/web app connection **does not need to be replaced or redeployed** for this upgrade.

1. Open the [existing Google Sheet](https://docs.google.com/spreadsheets/d/1HHL92HTiXhSvXAE-1g3mRIVUdyM16agTJw1lPf7pEf0/edit).
2. Click **Extensions → Apps Script**.
3. In the left panel beside **Files**, click **+ → Script**.
4. Name the file **CohortSync** (Google adds the `.gs` extension).
5. Open [CohortSync.gs](./CohortSync.gs) in this GitHub repository. Copy the **entire** source code, paste it into the new Apps Script `CohortSync.gs` file, and **Save**.
6. Return to Google Sheets and **reload** it. A new **English Dashboard** menu should appear.
7. This is a **one-time setup**. Keep your existing `Code.gs` untouched.

Only `Code.gs` handles the read-only public data feed. `CohortSync.gs` adds a **private spreadsheet-menu action** for importing new students. That action is **not** an endpoint callable by public dashboard visitors.

## Add a new batch in 2027

Open the **NEW_COHORT_INTAKE** tab and enter the following columns in order:

| SPM_Cohort | Academic_Year | Form | Class | Student_Name | Enrolment_Status |
|---|---|---|---|---|---|
| 2028 | 2027 | Form 4 | *student's actual class* | *student's actual name* | Active |

**The row above is an example only. Do not treat it as a student record.**

- Use one row per student. Leave **Sync_Status** blank; the importer fills it. The Notes column is optional.
- You can paste an entire class or several classes at once.
- Choose **English Dashboard → Sync New Cohort Intake**.
- The script creates a row in **STUDENT_MASTER**, the year/form/class row in **ENROLMENT**, five blank exam slots (TOV, AR1, AR2, AR3, SPM) in **RESULTS_DATA**, a new **HEADCOUNT_TRACKER** row, and a **HEADCOUNT_SUMMARY** row. Existing names, scores, targets and item-analysis records are not intentionally overwritten.
- **After the one-time roster sync**, enter actual marks in `RESULTS_DATA` column H; set Absent in column J when appropriate. Grades, pass flags, and headcount formulas update automatically.
- Open **COHORT_OVERVIEW** for a per-cohort rollup. Refresh the GitHub dashboard and choose **SPM 2028 → Form 4**.
- When the **same** SPM 2027 students become Form 5 in academic year 2027, paste their **same names and cohort (2027)** with Form 5 / actual class into **NEW_COHORT_INTAKE**, then sync again. This updates enrolment and previously blank exam class assignments without adding duplicate students.

Do **not** sort or move rows in `RESULTS_DATA`, `HEADCOUNT_TRACKER`, or `HEADCOUNT_SUMMARY`; older headcount formulas use fixed exam-row references. Use filter views instead.

## What updates automatically?

- Google Sheets marks / attendance / targets → calculated fields update in Sheets.
- New students synced from the intake tab → master, enrolment, results, tracker and headcount rows are prepared.
- COHORT_OVERVIEW → dynamic cohort counts and marks coverage without fixed 2027-only ranges.
- GitHub dashboard → reads the new cohort/form values from the existing live Sheets feed when refreshed.
- Dashboard layout/features → change `index.html` on GitHub only if the *design* needs changing.

**Important:** The original `OVERVIEW` tab was built specifically for SPM 2027. Use the new `COHORT_OVERVIEW` tab when you need multi-cohort results.

## Existing dashboard Apps Script setup (only if not connected yet)

The original [AppsScriptBackend.gs](./AppsScriptBackend.gs) is the read-only backend. If the dashboard isn't yet connected, install it as Code.gs, deploy it as a web app with **Execute as Me**, and connect its `/exec` URL at [Dashboard connection setup](https://fatinlea1001.github.io/sti-english-dashboard/?setup=1). If it already works, **do not repeat this step**.

## Confidentiality

Keep the Google Sheet restricted to authorised staff. The current public dashboard and its deployed read-only feed can still expose returned students' names and marks to people who obtain the URL. `noindex` requests avoidance of search listing, but it is **not** an access-control measure. Never upload exported student data or Google credentials to this public GitHub repository.
