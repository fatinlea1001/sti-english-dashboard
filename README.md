# English SPM Dashboard — Live Google Sheets Connection

This repository hosts the HTML dashboard. The source data stays in the existing Google Sheet, and you only update your Sheet when entering new marks, changing student rosters, recording intervention activities, or adding Paper 1–4 item analysis.

## One-time Apps Script setup
1. Open [English Performance Dashboard — SPM Cohorts](https://docs.google.com/spreadsheets/d/1HHL92HTiXhSvXAE-1g3mRIVUdyM16agTJw1lPf7pEf0/edit).
2. Choose **Extensions → Apps Script**.
3. Replace **Code.gs** with the full contents of [AppsScriptBackend.gs](./AppsScriptBackend.gs) from this repository. This is a read-only JSON/JSONP service (not the full dashboard UI).
4. Save, then choose **Deploy → New deployment → Web app**. Select **Execute as: Me**, and **Who has access: Anyone**, then **Deploy** and authorise.
5. Copy the web app URL ending in **/exec**.
6. Open **https://fatinlea1001.github.io/sti-english-dashboard/?setup=1**.
7. Paste the **/exec** URL into the connection box and click **Connect & refresh**.
8. Click **Copy shareable dashboard link**. **Share that complete link** (including its `?api=` parameter) or embed it in Google Sites; the bare GitHub Pages URL does not carry the configured endpoint to new visitors.

## Future updates
- Change marks or student details in Google Sheets → reload the dashboard or press **Refresh from Sheet**.
- Add 2027/2028 cohorts using the existing column structure → reload.
- Change dashboard structure/design/features → edit `index.html` in this GitHub repository.
- Modify data feed or add backend fields → update Apps Script and redeploy a new version under the same web app deployment.

## Caution: student privacy
Your original Google Sheet **can remain Restricted**. However **anyone who has access to the public dashboard or its Apps Script feed can access student data returned by the endpoint** (names, classes, scores, progress). The `noindex` meta tag is a request to search engines, not an access control. Never store credentials, API keys, or student data exports in this public GitHub repository.

## Troubleshooting
If the dashboard shows a connection error, check that the Apps Script deployment **Who has access** is **Anyone** and that you are using the **/exec** URL (not `/dev`). After you edit Apps Script backend code, use **Deploy → Manage deployments → Edit → New version → Deploy**. On GitHub, use **Refresh from Sheet** after the deployment is updated.

## Data tabs
`RESULTS_DATA`, `HEADCOUNT_TRACKER`, `ITEM_ANALYSIS_DATA`, `TARGETS`, `INTERVENTION`, and `SETTINGS` are the read-only data sources. Blank or absent marks remain blank.
