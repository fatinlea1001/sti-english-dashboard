/**
 * ONE-TIME INSTALL:
 * In the EXISTING Google Sheet: Extensions > Apps Script > + > Script
 * Name it CohortSync. Paste this file's code, Save, and reload the Sheet.
 * You do NOT need to replace Code.gs or redeploy your web app.
 *
 * Future workflow:
 *   1. Paste new cohort/class/name records into NEW_COHORT_INTAKE.
 *   2. English Dashboard menu > Sync New Cohort Intake.
 *   3. Enter scores in RESULTS_DATA; refresh the GitHub dashboard.
 *
 * This function is only callable from your bound Google Sheets editor.
 * It is not part of the public doGet API.
 */
const COHORT_SYNC_SPREADSHEET_ID = '1HHL92HTiXhSvXAE-1g3mRIVUdyM16agTJw1lPf7pEf0';

/**
 * NEW COHORT INTAKE — run from the private Google Sheets UI.
 *
 * Once installed: refresh the spreadsheet, then choose
 *   English Dashboard > Sync New Cohort Intake
 * Paste data in NEW_COHORT_INTAKE, columns A–F:
 *   Cohort year | Academic year | Form | Class | Student name | Status.
 * The script creates missing master/enrolment/exam/headcount rows.
 * It NEVER overwrites existing marks, targets, or item analysis.
 *
 * The doGet web endpoint above remains strictly read-only.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('English Dashboard')
    .addItem('Sync New Cohort Intake', 'syncNewCohortIntake')
    .addToUi();
}

function cohortKey_(year, name) {
  return String(year).trim() + '|' +
    String(name).trim().replace(/\s+/g, ' ').toUpperCase();
}

function intakeClass_(value) {
  return String(value || '').replace(/^\s*[45]\s+/, '')
    .trim().replace(/\s+/g, ' ').toUpperCase();
}

function ensureIntakeCapacity_(sheet, lastRequired) {
  if (lastRequired > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), lastRequired - sheet.getMaxRows());
  }
}

function pasteNewFormulas_(sheet, startRow, count, width) {
  if (!count) return;
  const source = sheet.getRange(startRow - 1, 1, 1, width);
  const target = sheet.getRange(startRow, 1, count, width);
  source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
}

function adjustResultFormula_(formula, templateRow, newRow) {
  if (!formula) throw new Error('The existing RESULT_DATA template is missing a calculated formula.');
  const pat = new RegExp('\\b([A-Z]{1,3})' + templateRow + '\\b', 'g');
  return formula.replace(pat, (_whole, col) => col + newRow);
}

function buildResultTemplate_(key, exam) {
  const year = Number(key.split('|')[0]);
  const first = exam === 'TOV' || exam === 'AR1';
  return [year, first ? year - 1 : year, first ? 'Form 4' : 'Form 5',
    '', '', exam, '', ''];
}

function classFromNewIntake_(intake, exam) {
  const first = exam === 'TOV' || exam === 'AR1';
  return (intake.form === (first ? 'Form 4' : 'Form 5') &&
    intake.academicYear === (first ? intake.cohort - 1 : intake.cohort))
    ? intake.clazz : '';
}

function resultMarkRef_(row) {
  if (!row) return ['', ''];
  return [
    '=IF(ISNUMBER(RESULTS_DATA!H' + row + '),RESULTS_DATA!H' + row + ',"")',
    '=IF(RESULTS_DATA!I' + row + '="","",RESULTS_DATA!I' + row + ')'
  ];
}

function syncNewCohortIntake() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    ui.alert('Another cohort sync is already running. Please retry later.');
    return;
  }
  try {
    const ss = SpreadsheetApp.openById(COHORT_SYNC_SPREADSHEET_ID);
    const intakeSheet = ss.getSheetByName('NEW_COHORT_INTAKE');
    if (!intakeSheet) throw new Error('Missing NEW_COHORT_INTAKE tab.');
    const lastInput = intakeSheet.getLastRow();
    if (lastInput < 2) {
      ui.alert('NEW_COHORT_INTAKE is empty. Paste rows under its headers first.');
      return;
    }
    const inputRows = intakeSheet.getRange(2, 1, lastInput - 1, 8).getDisplayValues();
    const masters = ss.getSheetByName('STUDENT_MASTER');
    const enrolments = ss.getSheetByName('ENROLMENT');
    const results = ss.getSheetByName('RESULTS_DATA');
    const tracker = ss.getSheetByName('HEADCOUNT_TRACKER');
    const summary = ss.getSheetByName('HEADCOUNT_SUMMARY');
    const required = [masters, enrolments, results, tracker, summary];
    if (required.some(sh => !sh)) throw new Error('One or more required dashboard tabs are missing.');

    // Read existing records by unique cohort + student name; do not assume
    // that physically adjacent rows belong to the same student.
    const masterValues = masters.getRange(2, 1, Math.max(1, masters.getLastRow() - 1), 2).getDisplayValues();
    const enrolValues = enrolments.getRange(2, 1, Math.max(1, enrolments.getLastRow() - 1), 4).getDisplayValues();
    const resultValues = results.getRange(2, 1, Math.max(1, results.getLastRow() - 1), 10).getDisplayValues();
    const trackValues = tracker.getRange(4, 1, Math.max(1, tracker.getLastRow() - 3), 3).getDisplayValues();
    const summaryValues = summary.getRange(2, 1, Math.max(1, summary.getLastRow() - 1), 2).getDisplayValues();
    const masterMap = new Map();
    masterValues.forEach((r, i) => {
      if (r[0] && r[1]) masterMap.set(cohortKey_(r[1], r[0]), i + 2);
    });
    const enrolMap = new Map();
    enrolValues.forEach((r, i) => {
      if (r[3] && r[0]) enrolMap.set(cohortKey_(r[0] + '/' + r[1], r[3]), {row: i + 2, clazz: r[2]});
    });
    const resultMap = new Map();
    resultValues.forEach((r, i) => {
      if (r[0] && r[4] && r[5]) resultMap.set(cohortKey_(r[0], r[4]) + '|' + r[5], i + 2);
    });
    const trackMap = new Map();
    trackValues.forEach((r, i) => { if (r[0] && r[2]) trackMap.set(cohortKey_(r[0], r[2]), i + 4); });
    const summaryMap = new Set(summaryValues.filter(r => r[0] && r[1]).map(r => cohortKey_(r[0], r[1])));

    const newMaster = [], newEnrol = [], newResults = [], newTracker = [], newSummary = [];
    const classPatches = [];
    const statuses = inputRows.map(r => [r[6] || '']);
    const exercises = ['TOV', 'AR1', 'AR2', 'AR3', 'SPM'];
    const trackerSourceRow = tracker.getLastRow();
    const targetGradeOriginals = tracker.getRange(trackerSourceRow, 1, 1, 19).getFormulas()[0];
    const nextMasterRow = masters.getLastRow() + 1;
    const nextEnrolRow = enrolments.getLastRow() + 1;
    const nextResultRow = results.getLastRow() + 1;
    const nextTrackerRow = tracker.getLastRow() + 1;
    const nextSummaryRow = summary.getLastRow() + 1;
    let newlyRegistered = 0, placementsAdded = 0, reviewCount = 0;

    // Plan all data before writing anything. No marks/targets are copied from
    // examples; new assessments have genuinely blank marks.
    inputRows.forEach((row, i) => {
      const filled = row.slice(0, 5).some(v => String(v).trim() !== '');
      if (!filled) return;
      const cohort = Number(row[0]), academicYear = Number(row[1]);
      const form = String(row[2] || '').trim(), clazz = intakeClass_(row[3]);
      const name = String(row[4] || '').trim().replace(/\s+/g, ' ');
      const status = String(row[5] || 'Active').trim();
      if (!Number.isInteger(cohort) || cohort < 2024 || cohort > 2100 ||
        !Number.isInteger(academicYear) || academicYear < cohort - 1 ||
        academicYear > cohort || !['Form 4', 'Form 5'].includes(form) ||
        !clazz || !name || !['Active','Transferred','Withdrawn','Completed'].includes(status) ||
        (form === 'Form 4' && academicYear !== cohort - 1) ||
        (form === 'Form 5' && academicYear !== cohort)) {
        statuses[i] = ['Review: cohort/form/year/class/name']; reviewCount++;
        return;
      }
      const key = cohortKey_(cohort, name);
      const entry = {cohort: cohort, academicYear: academicYear, form: form,
        clazz: clazz, name: name, status: status};
      let masterRow = masterMap.get(key);
      let didImport = false;
      if (!masterRow) {
        masterRow = nextMasterRow + newMaster.length;
        masterMap.set(key, masterRow);
        newMaster.push([name, cohort, status, '']);
        didImport = true; newlyRegistered++;
      }
      const enrolKey = cohortKey_(academicYear + '/' + form, name);
      let enrolRecord = enrolMap.get(enrolKey);
      if (enrolRecord && enrolRecord.clazz && enrolRecord.clazz !== clazz) {
        statuses[i] = ['Review: different class already registered'];
        reviewCount++;
        return;
      }
      if (!enrolRecord) {
        enrolRecord = {row: nextEnrolRow + newEnrol.length, clazz: clazz};
        enrolMap.set(enrolKey, enrolRecord);
        newEnrol.push([academicYear, form, clazz, name, status, '']);
        placementsAdded++;
      }
      // Ensure all 5 exam slots exist, even if the student had a partial
      // import in a previously interrupted sync.
      exercises.forEach(exam => {
        const resultKey = key + '|' + exam;
        if (resultMap.has(resultKey)) {
          const existingIndex = resultMap.get(resultKey);
          const past = resultValues[existingIndex - 2];
          const first = exam === 'TOV' || exam === 'AR1';
          const expectedForm = first ? 'Form 4' : 'Form 5';
          const expectedYear = first ? cohort - 1 : cohort;
          if (past && !past[3] && form === expectedForm && academicYear === expectedYear) {
            classPatches.push({row: existingIndex, clazz: clazz});
          }
          return;
        }
        const newRow = nextResultRow + newResults.length;
        resultMap.set(resultKey, newRow);
        const vals = buildResultTemplate_(key, exam);
        vals[3] = classFromNewIntake_(entry, exam);
        vals[4] = name;
        newResults.push(vals);
      });
      if (!trackMap.has(key)) {
        const tr = nextTrackerRow + newTracker.length;
        trackMap.set(key, tr);
        const mr = masterRow, er = enrolRecord.row;
        const ar = exercises.map(ex => resultMarkRef_(resultMap.get(key + '|' + ex)));
        const gradeFormula = (sourceIndex, markColumn) =>
          targetGradeOriginals[sourceIndex].split(markColumn + trackerSourceRow)
            .join(markColumn + tr);
        newTracker.push([
          '=IF(STUDENT_MASTER!B' + mr + '="","",STUDENT_MASTER!B' + mr + ')',
          '=IF(ENROLMENT!C' + er + '="","",ENROLMENT!C' + er + ')',
          '=IF(STUDENT_MASTER!A' + mr + '="","",STUDENT_MASTER!A' + mr + ')',
          ar[0][0], ar[0][1], '', gradeFormula(6,'F'),
          ar[1][0], ar[1][1], '', gradeFormula(10,'J'),
          ar[2][0], ar[2][1], '', gradeFormula(14,'N'),
          ar[3][0], ar[3][1], '', gradeFormula(18,'R')
        ]);
      }
      if (!summaryMap.has(key)) {
        summaryMap.add(key);
        const sr = nextSummaryRow + newSummary.length;
        const tr = trackMap.get(key);
        const mark = ex => resultMarkRef_(resultMap.get(key + '|' + ex))[0];
        newSummary.push([
          cohort, name, mark('TOV'),mark('AR1'),mark('AR2'),mark('AR3'),
          '=IF(ISNUMBER(HEADCOUNT_TRACKER!R' + tr + '),HEADCOUNT_TRACKER!R' + tr + ',"")',
          '=IF(OR(C' + sr + '="",G' + sr + '=""),"",G' + sr + '-C' + sr + ')'
        ]);
      }
      statuses[i] = [didImport ? 'Imported' : 'Enrolment synced'];
    });

    if (newResults.length) {
      const templateRow = results.getLastRow();
      const template = results.getRange(templateRow, 1, 1, 14).getFormulas()[0];
      const sourceColumns = [8, 10, 11, 12];
      if (sourceColumns.some(idx => !template[idx])) {
        throw new Error('RESULTS_DATA is missing grade/attendance formulas in its final template row.');
      }
      const start = templateRow + 1;
      ensureIntakeCapacity_(results, start + newResults.length - 1);
      pasteNewFormulas_(results, start, newResults.length, 14);
      results.getRange(start, 1, newResults.length, 8).setValues(newResults);
      results.getRange(start, 9, newResults.length, 1)
        .setFormulas(newResults.map((row, idx) =>
          [adjustResultFormula_(template[8], templateRow, start + idx)]));
      results.getRange(start, 11, newResults.length, 3)
        .setFormulas(newResults.map((row, idx) =>
          [10, 11, 12].map(col => adjustResultFormula_(template[col], templateRow, start + idx))));
    }
    if (newMaster.length) {
      const start = masters.getLastRow() + 1;
      ensureIntakeCapacity_(masters, start + newMaster.length - 1);
      masters.getRange(start, 1, newMaster.length, 4).setValues(newMaster);
      masters.getRange(start, 3, newMaster.length, 1)
        .setDataValidation(masters.getRange(2, 3).getDataValidation());
    }
    if (newEnrol.length) {
      const start = enrolments.getLastRow() + 1;
      ensureIntakeCapacity_(enrolments, start + newEnrol.length - 1);
      enrolments.getRange(start, 1, newEnrol.length, 6).setValues(newEnrol);
      enrolments.getRange(start, 5, newEnrol.length, 1)
        .setDataValidation(enrolments.getRange(2, 5).getDataValidation());
    }
    if (newTracker.length) {
      const start = tracker.getLastRow() + 1;
      ensureIntakeCapacity_(tracker, start + newTracker.length - 1);
      const template = tracker.getRange(start - 1, 1, 1, 19);
      const dest = tracker.getRange(start, 1, newTracker.length, 19);
      template.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      template.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      dest.setValues(newTracker);
    }
    if (newSummary.length) {
      const start = summary.getLastRow() + 1;
      ensureIntakeCapacity_(summary, start + newSummary.length - 1);
      const template = summary.getRange(start - 1, 1, 1, 8);
      const dest = summary.getRange(start, 1, newSummary.length, 8);
      template.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dest.setValues(newSummary);
    }
    classPatches.forEach(change => {
      const cell = results.getRange(change.row, 4);
      if (!cell.getDisplayValue()) cell.setValue(change.clazz);
    });
    intakeSheet.getRange(2, 7, statuses.length, 1).setValues(statuses);
    SpreadsheetApp.flush();
    ui.alert('Cohort sync complete!\n\n' +
      newlyRegistered + ' new students registered.\n' +
      placementsAdded + ' enrolment records added.\n' +
      newResults.length + ' missing exam slots prepared.\n' +
      reviewCount + ' records need review.\n\n' +
      'Enter marks in RESULTS_DATA; check COHORT_OVERVIEW and refresh the dashboard.');
  } catch (error) {
    ui.alert('Cohort sync stopped: ' + String(error && error.message || error) +
      '\nPlease review the workbook before running it again. Existing marks were not intentionally overwritten.');
    throw error;
  } finally {
    lock.releaseLock();
  }
}
