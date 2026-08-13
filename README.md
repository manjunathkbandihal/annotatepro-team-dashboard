# AnnotatePro Team Dashboard Pro

Professional React + Vite dashboard for annotation team operations.

## Excel / Google Sheets import
Use **Sheet Import** to upload the `.xlsx` exported from your Google Sheet. The importer reads the date blocks and:
- Name
- Project
- Annotation/Review
- Total images worked
- Link to the range

Imported records are stored in browser localStorage.

## Run
npm install
npm run dev

## Build
npm run build

## Daily dashboard
The Dashboard now derives daily production/review metrics from imported sheet records. It supports date selection and status rules: Completed, Processing, Pending, and On Leave. Production/Update target is 800; Review target is 400.


### Import behavior
Excel dates are preserved using local calendar dates (no timezone shift). The dashboard defaults to the latest date found in the imported sheet.


### Latest fixes
- Date filters normalize stored/imported dates so previous days and date ranges are included correctly.
- Projects: every project card has an Edit button; status and remaining are recalculated automatically.
