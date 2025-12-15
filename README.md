# Sheet-AppScript
Sheet automation for literally anything else than click

# Google Apps Script Automation Toolkit
 
## What This Repository Is

This designed to automate Google Sheets, Forms, and Drive workflows.

No fluff.
real operational problem that needed a clean, repeatable fix.
 

## What You Can Do With This

This project helps you:

* Automate data copying & dropdown propagation
* Lock formulas to prevent user mistakes
* Back up Sheets & Form responses automatically
* Convert Base64 images into Drive links
* Move and filter table data between sheets
* Generate location names from geolocation
* Track dates automatically
* Archive data on daily or monthly schedules

forms, reports, maintenance logs, or operational spreadsheets, this repo saves time immediately.

-----------<>-----------

## Script Overview

### Core Automation

* **AutoUpdateDateCell.gs**
  Automatically updates a date cells when changes occur.

* **DropdownListCopyDown.gs**
  Copies dropdown (data validation) rules downward dynamically as new rows are created.

* **LockFormulaValue.gs**
  Locks formulas into static values to prevent accidental edits.

---><---

### Data Movement & Duplication

* **CopyNameToNewSheet.gs**
  Creates new sheets dynamically based on selected names.

* **CopySheetToNewFile.gs**
  Duplicates sheets into entirely new Spreadsheet files.

* **Move(N)RowsToOtherSheet.gs**
  Moves a fixed number of rows to another sheet automatically.

* **MoveFilterValueTableToOtherTable.gs**
  Move filtered data table and transfers matching values.

---><---

### Backup & Archiving

* **BackupDataDaily.gs**
  Daily automated backup of spreadsheet data to new sheet.

* **BackupFormResponseDaily.gs**
  Daily backup of Google Form responses.

* **SaveTableDataMonthly.gs**
  Monthly data archiving for long‑term records.

---><---

### File & Media Utilities

* **ConvertBase64ToDriveLink.gs**
  Converts Base64 image data into Google Drive files and returns shareable links.

---><---

### Location & Metadata

* **AutomateGeoLocToLocationName.gs**
  Converts geolocation coordinates into readable location names automatically.

---><---

## Guide --> How to ?

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Copy the script you need into your project
4. Adjust sheet names / ranges
5. Add triggers if required

Each script is standalone ==> use only what you need.
 

