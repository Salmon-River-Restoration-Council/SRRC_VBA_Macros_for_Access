# VBA Macros for SRRC Water/Air Temp Access Database: A Quick Guide

This README provides a basic description and tutorial for a set of VBA (Visual Basic for Applications) macros designed for use with SRRC's `srrc_wtemp_db.accdb` Microsoft Access databases. These macros facilitate various operations such as exporting data, updating records, and logging changes. These macros should already be saved into the current version of the wtemp database.

## 1. ExportYearlySummary `./YearlySummary.bas`

### Description
This macro exports yearly summary data from specified tables within an Access database to Excel. It offers options to include a "Sites" sheet and to export data for all years or a specific year.

### Tutorial
1. **Prepare Excel:** Ensure the Microsoft Excel Object Library is imported via Tools > References in the VBA editor.
2. **Run the Macro:** Execute `ExportYearlySummary`.
3. **User Prompts:** Respond to prompts regarding including the "Sites" sheet and the range of years to export.
4. **Export Path:** The macro saves the Excel files in a predefined folder within the project's path.

## 2. UpdateSiteYears `./YearCorrection.bas`

### Description
Updates a "Years" column for each site in the `tbl_Sites` table based on the distinct years of data available across specified data tables.

### Tutorial 
1. **Run the Macro:** Execute `UpdateSiteYears`.
2. **Automatic Updates:** The macro automatically updates the "Years" column in `tbl_Sites` with a comma-separated list of years for which data exists.
3. **Completion Message:** A message box confirms the successful update.

## 3. ExportSitesToDate `./SitesToDate.bas`

### Description
Exports data for specified sites and years from the database to Excel. It allows filtering by site code and year, and offers an option to pare down columns in the exported data.

### Tutorial
1. **User Prompts:** Upon execution, respond to prompts for site code(s), year(s), and whether to pare certain columns.
2. **Export Path:** Exports are saved in a designated folder `ExportsFromAccess\SitesToDate` within the DB's path.
3. **Excel Creation:** For each site and year combination, an Excel file is created with the relevant data.

## 4. ReplaceSiteCodeAndLogChanges `./SiteCode_Backfill.bas`

### Description
Searches for and replaces instances of a specified word in the `SiteCode` column across multiple tables. It logs each replacement in a CSV file, including the date in the filename.

### Tutorial
1. **Run the Macro:** Execute `ReplaceSiteCodeAndLogChanges`.
2. **User Prompts:** Input the word to find and the replacement word when prompted.
3. **CSV Log:** The macro generates a CSV file in the project's path, logging all replacements made, with the filename including the current date.

### General Notes
- **Error Handling:** Each macro includes basic error handling to manage common issues, such as invalid input or database access errors.
- **Customization:** Macros can be customized to fit specific database schemas or requirements by adjusting table names, column names, and SQL queries as needed.
- **Backup:** Always back up your database before running these macros to prevent accidental data loss.
