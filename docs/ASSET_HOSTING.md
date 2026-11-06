# Portable ARISE asset hosting

GitHub raw URL pattern:
`https://raw.githubusercontent.com/madprohub/ARISE-EXCEL-ASSETS/main/assets/ARISE_Neon_Cursor.ico`

The URL works only after the actual binary asset has been uploaded. Download to a local folder or embed the cursor in the workbook before assigning MSForms `MouseIcon`. An `.ico` file may not be supported by every MSForms `LoadPicture` implementation; test compatibility and use a fallback cursor.

For offline reliability, embed the cursor or distribute `ARISE_NEON_ASSETS` alongside the workbook.
