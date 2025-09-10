# Metadata & Branding Sanitization

Follow these steps to ensure no original author identifiers remain.

## PDF
- Use ExifTool to clear and reset fields:
  - macOS (Homebrew): `brew install exiftool`
  - Strip and set: `exiftool -overwrite_original -all= -Title="Real Estate Investment Dashboard" -Author="Data Analyst" -Creator="Data Analyst" -Producer="Data Analyst" "path/to/report.pdf"`

## DOCX
- Open the file in Microsoft Word.
- File → Info → Check for Issues → Inspect Document.
- Remove: Document Properties and Personal Information, Custom XML Data, Comments, Revisions.
- Set File → Info → Properties → Advanced Properties → Summary: Title "Real Estate Investment Dashboard", Author "Data Analyst", Company "Acme Analytics".
- Save a copy.

## XLSX/XLSM
- Open in Excel.
- File → Info → Check for Issues → Inspect Workbook.
- Remove: Document Properties and Personal Information, Hidden Names, Custom XML Data.
- File → Info → Properties → Advanced Properties → Summary: set Title/Author/Company.
- Save a copy.

## Git & README
- Ensure links and text no longer reference external usernames or repositories.
- Replace titles with "Real Estate Investment Dashboard" where relevant.


