# SharePoint to Horilla Conversion

Converts a SharePoint staff list Excel export into a Horilla employee import file.

**Integration plans:** See [docs/](docs/README.md) for Active Directory and SharePoint API integration roadmaps.

## Files

| File | Purpose |
|------|---------|
| `Convert-SharePointToHorilla.ps1` | Main conversion script |
| `sharepoint-horilla.config.json` | Column mappings and defaults |
| `work_info_template.xlsx` | Horilla import template (included; or download fresh from Horilla) |
| `sharepoint_list.xlsx` | SharePoint export (place your file here) |
| `horilla_import.xlsx` | Generated output (created after running the script) |

## Setup (first time)

```powershell
git clone git@github.com:osokuka/conversion.git
cd conversion
```

## Run

1. Export your SharePoint list to Excel
2. Save/rename it as `sharepoint_list.xlsx` in this folder
3. Run:

```powershell
cd C:\apps\sharepoint_to_horilla\conversion
.\Convert-SharePointToHorilla.ps1
```

4. Import `horilla_import.xlsx` into Horilla (Employee → Actions → Import)

The script will:
- Download `ImportExcel` into a local `Modules` folder on first run (avoids broken OneDrive module paths)
- Find `sharepoint_list.xlsx` automatically
- Use `work_info_template.xlsx`, or create one with the correct Horilla headers if missing
- Write `horilla_import.xlsx` in the same folder

### Optional parameters

```powershell
.\Convert-SharePointToHorilla.ps1 `
  -SharePointExportPath .\sharepoint_list.xlsx `
  -TemplatePath .\work_info_template.xlsx `
  -OutputPath .\horilla_import.xlsx `
  -ConfigPath .\sharepoint-horilla.config.json `
  -WorksheetName 'query (26)'
```

Use `-WorksheetName` when the SharePoint export has multiple sheets and the data is not on the first sheet.

## Config

Edit `sharepoint-horilla.config.json` to change:
- `defaults.Company` — must match the company name in Horilla exactly
- `genderMap` — SharePoint `Mashkull` → `Male`, `Femër` → `Female`
- `columnMap` — if SharePoint column names change
- `filters.SkipInactive` — skip rows where `Statusi` is inactive

## Horilla import tips

- Department, job position, and company names must match what exists in Horilla (or Horilla will create them)
- Dates are written as `yyyy-MM-dd`
- Employees without email get a placeholder address using `defaults.PlaceholderEmailDomain`
