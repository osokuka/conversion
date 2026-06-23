# SharePoint to Horilla Conversion

Converts a SharePoint staff list Excel export into a Horilla employee import file.

**Latest update:** 2026-06-23 — auto-install ImportExcel, auto-detect files, gender maps to `Male`/`Female`

## Files

| File | Purpose |
|------|---------|
| `Convert-SharePointToHorilla.ps1` | Main conversion script |
| `sharepoint-horilla.config.json` | Column mappings and defaults |
| `sharepoint_list.xlsx` | SharePoint export (place your file here) |
| `work_info_template.xlsx` | Horilla import template |
| `horilla_import.xlsx` | Generated output (created after running the script) |

## Setup (first time)

```powershell
git clone git@github.com:osokuka/conversion.git
cd conversion
```

Or if you already have the folder:

```powershell
cd C:\Users\BleonaAdemi\Desktop\HRRR\conversion
git pull
```

## Run

1. Export your SharePoint list to Excel
2. Save/rename it as `sharepoint_list.xlsx` in this folder
3. Run:

```powershell
cd C:\Users\BleonaAdemi\Desktop\HRRR\conversion
.\Convert-SharePointToHorilla.ps1
```

4. Import `horilla_import.xlsx` into Horilla (Employee → Actions → Import)

The script will:
- Install `ImportExcel` automatically on first run
- Find `sharepoint_list.xlsx` and `work_info_template.xlsx` automatically
- Write `horilla_import.xlsx` in the same folder

## If you don't see updates

Your local folder may be outdated. Run:

```powershell
cd C:\Users\BleonaAdemi\Desktop\HRRR\conversion
git pull
```

Check you are on the latest commit:

```powershell
git log -1 --oneline
```

Expected latest: `Fix gender JSON duplicates and map values to Male/Female`

## Config

Edit `sharepoint-horilla.config.json` to change:
- `defaults.Company` — must match the company name in Horilla exactly
- `genderMap` — SharePoint `Mashkull` → `Male`, `Femër` → `Female`
- `columnMap` — if SharePoint column names change
