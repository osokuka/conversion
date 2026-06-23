# SharePoint → Horilla integration plan

## Goal

Keep the SharePoint staff list (*Lista e stafit*) as the HR source for job, contract, and payroll fields, and sync into Horilla automatically.

## Architecture

```
SharePoint List  →  Mapping layer  →  Horilla API
     │                    │                  │
  Graph API      sharepoint-horilla     employee-work-info-import
                 .config.json
```

## Phase 1 — Excel bridge (current)

**Status:** Implemented in this repo.

| Step | Tool |
|------|------|
| Export SharePoint list to Excel | Manual or scheduled |
| Transform | `Convert-SharePointToHorilla.ps1` |
| Import | Horilla UI → Employee → Actions → Import |

Keep this as a **fallback** when Graph or the API is unavailable.

### Column mapping

Defined in `sharepoint-horilla.config.json`:

| SharePoint column | Horilla field |
|-------------------|---------------|
| `Title` / `EEmri` | First Name / Last Name |
| `Nr personal` | Badge ID |
| `E-mail` | Email |
| `Gjinia` | Gender |
| `Projekti` | Department |
| `Pozita e punës` | Job Position |
| `Kualifikimi` | Job Role |
| `Lloji i punësimit` | Work Type / Employee Type |
| `Vendi i Punës` | Location |
| `Data e fillimit të punës` | Date Joining |
| `Data e përfundimit të kontratës` | Contract End Date |
| `Paga Bruto` | Basic Salary |
| `Statusi` | Filter (skip inactive) |

## Phase 2 — Microsoft Graph → Horilla API

**Duration:** ~4–5 weeks  
**Outcome:** Scheduled sync without manual Excel export.

### Azure app registration

| Permission | Purpose |
|------------|---------|
| `Sites.Read.All` | Read SharePoint list items |

Admin consent required.

### Graph request

```
GET /sites/{site-id}/lists/{list-id}/items?expand=fields
```

Store `site-id` and `list-id` in integration config (not in git if sensitive).

### Sync flow

1. Authenticate to Microsoft Graph (client credentials).
2. Read list items (use delta query in Phase 2b).
3. Apply the same rules as `Convert-SharePointToHorilla.ps1`:
   - Skip non-item rows
   - Skip inactive `Statusi` values
   - Map columns via `sharepoint-horilla.config.json`
   - Generate placeholder emails when `E-mail` is empty
4. Authenticate to Horilla (`POST /api/auth/login/`).
5. Submit import via `POST /api/employee/employee-work-info-import/`.
6. Write sync log; notify HR on errors.

### Implementation options

| Option | Notes |
|--------|-------|
| Extend PowerShell script (`-UseGraph`) | Reuses existing mapping; recommended |
| Python Azure Function | Good for cloud scheduling |
| Power Automate + HTTP | Possible; limited for complex mapping |

### Sync rules

| Rule | Action |
|------|--------|
| `defaults.Company` | Must match Horilla exactly (`Caritas Kosovo`) |
| Missing email | Placeholder using `PlaceholderEmailDomain` or block + alert |
| Inactive status | Skip (`Jo aktiv`, `Terminated`, `Pension`, etc.) |
| Duplicate badge/email | Skip row; log for HR review |

## Phase 2b — Delta sync

- Store last `lastModifiedDateTime` per SharePoint item.
- Process only changed rows.
- Reduces API load and duplicate import errors.

## Phase 3 — Optional write-back

Write to SharePoint list columns:

- Horilla employee ID
- Last sync timestamp
- Import status / error message

Only needed if HR works primarily in SharePoint and wants live status.

## Prerequisites

- [ ] Horilla API enabled and tested (see [API baseline](horilla-api-baseline.md))
- [ ] Azure app registration with `Sites.Read.All`
- [ ] SharePoint site ID and list ID documented
- [ ] Horilla master data aligned (company, departments, job positions)
- [ ] HR sign-off on placeholder email and inactive filter rules

## Timeline

| Phase | Duration | Outcome |
|-------|----------|---------|
| 1 — Excel | Done | Manual, reliable |
| 2 — Graph + API | 4–5 weeks | Automated sync |
| 2b — Delta | +2 weeks | Incremental updates |
| 3 — Write-back | Optional | Status in SharePoint |
