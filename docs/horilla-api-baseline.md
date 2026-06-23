# Horilla API baseline

Shared reference for Active Directory and SharePoint integrations.

## Authentication

| Item | Value |
|------|-------|
| Login endpoint | `POST https://YOUR_HORILLA/api/auth/login/` |
| Header | `Authorization: Bearer <access_token>` |
| Token type | JWT (`rest_framework_simplejwt`) |

## Documentation

| Resource | URL |
|----------|-----|
| Swagger UI | `https://YOUR_HORILLA/api/swagger/` |
| ReDoc | `https://YOUR_HORILLA/api/redoc/` |
| Horilla docs | https://docs.horilla.com/technical/v2.0/doc/api/baseapi.html |

Official PDFs are also attached to [GitHub issue #345](https://github.com/horilla/horilla-hr/issues/345) (Base module + Employee API).

## Key employee endpoints

| Method | Endpoint | Purpose |
|--------|----------|---------|
| `GET` | `/api/employee/list/employees/` | List employees |
| `GET` | `/api/employee/employees/{id}/` | Employee detail |
| `GET/POST` | `/api/employee/employee-work-information/` | Work info records |
| `POST` | `/api/employee/employee-work-info-import/` | Bulk import (same data as Excel template) |
| `GET` | `/api/employee/employee-work-info-export/` | Export work info |
| `POST` | `/api/employee/employee-bulk-update/` | Bulk updates |
| `POST` | `/api/employee/employee-archive/{id}/{is_active}/` | Archive / activate |

## Base master data

Under `/api/base/` — departments, job positions, job roles, work types, shifts, etc. Create or align these before bulk employee import.

## Integration service account

Create a dedicated Horilla user for automation (not a personal account):

1. Log in via API and store credentials securely (Key Vault or Credential Manager).
2. Use a service account with permission to import employees.
3. Log every sync run: created, updated, skipped, failed.

## Security notes

- Use HTTPS only in production.
- Rotate API credentials on a schedule.
- Test on staging before production import.
- Review your Horilla version for API permission defaults; keep the instance patched.

## Canonical employee key

Pick one matching field and use it consistently across SharePoint and AD sync:

| Option | SharePoint field | AD / Entra field |
|--------|------------------|------------------|
| Badge ID (recommended) | `Nr personal` | `employeeId` or extension attribute |
| Email | `E-mail` | `mail` |

Document the choice in `sharepoint-horilla.config.json` and AD attribute mapping.
