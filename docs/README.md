# Horilla integration documentation

Plans for connecting Horilla HRMS with Microsoft 365 services used by Caritas Kosovo.

| Document | Description |
|----------|-------------|
| [Horilla API baseline](horilla-api-baseline.md) | Auth, endpoints, and shared prerequisites |
| [SharePoint integration](sharepoint-integration.md) | SharePoint staff list → Horilla (Excel + Graph API) |
| [Active Directory integration](active-directory-integration.md) | Entra ID / AD login and employee sync |

## Current tooling

The `Convert-SharePointToHorilla.ps1` script in the repo root is **Phase 1** of the SharePoint plan: Excel export → Horilla import template.

See the [main README](../README.md) for how to run it.

## Recommended order

1. SharePoint Graph → Horilla API (extends existing conversion script)
2. Active Directory LDAP login
3. Entra ID → Horilla identity sync
