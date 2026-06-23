# Active Directory / Entra ID → Horilla integration plan

## Goal

- **Login:** Staff sign into Horilla with corporate credentials.
- **Provisioning:** Identity fields stay aligned with Entra ID / AD.
- **Leavers:** Disabled AD accounts are deactivated in Horilla.

## What Horilla provides today

| Capability | Status |
|------------|--------|
| LDAP / AD login | Available via `horilla_ldap` module |
| REST API for employees | Available (used by mobile app) |
| Native AD employee sync | **Not built-in** — custom integration required |
| SAML / SCIM | On roadmap |

Integration = **LDAP for authentication** + **custom sync service via API**.

## Architecture

```
Entra ID / AD  →  Microsoft Graph (or LDAP)  →  Sync service  →  Horilla API
                                                      │
                                               Match by email / Badge ID
```

HR-specific fields (salary, contract, licenses) remain in **SharePoint** or Horilla — not in AD.

---

## Phase 1 — LDAP login

**Duration:** 2–3 weeks  
**Outcome:** Corporate login; employees must already exist in Horilla.

### Steps

1. Confirm `horilla_ldap` is available on your Horilla build (may require Enterprise).
2. Configure in Horilla → Settings → LDAP:
   - LDAP server (on-prem AD or Entra LDAPS if applicable)
   - `BASE_DN`, `BIND_DN`, service account password
   - User search filter (e.g. `(userPrincipalName=%(user)s)`)
3. Create read-only LDAP bind account in AD.
4. Pilot with 5–10 users.
5. Roll out org-wide.

### Limitation

LDAP login does **not** create Horilla employees. Accounts must exist before first login until Phase 2 is live.

---

## Phase 2 — Entra ID → Horilla employee sync

**Duration:** 4–6 weeks  
**Outcome:** Auto create, update, and archive employees.

### Azure app registration (Graph)

| Permission | Purpose |
|------------|---------|
| `User.Read.All` | Read user profiles |
| `Directory.Read.All` | Org structure, managers |

Admin consent required. Use **client credentials** (app-only) for the sync job.

### Field mapping

| Entra / AD | Horilla | Notes |
|------------|---------|-------|
| `employeeId` or extension attribute | Badge ID | Align with SharePoint `Nr personal` |
| `givenName` | First Name | |
| `surname` | Last Name | |
| `mail` | Email | Required |
| `mobilePhone` / `businessPhones` | Phone | |
| `department` | Department | |
| `jobTitle` | Job Position | |
| `officeLocation` | Location | |
| Manager reference | Reporting Manager | Resolve to Horilla employee name |
| `accountEnabled` | Active / archive | `false` → deactivate in Horilla |

**Do not sync from AD:** salary, contract dates, gender (unless reliably stored in AD), Albanian HR-specific fields.

### Sync service logic

```
Every 4–24 hours:
  1. GET users from Microsoft Graph
  2. POST /api/auth/login/ → Horilla JWT
  3. GET /api/employee/list/employees/
  4. Match by email or Badge ID
  5. CREATE new (accountEnabled=true, not in Horilla)
  6. UPDATE changed fields
  7. ARCHIVE when accountEnabled=false
  8. Log results; alert HR on errors
```

### Horilla API endpoints

| Action | Endpoint |
|--------|----------|
| List | `GET /api/employee/list/employees/` |
| Bulk create | `POST /api/employee/employee-work-info-import/` |
| Update work info | `PUT/PATCH /api/employee/employee-work-information/{id}/` |
| Deactivate | `POST /api/employee/employee-archive/{id}/false/` |

### Conflict rules

| Situation | Action |
|-----------|--------|
| Source of truth — identity | Entra ID |
| Source of truth — HR/payroll | SharePoint or Horilla |
| Duplicate email | Skip + alert HR |
| New AD user, active | Create in Horilla |
| AD account disabled | Archive in Horilla |
| In Horilla, not in AD | Flag for HR (no auto-delete) |

### Security

- Store secrets in Azure Key Vault or Windows Credential Manager.
- Dedicated Horilla integration account (not a person).
- HTTPS only; audit log all changes.

---

## Phase 3 — SSO polish (future)

When Horilla SAML support is generally available:

- Entra ID SAML for browser login (replaces LDAP bind).
- SCIM provisioning (if Horilla adds it) may replace custom sync.

Until then: **LDAP login + Graph sync** is the practical enterprise pattern.

---

## Combined roadmap with SharePoint

| Track | Priority | Reason |
|-------|----------|--------|
| SharePoint → Horilla API | **First** | Extends existing conversion script; highest ROI |
| AD LDAP login | Second | Security and UX |
| Entra ID sync | Third | Identity automation; complements SharePoint HR data |

SharePoint carries fields AD does not have (salary, contracts, licenses). AD carries identity and leaver status.

---

## Prerequisites

- [ ] Horilla URL and API tested ([API baseline](horilla-api-baseline.md))
- [ ] `horilla_ldap` confirmed on your instance
- [ ] Azure app registration for Graph
- [ ] Badge ID stored in AD (`employeeId` or extension attribute) — same as SharePoint `Nr personal`
- [ ] HR sign-off on conflict and deactivation rules
- [ ] Staging environment for sync testing

## Timeline

| Phase | Duration | Outcome |
|-------|----------|---------|
| 1 — LDAP login | 2–3 weeks | AD credentials in Horilla |
| 2 — Graph sync | 4–6 weeks | Auto provision / deactivate |
| 3 — SAML | TBD | Modern SSO when available |
