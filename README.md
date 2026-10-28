# MeterVault — meter reading management

A mobile-first Estonian-language Flask application for managing shopping centres, categories, utility meters, monthly readings, users and Excel exports.

## Features in this starter implementation
- Username/password login with password hashing and admin/user roles.
- Admin-managed shopping centres, categories, meters, ordering, archive/reactivation and user accounts.
- Bulk entry for all active meters in a category, one category per screen.
- Monthly reading history; previous reading and consumption calculations.
- XLSX export filtered by month and optionally shopping centre.
- Basic audit log for reading saves, exports and admin changes.
- SQLite for local testing; PostgreSQL URL supported for deployment.
- Responsive UI, Estonian interface.

## Run locally

Python 3.11+ recommended.

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

Set bootstrap credentials before first run. On Windows PowerShell:

```powershell
$env:SECRET_KEY = "replace-with-a-long-random-secret"
$env:BOOTSTRAP_ADMIN_USER = "admin"
$env:BOOTSTRAP_ADMIN_PASSWORD = "use-a-unique-password-at-least-10-chars"
python app.py
```

Open `http://127.0.0.1:5000`. The initial admin is created only if both bootstrap variables are supplied and no admin exists. Add the three ordinary user accounts through **Haldus → Kasutajad**.

## Docker / Coolify

Build using the included Dockerfile. Set environment variables in Coolify:
- `SECRET_KEY`: long random secret.
- `BOOTSTRAP_ADMIN_USER`: desired first admin username.
- `BOOTSTRAP_ADMIN_PASSWORD`: unique password (10+ chars); after first successful bootstrap, remove this variable.
- `DATABASE_URL`: PostgreSQL connection URL for production, e.g. `postgresql://...` (use a persistent managed PostgreSQL service).
- `COOKIE_SECURE=1` when served over HTTPS.

Expose container port `5000`. **Configure a persistent PostgreSQL database before real use.** Do not rely on container-local SQLite for production because container rebuilds can lose data unless a persistent volume is configured.

## Production readiness
This is an initial working version, not a security-audited production release. CSRF protection is enabled for forms. Before using real company data, configure HTTPS and `COOKIE_SECURE=1`, use a persistent PostgreSQL database, configure and test backups/restores, rotate the secret key, and test the workflows with representative data. Validate the company's data-retention and access requirements. Shopping centres, categories and meters are archived/reactivated rather than physically deleted so historical readings remain intact.

## Data model
- `User`: account, password hash, role and active state.
- `Center`: shopping centre.
- `Category`: tenant/area category inside a centre.
- `Meter`: utility meter assigned to a category.
- `Reading`: unique record for each meter and `YYYY-MM` period.
- `AuditLog`: recorded actions.
