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
$env:SECRET_KEY = "paste-a-random-secret-of-at-least-32-characters"
$env:BOOTSTRAP_ADMIN_USER = "admin"
$env:BOOTSTRAP_ADMIN_PASSWORD = "use-a-unique-password-at-least-14-chars"
python app.py
```

Open `http://127.0.0.1:5000`. The initial admin is created only if both bootstrap variables are supplied and no admin exists. Add the three ordinary user accounts through **Haldus → Kasutajad**.

## Docker / Coolify

Build using the included Dockerfile. Set environment variables in Coolify:
- `SECRET_KEY`: long random secret.
- `BOOTSTRAP_ADMIN_USER`: desired first admin username.
- `BOOTSTRAP_ADMIN_PASSWORD`: unique password (10+ chars); after first successful bootstrap, remove this variable.
- `DATABASE_URL`: PostgreSQL connection URL for production, e.g. `postgresql://...` (use a persistent managed PostgreSQL service).
- `APP_ENV=production`
- `COOKIE_SECURE=1` when served over HTTPS (required in production).
- `SECRET_KEY` must be a unique random value of at least 32 characters; the app refuses to start without it.
- Use HTTPS at the reverse proxy, and configure that proxy to be the only trusted proxy in front of the app.

Expose container port `5000`. **Configure a persistent PostgreSQL database before real use.** Do not rely on container-local SQLite for production because container rebuilds can lose data unless a persistent volume is configured.

## Security hardening included
- Startup fails if `SECRET_KEY` is missing or weak; production requires `COOKIE_SECURE=1`.
- CSRF protection on state-changing requests, secure session-cookie settings, an eight-hour session lifetime, and security response headers.
- Database-backed login throttling shared across application workers.
- Admin-only controls remain enforced server-side; the dashboard now also requires an active login.
- XLSX text cells are protected against spreadsheet formula injection.
- Docker runs the application as a non-root user.
- Passwords set through the admin interface must be at least 14 characters.

## Production deployment checklist
This code hardening does not replace an infrastructure penetration test. Before using real company data: use HTTPS and PostgreSQL; keep secrets only in the hosting platform's secret store; remove bootstrap credentials after the first admin is created; restrict database/network access; configure automated encrypted backups and test restoration; monitor logs and failed logins; keep the host OS, reverse proxy and dependencies patched; and verify access permissions with separate ordinary/admin accounts. Never publish `.env`, database files, backups or source-control secrets.

## Data model
- `User`: account, password hash, role and active state.
- `Center`: shopping centre.
- `Category`: tenant/area category inside a centre.
- `Meter`: utility meter assigned to a category.
- `Reading`: unique record for each meter and `YYYY-MM` period.
- `AuditLog`: recorded actions.
