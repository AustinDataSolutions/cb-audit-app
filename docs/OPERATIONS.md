# Operations runbook: Automatic Audit

All commands run from the repository root locally, or from the Render Shell
for the hosted app (`cd /opt/render/project/src` is the default there).
Nothing here prints a password hash. Temporary passwords are printed once.
Accounts, sessions, usage events, backups, and the commands below come from
the shared `ads-auth` package; `manage.py` runs them with this app's settings.

## Accounts

| Task | Command |
|---|---|
| First administrator | `python manage.py create-user <username> --role admin --full-name "<Name>"` (prompts twice for the password) |
| New user with a temporary password | `python manage.py create-user <username> --full-name "<Name>" --temporary` then give the printed password to the person privately; it must be changed at first sign-in |
| Manager (sees the Usage statistics page) | `python manage.py create-user <username> --role manager --temporary`, or `python manage.py set-role <username> manager` for an existing account |
| Reset a password | `python manage.py set-password <username>` (same temporary-password behaviour; also signs the person out everywhere and clears any lockout) |
| Locked out | Wait 15 minutes, or `python manage.py unlock-user <username>` |
| Promote or demote | `python manage.py set-role <username> user` (roles: user, manager, admin) |
| Deactivate / reactivate | `python manage.py deactivate-user <username>`, `python manage.py reactivate-user <username>` |
| Delete permanently | `python manage.py delete-user <username> --force` (usage history keeps the name; cannot be undone) |
| List accounts | `python manage.py list-users` |

The last active administrator cannot be deactivated, demoted, or deleted.
Five failed sign-ins within 15 minutes lock an account for 15 minutes. A
session ends 8 hours after sign-in, at sign-out, or when the password is reset.

The Usage statistics page leaves out the accounts in `USAGE_EXCLUDE_USERS`
(default `ads-dev`) unless a manager changes the "Exclude accounts" box.

## Copying accounts from the Model Rules Explorer or Themes Explorer

Both sibling apps use the same package, so their database holds accounts in
the same shape and people keep their passwords. The copy needs `ads-auth`
v1.1.1 or later on this app.

1. In the source app's Render Shell, take a consistent copy of its database:
   `sqlite3 /var/data/app.db ".backup /tmp/users-copy.db"` (or
   `python manage.py backup-db`, which writes to `/var/data/backups/`).
2. Move the copy to this service's Shell through an authenticated channel
   (never email, tickets, or Git). The deployment runbook, section 5, uses a
   private upload-and-download step for this.
3. Dry run, one `--allow` per username to import:
   `python manage.py import-users /tmp/users-copy.db --allow <user1> --allow <user2>`
4. Review the report, then add `--apply`. Imported accounts get role `user`
   and keep their existing password.
5. Set manager and admin roles with `set-role`, then delete the copies on
   both shells: `rm /tmp/users-copy.db`.

## Usage events

The app records `login`, `login_failed`, `logout`, `upload_audit_file`,
`upload_model_tree`, `run_audit`, `email_results`, and `download`: time,
account, and event type only. Rows older than 13 months are removed by the
daily housekeeping pass. Managers download CSVs from the Usage statistics
page; `python manage.py export-events --out <file>` writes the same four-column
export (`ts_utc,app,username,event`, with `app=automatic-audit`) from the
shell, and `python manage.py prune-events` runs the retention pass by hand.

## Backups

The app writes a consistent copy of `app.db` to `DATA_DIR/backups/` once a
day and keeps the newest seven. Uploaded workbooks and results are never on
the disk, so a backup holds accounts and usage events only. Nothing leaves
the Render disk.

| Task | Command |
|---|---|
| Backup now | `python manage.py backup-db` |
| Restore | Suspend the service in the Render dashboard, then `python manage.py restore-db /var/data/backups/<file> --force`, then Resume. `--force` is needed because the lock file from the last start is still present. |
| Everything lost, no backup | Re-import from a sibling app (above) or recreate accounts with `create-user`. Usage history cannot be recovered. |

## Deploying

Deploys are manual: pushing the branch does not deploy. In the Render
dashboard open the service, **Manual Deploy**, **Deploy latest commit** (or a
specific commit). The build installs the pinned packages, including `ads-auth`
from its GitHub tag; the start script runs the package's preflight checks
(settings valid, sign-in required, `DATA_DIR` writable, database migrated) and
refuses to serve if one fails, which the Logs tab shows as a `preflight:` line.

A deploy restarts the process, which ends any audit in progress. Deploy when
nobody is running one (the app shows a progress banner to everyone while a run
is active). Roll back by deploying the previous commit the same way. Accounts
and usage events live on the disk and survive deploys and restarts.

## Settings on the service

| Variable | Purpose |
|---|---|
| `ANTHROPIC_API_KEY`, `OPENAI_API_KEY` | Provider keys. Rotate by changing the value and redeploying. |
| `ORGANIZATION`, `AUDIENCE` | Named in the prompts and the page title. |
| `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`, `EMAIL_FROM` | Email delivery. Blank `SMTP_HOST` switches the feature off. The planned Brevo migration uses `smtp-relay.brevo.com`, port 587, the SMTP login and an SMTP key from Brevo's SMTP & API page, and a sender address verified in Brevo. See "Pending Brevo email migration" below. |
| `DATA_DIR`, `AUTH_MODE` | `/var/data` and `required`; the Blueprint sets them. |
| `USAGE_EXCLUDE_USERS` | Optional; accounts hidden from the statistics page by default. |

Changing a value restarts the service.

## Pending Brevo email migration

The Audit app already supports sending a finished workbook through an SMTP
service. The app has not yet been switched to Brevo in production. Once the
Brevo account, sending domain, and SMTP settings are ready, "Email results to
me" will send the finished workbook through Brevo's SMTP relay. The free plan
allows 300 emails a day and adds a "Sent with Brevo" footer; the Audit app
sends one message per emailed run, so the daily limit is not a practical
concern.

- **Sending domain:** authenticate `austindatasolutions.com` in Brevo
  (Settings, Senders, Domains, IPs, Domains). The DNS records needed for that
  step are described in `brevo-dns-instructions.md` in this folder.
- **Sender:** when the domain is ready, create or verify the sender address in
  Brevo. `EMAIL_FROM` must match that sender; Brevo refuses mail from any other
  address.
- **SMTP key:** generate an SMTP key on Brevo's SMTP & API page, name it after
  this app, and set it as `SMTP_PASSWORD` on the service. Brevo can deactivate
  a key that has not sent anything for 90 days; reactivate the key there if
  email later stops working after a quiet period.
- **Checking the connection:** after setting the Brevo values in Render, run
  this from the Render Shell without printing the key:

  ```bash
  python - <<'EOF'
  import os, smtplib, ssl
  host, port = os.environ.get("SMTP_HOST", ""), int(os.environ.get("SMTP_PORT", "587"))
  print("host:", host, "port:", port)
  try:
      with smtplib.SMTP(host, port, timeout=30) as s:
          print("connected")
          s.starttls(context=ssl.create_default_context()); print("encryption on")
          s.login(os.environ["SMTP_USER"], os.environ["SMTP_PASSWORD"]); print("login accepted")
  except Exception as exc:
      print("failed at the step after the last line above:", type(exc).__name__, str(exc)[:200])
  EOF
  ```

  "login accepted" means the settings are right; a failure names the step
  that broke (connection, encryption, or sign-in). A sign-in failure usually
  means a wrong or deactivated key, or `SMTP_USER` set to the account's
  sign-in email instead of the SMTP login shown on the SMTP & API page.
- **Delivery log:** once email delivery is live, Brevo's Transactional, Logs
  page shows each message and whether it was delivered, bounced, or refused.

## Local runs with accounts

`.env` with `AUTH_MODE=required` and `DATA_DIR=.data` gives a local database
under the git-ignored `.data/` folder. Create an account with
`python manage.py create-user <username>` and start the app with
`python -m streamlit run app.py`. With `AUTH_MODE=disabled` the sign-in screen
is skipped and nothing is recorded; Render refuses that mode.
