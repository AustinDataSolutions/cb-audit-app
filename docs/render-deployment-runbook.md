# Automatic Audit: Render deployment runbook

This guide is written for the app's owner. It moves the Automatic Audit from
Streamlit Community Cloud to Render, puts it behind individual accounts, and
copies the team's existing accounts over so nobody needs a new password. Each
numbered step says what to click, what to type, and what you should see. Do
the sections in order.

What you end up with: one small web service on Render named
`aarp-automatic-audit`, in Render's Virginia region, with a 1 GB disk that
holds only the accounts and usage database, deploying only when you click
Deploy. The address will be `https://aarp-automatic-audit.onrender.com/`
unless Render reports the name is taken, in which case it adds a suffix; note
the address Render shows you.

## 1. Before you start: gather the values

Have these ready in a password manager or a private note. You will type them
into Render's form. Never paste them into chat, email, or a ticket.

| Setting | What it is | Where to find it |
|---|---|---|
| `ANTHROPIC_API_KEY` | The Anthropic key the app uses today | The Community Cloud app's Settings, Secrets tab, or your key storage |
| `OPENAI_API_KEY` | The OpenAI key the app uses today | Same place |
| `AUDIENCE` | The short phrase the prompts use for who wrote the feedback, for example "members" | The Community Cloud secrets (the `AUDIENCE` line); not a secret, just wording |
| `SMTP_HOST`, `SMTP_USER`, `SMTP_PASSWORD`, `EMAIL_FROM` | The email-sending settings behind "Email results to me" | The Community Cloud secrets. If email was never set up, leave them blank and the feature stays off. |

The shared app password is not needed any more and is not entered anywhere.

## 2. Create the service from the Blueprint

The repository contains a file named `render.yaml` that describes the
service. Render reads it and builds the form for you.

1. Open the Render dashboard and make sure you are in the same workspace as
   the Model Rules Explorer and the Themes Explorer.
2. Click **New** (top right), then **Blueprint**.
3. Find the GitHub repository **AustinDataSolutions/cb-audit-app** and click
   **Connect**. If Render says it cannot see the repository, click the link to
   configure GitHub access, tick that repository, and come back.
4. Render shows a page titled with the repository name. Leave **Blueprint
   Name** as `cb-audit-app` or change it to `aarp-automatic-audit`; either is
   fine. For **Branch**, pick `ads-auth-accounts`. The Blueprint file path stays
   `render.yaml`.
5. Render lists one web service, `aarp-automatic-audit`, with its settings
   read from the file. Check these before continuing:
   - Region: Virginia (US East)
   - Instance type: 512 MB, 0.5 CPU ("Starter")
   - A disk named `automatic-audit-data`, 1 GB, at `/var/data`
6. Below the service, Render asks for the values it cannot read from the file.
   Type each one exactly as you gathered it in section 1. The box names are
   `ANTHROPIC_API_KEY`, `OPENAI_API_KEY`, `AUDIENCE`, `SMTP_HOST`, `SMTP_USER`,
   `SMTP_PASSWORD`, and `EMAIL_FROM`. For a value you do not have (for example
   the email settings), leave the box empty.
7. Click **Apply** (sometimes labelled **Deploy Blueprint**). Render creates
   the service and starts the first build.

If Render created the service without asking for the values, do not worry:
open the service, click **Environment** in the left menu, and fill in each
name listed above there, then click **Save Changes**. The service restarts.

## 3. Watch the first build

Click the new service, then **Logs**. A first build takes two to four
minutes. You should see, in order:

1. Lines installing Python packages, ending with one that mentions
   `ads-auth` and GitHub (the shared sign-in package).
2. `Build complete.`
3. A line beginning `preflight: ok` showing `auth=required`. This is the
   app checking its own settings before it agrees to serve.
4. Streamlit starting, then Render reporting the service is **Live**.

If the build fails, copy only the last few non-secret lines of the log. Never
copy a line containing a key, a password, or a token.

## 4. Switch off automatic deploys and check the sign-in screen

1. Still in the service, click **Settings**. Scroll to **Build & Deploy** and
   confirm **Auto-Deploy** is **Off**. The Blueprint set this, but check.
2. Go back to the dashboard, open **Blueprints** in the left menu, open this
   Blueprint, and set **Auto Sync** to **No**. This means changes to the
   repository never deploy by themselves; you always click **Manual Deploy**.
3. Open the service's address in a private browser window. You should see a
   sign-in screen titled "AARP Automatic Audit" with username and password
   boxes, a link "About this app and data handling", and nothing else. It will
   say the app has no accounts yet, which is expected at this point.
4. Click the About link. The data-handling page should open without signing
   in.

## 5. Copy the team's accounts

The Model Rules Explorer and the Themes Explorer hold the same accounts, with
the same passwords. You copy one app's account database onto this service's
disk, then tell this app to import the accounts from it. Passwords are stored
as scrambled hashes, so the copy carries no readable passwords, and people
keep the password they already have.

This takes two Render Shells: one on the source app, one on the new app. A
Shell is a command line running inside the service.

**On the source app** (the Model Rules Explorer, `model-rules-explorer`):

1. Open that service, click **Shell** in the left menu, and wait for the
   prompt to appear.
2. Type the following and press Enter. It writes a consistent copy of the
   accounts database to a temporary file:

   ```bash
   python manage.py backup-db
   ```

   It prints the path of the file it wrote, under `/var/data/backups/`.
3. The copy must now get to the new service. Render Shells cannot send files
   to each other directly, so use a private, temporary channel that both
   shells can reach. The simplest is a short-lived private upload: the person
   guiding you will give you one command to run in this shell that uploads the
   backup file and prints a one-time download address, and one command to run
   in the other shell that downloads it. The address works once and expires
   within minutes. Do not post the address anywhere.

**On the new app** (`aarp-automatic-audit`):

4. Open the new service, click **Shell**, and download the copy with the
   command you were given. It lands at `/tmp/users-copy.db`.
5. Do a dry run. This lists what would be imported and changes nothing. One
   `--allow` per username; the person guiding you will give you the exact
   line with the team's usernames filled in:

   ```bash
   python manage.py import-users /tmp/users-copy.db --allow <user1> --allow <user2> ...
   ```

   Read the list it prints. It should name each person on the team under
   "Will import" and nothing under "Not found" or "Unsupported".
6. If the list is right, run the same line again with `--apply` added at the
   end. The accounts are created with their existing passwords.
7. Give the manager and administrator roles back, one line per person:

   ```bash
   python manage.py set-role <your-username> admin
   python manage.py set-role <manager-username> manager
   ```

8. Confirm with `python manage.py list-users`, then delete the copy here and
   the backup on the source shell:

   ```bash
   rm /tmp/users-copy.db
   ```

9. In a private browser window, open the new app's address and sign in with
   your usual username and password. Upload a small test file, run a short
   audit, download the workbook, open **Usage statistics** from the sidebar
   (it should show your sign-in, upload, run, and download), then **Log out**.

## 6. Cut over from Community Cloud

Until this section, the Community Cloud app still works with the shared
password and the Render app runs from the `ads-auth-accounts` branch.

1. Tell the team the date and time. Any audit running on Community Cloud at
   that moment will stop.
2. The person guiding you merges the branch into `main` and pushes. That one
   push does two things: Community Cloud redeploys itself and now shows only a
   notice with the new address, and the repository's `main` branch becomes
   the one Render deploys from.
3. In Render, open the service, **Settings**, **Build & Deploy**, change
   **Branch** to `main`, and save. Then click **Manual Deploy**, **Deploy
   latest commit**. Watch the Logs for `preflight: ok` and **Live**. Accounts
   are on the disk and survive the deploy.
4. Open the old Community Cloud address. It should show "Automatic Audit has
   moved" with a link to the new address. Nothing is uploaded or recorded
   there any more.
5. After 30 days, delete the Community Cloud app: open share.streamlit.io,
   find the app, open its menu (the three dots), choose **Delete**, and type
   the confirmation it asks for. Then remove the shared password from the
   password manager and tell former users it no longer works.

## 7. Day to day

- **Deploying a change:** open the service, **Manual Deploy**, **Deploy
  latest commit**. Deploy when no audit is running; a deploy ends it.
- **Changing a key or the email settings:** **Environment**, edit the value,
  **Save Changes**. The service restarts.
- **Adding, resetting, or removing an account:** **Shell**, then the commands
  in `docs/OPERATIONS.md`.
- **Something is wrong:** **Logs** shows the last lines; **Events** shows
  deploys and restarts. Roll back by deploying the previous commit.
