# How this app handles your data

Last updated 19 October 2026. This describes the hosted AARP Automatic Audit
on Render at aarp-automatic-audit.onrender.com. Each person now signs in with
their own account, the same username and password used for the Model Rules
Explorer and the Themes Explorer; the shared password used until October 2026
no longer works. The older copy of the tool on Streamlit Community Cloud shows
only a notice pointing here; it has no accounts and records nothing, and it
will be removed 30 days after the move. This page is reachable from the
sign-in screen and from the app's sidebar.

## What the app is for

The app checks how accurately a CX Designer topic model labels sentences. You
upload an audit workbook exported from CX Designer, optionally a model tree
file to pick topics and supply their descriptions, and the app asks an
external language model, topic by topic, whether each sentence was labelled
correctly. It produces a completed audit workbook with an accuracy figure for
every topic and, for topics below the accuracy threshold you set, a short
summary of what went wrong. You can download the workbook or have it emailed
to you.

The sentences in an audit workbook are excerpts of member feedback. Upload
de-identified exports only. Do not include member names, email addresses,
street addresses, or other direct identifiers.

## What is stored, where, and for how long

Signing in, signing out, uploading an audit workbook, uploading a model tree,
starting an audit run, having results emailed, and downloading a workbook are
recorded with your account name and the time, so the team can see how much the
tool is used. The files, their contents, the sentences, the prompts, the
model's answers, and email addresses are never recorded.

| Data | Where it goes | How long it is kept | Who can see it |
|---|---|---|---|
| The audit workbook you upload, and a model tree if you add one | Held in the app's memory on the server for your session only. They are never written to the server disk. | Until you replace them, sign out, your session ends (8 hours after sign-in), or the server restarts or is redeployed | Only you, during that session |
| The audit in progress and its results | Kept in the server's memory while the run continues, including a checkpoint of the work so far, so that a run keeps going if your browser disconnects. Only one audit runs at a time; anyone else who opens the app while it runs sees a progress notice and nothing else. | Until the run ends and you have downloaded or been emailed the result, you start another run, or the server restarts (a restart ends the run; the last emailed or downloaded checkpoint is what survives) | You, in the browser tab that started the run. Other signed-in people see only that a run is in progress and how far along it is. |
| Workbooks you download | Built in memory and sent to your browser. The app keeps no copy. | Not kept by the app | Whoever you share the download with. Downloads contain the original sentences and the model's judgments, so handle them like the source file. |
| Workbooks emailed to you | Sent as an attachment from the team's sending address, through the email service the administrator set up, to the address you type and confirm. The app checks that the address is an AARP or Austin Data Solutions address and asks you to confirm any other address. The address is held in the server's memory for that run only. | The app keeps nothing. Your mailbox and the email service keep the message under their own retention. | You, and whoever can read that mailbox |
| Sentences sent to a model provider | Sent over an encrypted connection to Anthropic's or OpenAI's API, whichever provider you choose in the sidebar, during an audit run and when summaries are generated. The next section says exactly what each request carries and what each provider keeps. | The app keeps the requests only in your session. | The chosen provider, as the company processing the request |
| Your account | Username, display name, role, and a salted hash of your password, in a database on the server disk. The password itself is never stored. | Until the administrator deactivates or removes the account | The administrator |
| Failed sign-in attempts for an existing account | Account name and time, used only to lock the account for 15 minutes after five failures within 15 minutes | 15 minutes | Nobody views these; they are deleted automatically |
| Usage events | One row per sign-in, failed sign-in, sign-out, uploaded audit workbook, uploaded model tree, audit run started (including a resumed checkpoint and a summary-only rerun), results email sent, and workbook download: time, account name, and event type. Failed sign-ins for names that match no account are recorded with no name. File names, file contents, sentences, prompts, model answers, and email addresses are never recorded. | 13 months, then deleted automatically | Managers and the administrator, as weekly and monthly counts, per-user totals, and CSV downloads |
| Backup copies of the account and usage database | Written once a day to a backups folder on the server disk. Uploaded files and results are never on the disk, so backups hold accounts and usage events only. | The newest seven copies, about a week | The administrator |
| Server log | Progress lines for each run: which topic number is being audited and its topic name, how long a model request took, whether a request was retried, and whether the results email was sent. Also start-up and housekeeping notices. Never sentences, prompts, model answers, file names, email addresses, usernames, or passwords. | For the hosting provider's own log retention period | The operator, in the hosting dashboard |

### What is sent to Anthropic or OpenAI, and when

Only one provider is used for a run: the one selected under **Provider** in
the sidebar. The app connects to that provider with a key the administrator
set on the server; nobody types a key into the app.

- **The provider's model list.** Each time the page loads, the app asks the
  chosen provider for the list of models you may pick from. That request
  carries the key and nothing else.
- **Auditing a topic.** During a run the app sends one request per topic. It
  carries the topic's name, its description from the model tree if you added
  one, up to the number of sentences you allowed under **Max sentences per
  category** (each with its row number), the model's title and description,
  anything you typed under **About this model**, and the organization and
  audience names the administrator configured. The written instruction
  shown under **Audit prompt** goes with it; you can read and edit that text
  before you run.
- **Summarizing a low-accuracy topic.** After the audit, for each topic whose
  accuracy is below the threshold you set, the app sends one request carrying
  the model's own explanations of the sentences it judged mislabelled, the
  topic name, and your notes about the model, with the instruction shown
  under **Summary settings**. Sentences themselves are not sent again; the
  explanations may quote them.

**Anthropic.** The requests go to Anthropic's commercial API under its
standard terms: Anthropic states that it does not use API requests to train
its models and that it deletes request and response data from its systems
within 30 days. Anthropic may keep content longer if its automated safety
systems flag it, for up to two years, or where the law requires it. No
account-specific retention agreement is in place.

**OpenAI.** The requests go to OpenAI's commercial API under its standard
terms: OpenAI states that data sent to its API is not used to train or
improve its models unless the customer opts in, which this app does not. OpenAI
keeps abuse-monitoring logs, which may include the prompts and responses, for
up to 30 days, or longer where the law requires it or to protect its services
or others from harm. The app does not ask OpenAI to store conversations
beyond that. No zero-data-retention agreement is in place.

Clearing data in the app does not remove anything a provider still holds.

### Limits and other details

Uploads are limited to 50 MB per file. The sidebar limits on topics,
sentences per topic, and tokens per request cap what one run can send. The
app runs one audit at a time for everyone; a second person who starts one is
asked to wait. Your session ends after 8 hours, when you sign out, or when an
administrator resets your password.

The app uses no third-party services other than the hosting service, Render,
the model provider you choose, and the email service used to send results. It
uses no analytics (Streamlit's usage statistics are switched off) and no
cookies beyond the ones Streamlit needs to keep your browser connected to your
session.

The server runs on one server in Render's Virginia (US) region. The accounts
and usage database lives on a small disk attached to that server. Render keeps
daily snapshots of that disk for its own retention period. Those snapshots are
a separate layer from the app's own backups and may hold copies of the database
for longer than the app itself keeps them. Uploaded files, results, and emails
never touch that disk.

## What will change soon

Nothing further is planned at the moment.

## Who to contact

J.T. Fales, Austin Data Solutions, jt.fales@austindatasolutions.com
