# Preparing Brevo domain authentication for austindatasolutions.com

Instructions for whoever manages the DNS records of austindatasolutions.com as
part of the planned transition from SendGrid to Brevo. The Automatic Audit app
has not yet been switched to Brevo in production. Brevo needs three or four
DNS records on the domain before mail from it is trusted. Nothing about the
existing mail setup changes: these records do not affect where incoming mail
goes, and the SendGrid records can stay or go as you prefer.

The values below in angle brackets come from Brevo's domain setup screen
(Settings, then Senders, Domains, IPs, then Domains, then Authenticate next to
austindatasolutions.com, then "Authenticate the domain yourself"). The app's
owner copies them from that screen into this note before sending it. Brevo
shows the exact Host and Value for each record; copy them exactly, and do not
add quotation marks that Brevo does not show.

## Records to add

| Purpose | Type | Host / Name | Value |
|---|---|---|---|
| Brevo ownership check ("Brevo code") | TXT | `<host Brevo shows, usually @ or austindatasolutions.com>` | `<brevo-code=... value Brevo shows>` |
| DKIM signature (so receivers can tell the mail is genuine) | TXT, or two CNAME records | `<host Brevo shows, for example mail._domainkey or brevo1._domainkey and brevo2._domainkey>` | `<value Brevo shows>` |
| DMARC policy (tells receivers what to do with mail that fails the checks) | TXT | `_dmarc` | `<value Brevo shows, for example v=DMARC1; p=none; rua=mailto:...>` |

Notes for the DNS editor:

- Some DNS providers want the host written relative to the domain (`_dmarc`)
  and some want it in full (`_dmarc.austindatasolutions.com`). Use whichever
  form the provider's other records already use.
- If the domain already has a `_dmarc` TXT record, do not add a second one.
  Leave the existing record in place and say so; Brevo accepts an existing
  DMARC record as long as one exists.
- The DKIM record is a long string. Keep it on one line. If the provider
  splits long TXT values into chunks, that is fine.
- Nothing needs to be removed. An SPF record (a TXT record starting with
  `v=spf1`) is not required by Brevo for this setup.
- TTL can stay at the provider's default.

## How to tell it worked

DNS changes can take anywhere from a few minutes to 48 hours to spread. After
the records are saved, the app's owner goes back to the same Brevo screen and
clicks "Authenticate this email domain". Brevo then shows the domain as
**Authenticated** with a green check. Until then it shows **Not
authenticated** and a message that the domain is not authenticated yet, which
just means to try again later.

If it still shows as not authenticated after 48 hours, the usual causes are a
typo in a host name, a value saved with extra quotation marks, or the record
saved under the wrong domain. Comparing each record in the DNS editor against
the Brevo screen character by character finds it.

When the domain is authenticated, reply to the app's owner so they can finish
the remaining steps: confirm the sending address in Brevo, set the SMTP values
in Render, and send a test from the app.
