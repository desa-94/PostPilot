# PostPilot

Every day the mailroom gets a pile of scanned letters that need to go to the
right department. PostPilot does that for me: it pulls the scans out of
Outlook, works out where one letter ends and the next begins from the incoming
stamp, reads who the sender is, and forwards each one to the right address.

The whole thing runs on fixed rules – no AI model, no cloud, nothing leaves the
machine. That's on purpose. Letter contents are sensitive, and I always want to
be able to tell *why* a letter ended up where it did. Every decision traces
back to one concrete rule.

When a letter isn't a clear match, it doesn't get sent blindly – it goes into a
review folder instead. Better to glance at it myself than mail something to the
wrong tax office.

## What it does

- Splits batch PDFs (many letters scanned back-to-back) at the incoming stamp,
  purely by colour detection, no OCR.
- Pulls sender and date from the scan's text layer, tolerating OCR noise.
- Scores each letter to a department, and if needed drills down further by the
  address line (e.g. Garnishment vs. Tax Department).
- Prepares the result as an Outlook draft, or sends it straight away.

## Getting started

```bash
pip install -r requirements.txt
```

Runs on Windows with Outlook (for sending). Tesseract isn't required – it's
only there as an optional fallback for the rare scan without a text layer.

See every command:

```bash
py -m postpilot
```

The usual run – grab the scans from Outlook and prepare drafts:

```bash
py -m postpilot --fetch
```

Once the drafts look right, send for real with `--send`:

```bash
py -m postpilot --fetch --send
```

## Configuration

Which department gets which rules and recipient address lives in
`config/departments.yaml`. The real addresses and my signature sit in a local
copy (`departments.local.yaml`) that stays out of the repo – the versioned file
only holds placeholders.
