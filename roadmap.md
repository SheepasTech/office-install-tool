# Roadmap

Rough goals, not a task list. Nothing here is built yet unless the README says so.

Done so far: guided menu joining diagnose, cleanup and install; plain-words error hints after a failed setup; per-job log in `Cases`. Next: presets, change/undo tracking, live log following, verified product list.

## Shape of the project
- One project: the installer and the diagnose/cleanup kit now live together here. Next is joining them into a single guided flow.
- Started with a one-liner pasted into a PowerShell terminal, the same way as the Windows Transfer Tool (`irm <raw url>/bootstrap.ps1 | iex`). It downloads the tool to a local folder, asks for admin rights itself, and starts.
- Options by environment variable on the same line, e.g. a dry-run switch or a pinned version.

## Install flow
- Ask what to install (edition, language, architecture, excluded apps) and generate the install script and XML from the answers.
- Remember finished choices as named presets, so a repeat job is "pick the preset" instead of answering everything again.
- Product key typed into the terminal at run time. It is never written to a file, preset, log or the ticket.
- Download and extract the Office Deployment Tool from Microsoft, and check that it is signed by Microsoft.

## Readiness check and changes
- Dry run first: test whether the PC is ready (disk space, file system, services, antivirus or filter drivers, leftover Office, network, pending reboot) and report without changing anything.
- Then let me tick which fixes to apply, or apply them all after a clear warning.
- Keep a note of every system change made for the install, so each one can be put back afterwards.
- An undo step at the end that restores what the preparation changed.

## Logs and errors
- Follow the install while it runs: the setup logs and the Click-to-Run processes.
- Catch Office installer errors as they happen and explain them in plain words, with the likely fix (the error table in `README.md` is the starting point).
- Save a short log per job into the ticket folder, with any key redacted.

## Supported Office versions
- Work out and keep a verified list of every Office version the installer can handle: product IDs, update channels and which need a key.
- Only the official Microsoft lists count as proof. Third-party pages are leads, not facts.
- Early findings, still to verify against Microsoft's "Product IDs supported by the Office Deployment Tool for Click-to-Run" page:
  - Microsoft 365 Apps and the 2019, 2021 and 2024 volume editions (Professional Plus, Standard, Project, Visio) use `PerpetualVL2019`, `PerpetualVL2021` and `PerpetualVL2024` channels.
  - The 2024 retail IDs (Home & Student, Home & Business, Professional) are not confirmed on a Microsoft page yet. Wrong IDs would fail, so they need a check.
  - Language codes come from Microsoft's languages list for Microsoft 365 Apps.

## Later
- Offline mode: download Office once, install on several PCs.
- Catalog of editions and languages kept as a data file, so adding one is a one-line change.
