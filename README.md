# Remind

A small Windows app for emailing a reminder to every employee on a list.

1. **File.** Drop in the employee list (CSV or Excel). Duplicate numbers are removed. The employee-number column is
   detected automatically, and you can pick another column if the guess is wrong.
2. **Recipients.** Each number is looked up in Spectrum (Viewpoint) with your own Spectrum SQL login to get the
   name and email. People with no email, or with an email that doesn't look right, are flagged. You can edit an
   email, add or remove people, and choose who gets the message.
3. **Message.** Write the email, or start from a saved **template**. Mail-merge fields such as `{FirstName}`,
   `{LastName}` and `{EmployeeNumber}` are filled in for each person. Preview it, send yourself a test, then send
   everyone their own email through classic Outlook. You can also save them all to Drafts instead.

Templates can be saved, renamed and deleted, and one is marked as the **default** that Remind opens with.
Everything you send is kept in History.

## Install

Download `Remind-win64.zip` from the latest [release](../../releases/latest), extract it, and double-click
`Install.cmd`. No admin rights are needed.

Remind needs:

- classic (desktop) Outlook, signed in
- Microsoft ODBC Driver 17 or 18 for SQL Server

The first time you look up emails, Remind asks for the Spectrum **server**, **database** and, optionally, the
**email view**, plus your Spectrum user name and password. Ask whoever set Remind up for the server details.
They are saved only on your PC (`%APPDATA%\Remind`), and the password is encrypted with Windows DPAPI.

## Updates

Installed copies check this repository's latest GitHub Release when they start, and from
**Settings > Check for updates**. When a newer version is out, Remind offers to download it (checking GitHub's
SHA-256 for the file), replace itself, and restart. Settings, templates, the saved login and history are kept.

To publish a new version:

```powershell
.\release.ps1 1.2.0 "What changed"
```

This bumps `src/version.py`, tags `v1.2.0` and pushes. GitHub Actions (`.github/workflows/release.yml`) then builds
`Remind-win64.zip` and attaches it to the release.

## Develop

```powershell
py -3.14 -m pip install -r requirements.txt
py -3.14 src\main.py          # run from source
.\build.ps1                   # local build -> dist\Remind-win64.zip
```

Set `REMIND_SPECTRUM_SERVER`, `REMIND_SPECTRUM_DATABASE` and `REMIND_SPECTRUM_EMAIL_VIEW` to test against a server
without saving it in settings.

**Never commit** server names, logins, employee lists or anything from `%APPDATA%\Remind`. This repository and
its releases are public. `.gitignore` blocks the usual data files.

| File | What it does |
| --- | --- |
| `src/main.py` | the window: file → recipients → message steps, templates, history, settings, updates |
| `src/loader.py` | reads the CSV/Excel list and finds the employee-number column |
| `src/spectrum.py` | looks up names, status and emails in Spectrum |
| `src/mailer.py` | builds each email in Outlook (merge fields, signature) and sends it |
| `src/store.py` | settings, templates, the encrypted login and send history in `%APPDATA%\Remind` |
| `src/updater.py` | checks GitHub Releases and installs a newer version |
| `installer/` | `Install.cmd` / `Install.ps1`, which install for the current user |
