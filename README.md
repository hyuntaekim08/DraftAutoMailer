# Outlook draft generator

This Python tool reads recipients from a CSV or Excel workbook, fills in
personalized email subjects and bodies, and saves messages in the signed-in
**classic Outlook desktop app**. It does not send messages.

## Requirements

- Windows
- Python 3.10 or later
- Classic Outlook desktop installed and set up with your mailbox
- CSV or Excel workbook (`.xlsx`) with a header row and a column containing
  recipient email addresses

This version uses Outlook's local Windows automation. It does not use
Microsoft Graph, an Entra app registration, a client ID, or a password.
It does not work with only new Outlook or Outlook in a browser.

## Install and start

Run `OutlookDraftMailerSetup.exe` and follow the setup prompts. The installer
bundles the application and its Python dependencies; Python does not need to be
installed on the target computer. Classic Outlook for Windows must still be
installed and configured.

The installer creates a Start menu shortcut and can optionally create a
desktop shortcut. The activity log is saved in
`%LOCALAPPDATA%\Outlook Draft Mailer\mailer.log`.

In the window:

1. Browse to your `.csv` or `.xlsx` file and load it. For Excel files, the
   program reads the first worksheet. If the first row contains column names,
   keep **First row contains column names** checked. If it contains contact
   data, uncheck it before loading; columns will be named `A`, `B`, `C`, and so
   on. The rows and columns appear in a table.
2. Choose the column containing recipient email addresses from the dropdown.
   The file does not need to use a specific name for that column.
3. Check the box in the **Draft?** column for each contact that should receive
   a draft. Use **Select all** or **Clear selection** if helpful.
4. Enter a subject and body. Use column names inside braces, for example
   `{mail}` or `{name}`. You can also select a column and insert its placeholder
   at the cursor.
5. Preview the first checked contact, then choose **Create Outlook drafts**.
   The program creates one draft for each checked row and does not send them.

The **Activity log** at the bottom of the window shows loading, preview, and
draft progress, including errors. The same log is saved as `mailer.log` in the
project folder. **Clear display** clears only the visible log; it does not
delete the log file.

The sample contact in `contacts.csv` uses the reserved `example.com` domain.
Do not create a draft from it without replacing the sample address.

## Build the Windows installer

On a Windows build machine, install Python 3.10 or later and Inno Setup 6.
Open PowerShell in the project folder and run:

```powershell
.\build.ps1
```

The script creates the packaged application in `dist\Outlook Draft Mailer`
and the installer at `dist\installer\OutlookDraftMailerSetup.exe`. It installs
the Python build dependencies into `.venv` as needed. The setup executable
contains the Python runtime and application dependencies, but does not include
Microsoft Outlook.

## Verify the template handling

```powershell
& ".\.venv\Scripts\python.exe" -m unittest -v
```
