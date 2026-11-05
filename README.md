# Albrost Accounts

A simple app for a restaurant's daily sales, expenses and supplier balances. It runs on the shop computer and saves everything into one Excel file. People open it from a web link.

> **Login is off for now.** Anyone with the link has full access, so keep the link private. Logins and the manager role below start working once login is turned on (planned for December).

> **This repo only holds the app's code.** The Excel file, backups, logs and passwords stay on the shop computer and are never uploaded.

## Who can do what

| | Owner | Manager |
|---|---|---|
| Dashboard, Supplier Balances and All Records | ✅ | ✅ |
| Add a daily sale or supplier entry | ✅ | ✅ |
| Edit or delete saved entries | ✅ | ❌ |
| "Last updated" times in All Records | ✅ | ❌ |
| Logs | ✅ | ❌ |

## Run it online

You need [Python](https://www.python.org/downloads/) and [cloudflared](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/) installed.

1. Install the one dependency, once:

   ```bash
   pip install -r requirements.txt
   ```

2. Put `Al_Brost_Finance_Supplier_Balances.xlsx` in this folder.

3. To use logins: remove `rem ` from the `set ALBROST_LOGIN=on` line in **START ONLINE.bat**, then create a login for each person. It asks for the password:

   ```bash
   python set_password.py USERNAME owner "Display Name"
   ```

   Use `manager` instead of `owner` for a manager. Run the same command again to change a password.

4. Double-click **START ONLINE.bat**. The link appears in the window as `https://....trycloudflare.com`.

The link only works while that computer is on and the window is open. Each restart gives a new link.

## Run it on the shop computer only

Double-click **START ALBROST APP.bat**. This has no login and is not on the internet. See `HOW TO USE.txt`. Changes made this way, or directly in Excel, show in Logs as "Shop PC / Excel" the next time the online app is opened.

## Where things are

| Path | What it is |
|---|---|
| `server.py` | Online server: logins, roles, logs, saving |
| `set_password.py` | Adds logins and changes passwords |
| `web/` | Sign-in page and the owner and manager screens |
| `app/` | The main app page and the shop computer version |

Created when it runs, and never uploaded:

| Path | What it is |
|---|---|
| `data/` | Password hashes and the activity log |
| `Backups/` | A copy from every save, plus the first copy of each day in `Backups/Daily` |

## Restore an old copy

Stop the server, copy the file you want from `Backups` over `Al_Brost_Finance_Supplier_Balances.xlsx`, then start it again.

## Updates to this repo

While the server is running, it checks for code changes every minute and pushes them here. Only the files listed in `.gitignore` can ever be uploaded.

## Questions

Open an issue on this repo, or ask the owner.
