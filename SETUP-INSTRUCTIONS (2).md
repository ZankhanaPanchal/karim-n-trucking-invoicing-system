# Karim N Trucking — POD System Setup

Two parts, both free: a **Google Sheet backend** (the dispatch dashboard) and **GitHub Pages hosting** (the public link drivers use). About 15 minutes total, one-time.

\---

## Part 1 — Google Sheet + Apps Script backend

**1. Create the Sheet**

* Go to [sheets.google.com](https://sheets.google.com) → Blank spreadsheet.
* Name it something like `KNT POD Tracker`.

**2. Attach the script**

* In the Sheet, go to **Extensions → Apps Script**.
* Delete anything in the editor, then paste in the entire contents of `Code.gs` (provided alongside this file).
* Click the disk icon (or Ctrl/Cmd+S) to save. Name the project `KNT POD Backend`.

**3. Deploy as a Web App**

* Click **Deploy → New deployment**.
* Click the gear icon next to "Select type" → choose **Web app**.
* Settings:

  * Description: `POD intake v1`
  * Execute as: **Me**
  * Who has access: **Anyone**
* Click **Deploy**.
* The first time, Google will ask you to authorize the script — click **Authorize access**, pick your Google account, click **Advanced → Go to KNT POD Backend (unsafe)** (this warning is normal for your own scripts), then **Allow**.
* Copy the **Web app URL** it gives you (ends in `/exec`). This is your `SCRIPT\\\_URL`.

**4. Wire it into the app**

* Open `karim-n-trucking-pod-system.html` in any text editor.
* Near the top, find `CONFIG.SCRIPT\\\_URL` and paste your Web App URL between the quotes.
* Save the file.

**5. Restrict the dashboard to dispatch only**

* Back in the Google Sheet, click **Share** (top right).
* Add only the email addresses of your dispatch staff, with **Viewer** or **Editor** access as you prefer.
* Turn off "Anyone with the link" if it's on — this keeps the data private to the people you name.
* Tabs are created automatically the first time a POD lands in them, named by type and delivery month — e.g. **`TL - Jul 2026`**, **`BOL - Jul 2026`**, **`RM - Aug 2026`**. Each one fills in with a row per ticket, a photo thumbnail preview, and clickable Drive links to the full-size photos. You'll end up with up to 4 tabs per month (fewer if a type wasn't used that month).

> Note: the Web App itself runs "as you," so it can write to the Sheet even though drivers never see or access the Sheet directly — they only interact with the HTML form.

**Re-deploying after edits:** if you ever change `Code.gs`, use **Deploy → Manage deployments → edit (pencil) → New version → Deploy**. Editing the script alone doesn't update the live URL's behavior until you do this.

\---

## Part 2 — Publish the app for drivers (GitHub Pages, free)

**1. Create a GitHub account** at [github.com](https://github.com) if you don't have one.

**2. Create a new repository**

* Click **+ → New repository**.
* Name it e.g. `knt-pod-system`. Set it to **Public**. Click **Create repository**.

**3. Upload the file**

* Click **Add file → Upload files**.
* Upload `karim-n-trucking-pod-system.html`, but **rename it to `index.html`** during upload (click the filename to edit it).
* Commit the changes.

**4. Turn on Pages**

* Go to the repo's **Settings → Pages**.
* Under "Build and deployment," set Source to **Deploy from a branch**, branch `main`, folder `/ (root)`. Save.
* Wait \~1 minute, then refresh — GitHub shows your live URL, something like:
`https://yourusername.github.io/knt-pod-system/`

**5. Share that link with drivers**

* Text or WhatsApp the link to your drivers. They can add it to their phone's home screen (Safari/Chrome → Share → Add to Home Screen) so it opens like an app.

**Updating the app later:** edit `index.html` in the repo (or re-upload it) and commit — GitHub Pages redeploys automatically within a minute.

\---

## How it all fits together

```
Driver's phone (public link, no login)
        │  fills form, submits
        ▼
Google Apps Script Web App  (CONFIG.SCRIPT\\\_URL)
        │  writes row + saves photos
        ▼
Google Sheet — tabs "TL - Jul 2026" / "BOL - Jul 2026" / etc.  ←──  shared only with dispatch (the dashboard)
        +
Google Drive folder "KNT POD Photos"  (photo files, linked from the sheet)
```

* Drivers never see the Sheet — they only ever see the submission form.
* Dispatch works entirely inside the Google Sheet: sort, filter, add a pivot table, color-code by status, whatever you need. It updates in real time as drivers submit.
* If a driver taps "Update an existing POD" and enters a ticket number, the app now checks the Sheet directly (not just that phone's memory), so updates work from any device.

