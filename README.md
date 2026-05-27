# Fateh Time Check-In

A simple mobile-friendly worker check-in/check-out page for Fateh Plumbing & Electric. It can be published as a public GitHub Pages site while keeping worker PINs inside Google Sheets / Apps Script.

## What Workers Do

1. Open the site on their phone.
2. Enter their name and PIN.
3. Optionally enter the job and notes.
4. Allow location.
5. Tap **Check In** or **Check Out**.

The app records the local time, GPS location, action, worker name, optional job, and optional notes.

## Google Sheet Setup

1. Create a new Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Delete the starter code and paste in `google-apps-script/Code.gs`.
4. Click **Save**.
5. Select the `setup` function, click **Run**, and allow permissions. This creates the sheets.
6. In the `Workers` sheet, add workers like this:

| Name | PIN | Active |
| --- | --- | --- |
| Jas | 4321 | TRUE |
| Aman | 7788 | TRUE |

Use the exact name the worker will type. Set `Active` to `FALSE` to disable a worker.

## Deploy Apps Script

1. In Apps Script, click **Deploy > New deployment**.
2. Choose **Web app**.
3. Set **Execute as** to **Me**.
4. Set **Who has access** to **Anyone**.
5. Click **Deploy** and copy the Web App URL.
6. Paste that URL into `config.js`:

```js
window.FATEH_TIME_CONFIG = {
  scriptUrl: "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec",
};
```

## Publish on GitHub Pages

1. Push these files to a public GitHub repository.
2. In GitHub, open **Settings > Pages**.
3. Set the source to the main branch and root folder.
4. Send the GitHub Pages link to workers.

## Important Notes

- PINs are not stored in the public website code.
- GPS requires HTTPS, so GitHub Pages is a good fit.
- The Google Sheet is the source of truth for workers and time logs.
- Job and notes are optional.
