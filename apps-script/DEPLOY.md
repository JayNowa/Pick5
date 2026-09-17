# Deploying Code.gs to Google Apps Script

> **Important:** Always use "Manage deployments" to update an existing deployment.
> Creating a "New deployment" generates a different URL, which breaks the app.

## Steps

1. Open [Google Apps Script](https://script.google.com) and open the Pick 5 project.
2. Copy the contents of `Code.gs` and paste it into the editor, replacing everything.
3. Click **Save** (floppy disk icon or Ctrl+S).
4. Click **Deploy** → **Manage deployments**.
5. Click the **pencil (edit) icon** next to the active deployment.
6. Under "Version", select **New version** from the dropdown.
7. Click **Deploy**.
8. Done — the deployment URL stays the same, no changes needed in the app.

## Verify it worked

- The version number in `doPost` logs (e.g. `doPost called - v12`) will reflect the latest version.
- To test, run one of the test functions (`testConfirmation`, `testReminder`, `testRecap`) directly from the editor and check your inbox.

## Deployment URL

```
https://script.google.com/macros/s/AKfycbw6dH6UV1BVH6Nzm_z_BwDi6oecp4AGcJ-0OkJhhUzSS07QD1cb84LCHjhQiSSremAd6Q/exec
```

This URL is also set in `worker.js` (`APPS_SCRIPT_URL`) and `index.html` (`PICKS_SCRIPT_URL`).
If it ever changes (from accidentally creating a new deployment), update it in both of those files.
