# Google Cloud & Google Workspace Automation
Extend Google Workspace using Apps Script, REST APIs, and automated workflows.

---

## 1. Automating Emails with Google Apps Script (`Mail_from_Almaty.json`)

This project demonstrates how anyone—even without a traditional software engineering background—can quickly automate email workflows in Google Workspace using **Google Apps Script** and the built-in `MailApp` service.

📺 **Video Tutorial:** [Automating emails via Appscript Google script](https://www.youtube.com/watch?v=j9Iggaiw02I)

### What's Inside [`Mail_from_Almaty.json`](./Mail_from_Almaty.json)
* **Apps Script Manifest (`appsscript.json`)**: Pre-configured for the `Asia/Almaty` timezone on the Google Apps Script **V8 runtime** with Stackdriver exception logging enabled.
* **`myEmailHi()` Function (`Code.gs`)**: A starter encapsulation template using `MailApp.sendEmail(recipient, subject, body)` to trigger automated emails directly from Google Drive / Google Workspace, or to connect as a webhook/action inside **Google AppSheet** via the Apps Script connector.

```javascript
function myEmailHi() {
  MailApp.sendEmail(
    "Put your email here",
    "Hello from Almaty",
    "Hello, thank you. Kind regards, Askar!"
  );
}
```

### How to Use (Step-by-Step from the Video)
1. **Launch Google Apps Script**: Open your **Google Drive** -> click **New** -> **More** -> **Google Apps Script** (`01:00` in the video).
2. **Create or Import the Script**:
   * Copy the `myEmailHi()` function from [`Mail_from_Almaty.json`](./Mail_from_Almaty.json) into the `Code.gs` editor (`01:30`), **or** import the JSON project using [`clasp`](https://github.com/google/clasp) / the Apps Script API.
3. **Configure Parameters**: Replace `"Put your email here"` with the target recipient email address, and customize the subject and body (`02:20`).
4. **Authorize & Run**: Click **Run** in the Apps Script editor, grant the required `MailApp` permissions, or connect the function to an **AppSheet** automation bot.

---

## 2. Web Scraping & Cloud Skills Boost Badge Verification (`sup.js` & `main.js`)

Developed by **Dmitrii Kainekin** based on business requirements and architecture by **Askar Aituov**. Together, these scripts automate deduplicating student records in Google Sheets and scraping public Google Cloud Skills Boost profiles via `UrlFetchApp` to count earned badges.

### Workflow & Deployment Order
1. **Step 1 — Deduplicate Data (`sup.js`)**:
   * Deploy and run `removeDuplicatesAndCopyToNewSheet()` in [`sup.js`](./sup.js) first.
   * Reads raw form responses from the `"Base"` sheet, filters out duplicate entries by `"Your student ID"`, and outputs clean records to the `"test"` sheet.
2. **Step 2 — Scrape & Extract Badge Counts (`main.js`)**:
   * Run `extractDataAndCreateNewSheet()` in [`main.js`](./main.js).
   * Reads the deduplicated `"test"` sheet, validates and normalizes each student's Cloud Skills Boost public profile URL, fetches the HTML content via `UrlFetchApp.fetch()`, counts the earned badges, and writes the final report to the `"test_result"` sheet.
