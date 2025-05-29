# Excel Web Add-in – Calendar Form System (Multi-Sheet Enabled)

## 📘 Overview

This Excel Web Add-in enables structured data entry tied to specific calendar cells across multiple sheets. Users click a calendar cell, open a taskpane form, and store all input in a centralized internal Excel Table for reporting, searching, and history.  
Ideal for use cases such as test scheduling, appointment booking, internal project logging, etc.

---

## 🛠 Project Summary

| Component    | Description                                                 |
|--------------|-------------------------------------------------------------|
| Platform     | Excel Desktop + Online                                      |
| Host         | Taskpane Add-in                                             |
| Language     | JavaScript (Office.js), HTML, CSS                           |
| Backend      | Internal Excel Table (`tblDataEntries`)                     |
| Hosting      | GitHub Pages (production) + live-server (for local testing) |
| Entry Format | Each row linked to a unique `SheetName!CellAddress`         |

---

## 🛡️ Architecture Summary

| Layer      | Technology                       |
|------------|----------------------------------|
| UI         | HTML, CSS (Office Taskpane)      |
| Logic      | JavaScript (ES6) + Office.js     |
| Data Store | Excel Table (`tblDataEntries`)   |
| Hosting    | GitHub Pages (static HTTPS)      |
| Platform   | Excel Online & Desktop (Office 365) |

---

## 🔄 Identifier Format

Each record is linked to a calendar cell using:

```text
SheetName!CellAddress
```

**Examples:**
- `JanCalendar!F12`
- `February!D3`

This allows support for multiple calendar sheets within one unified table.

---

## 📚 Data Table Specification (`tblDataEntries`)

| Column        | Required | Description                             |
|---------------|----------|-----------------------------------------|
| Identifier    | ✅        | `Sheet!Cell`, e.g., FebCalendar!D3      |
| ClientName    |          | Name of client                          |
| TO            |          | Technical Officer                       |
| CAD           |          | CAD engineer                            |
| JobNumber     |          | Job reference                           |
| CRM           |          | CRM record or note                      |
| TestFee       |          | Test fee as number                      |
| BookingStatus | ✅        | Confirmed / Provisional                 |
| PF            |          | Yes / No                                |
| Duration      |          | Duration of the event                   |
| Description   |          | Description of the task/test            |
| TestStartTime |          | Start time (e.g., 09:00)                |
| Cell          |          | Optional text about cell                |
| TestDate      | ✅        | Date of scheduled test                  |
| DateAdded     | ✅        | Entry creation date                     |
| LastModified  | ✅        | Auto-filled last modified timestamp     |

---

## 📂 File Structure

```text
.
├── manifest.xml
├── README.md
├── taskpane/
│   ├── taskpane.html
│   ├── taskpane.css
│   └── taskpane.js
├── scripts/
│   ├── calendarController.js
│   ├── formController.js
│   ├── worksheetService.js
│   ├── utils.js
│   └── testScript.js (REMOVE for production)
├── assets/icons/
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
```

---

## 📄 Module Breakdown

### `calendarController.js`
- **getSelectedCellIdentifier()** – Returns `Sheet!CellAddress` string for selected cell.
- **updateCalendarCell(identifier, summary, bookingStatus)**:
  - Parses identifier into sheet + cell.
  - Writes formatted summary.
  - Applies background color (`green` = Confirmed, `pink` = Provisional).

### `worksheetService.js`
- **getRowByIdentifier(identifier)** – Searches `tblDataEntries` for matching identifier.
- **saveOrUpdateEntry(entry)** – Writes or updates a table row using identifier as key.
- **deleteEntry(identifier)** – Removes the row with matching identifier from the table.

### `formController.js`
- **loadFormData()** – Reads from selected cell and loads values into form.
- **submitFormData()** – Collects inputs, validates, saves to table, updates calendar cell.
- **cancelTest()** – Deletes entry and clears calendar cell.

### `utils.js`
- **cleanField(value)** – Returns "N/A" if value is blank or null.
- **validateFields(requiredIds)** – Validates presence of required inputs.
- **getFormattedDate(inputId)** – Converts input date field to ISO string.

---

## 📂 UI Structure

### `taskpane.html`
- **14 Inputs / Selects:**
  - Text: ClientName, TO, CAD, JobNumber, CRM, Duration, Description, TestStartTime, Cell
  - Number: TestFee
  - Select: BookingStatus (Confirmed/Provisional), PF (Yes/No)
  - Date: TestDate, DateAdded
- **2 Buttons:**
  - Submit (saves data)
  - Cancel Test (deletes data)
- Modal Popup: Confirms cancellation
- Loading Overlay: Shows during processing

### `taskpane.js`
- Hooks all buttons
- Controls modal and loading UI
- Calls `formController` on load and actions

### `taskpane.css`
- Grid layout
- Modal box
- Loading overlay
- Accessibility and viewport meta

---

## 🔦 Functional Highlights

- Multi-sheet calendar support
- Centralized data in one Excel Table
- Real-time updates to selected cell
- Conditional formatting (color-coding)
- Field-level validation
- Lightweight UI in sidebar (taskpane)
- Excel Online & Desktop compatible
- Compatible with GitHub Pages (static HTTPS)

---

## 🌐 GitHub Pages Deployment

### Steps

1. Push your full folder to a GitHub repository.
2. Enable **GitHub Pages**:
   - Go to `Settings → Pages`
   - Choose branch: `main`, folder: `/ (root)`
3. GitHub will host your project at:
   ```
   https://<your-username>.github.io/<repo-name>/
   ```
4. Update `manifest.xml`:
   - Replace all `http://localhost:3000/...` URLs
   - Use GitHub Pages links (see above)
   - Ensure all icons and `taskpane.html` are hosted

---

## 📄 `manifest.xml` Production Notes

| Field             | Must Be                                  |
|-------------------|------------------------------------------|
| All URLs          | `https://` and hosted (no localhost)     |
| AppDomain (opt.)  | Match GitHub domain                      |
| Icon/HighResIcon  | Use direct GitHub URLs to PNG icons      |
| FunctionFile/Commands | Optional unless using ribbon logic   |
| SourceLocation    | Taskpane HTML on GitHub Pages            |

---

## 🧪 Local Testing Setup

```bash
npm install -g live-server
live-server --port=3000
```

Update `manifest.xml` to:
```
http://127.0.0.1:3000/taskpane/taskpane.html
```
Then:
- Open Excel → Insert → Office Add-ins → Upload My Add-in
- Test all form functions and calendar updates

---

## ✅ Pre-Production Checklist

- [x] Host everything public (GitHub Pages enabled)
- [x] Test `taskpane.html` loads directly in browser

---

## 📌 To-Do List

- [ ] Replace all `localhost` links with GitHub Pages links
- [ ] Add versioning to script files
- [ ] Create GitHub Pages deployment GitHub Action (optional)
- [ ] Add sample screenshots to `README.md`
- [ ] Confirm icons load correctly
- [ ] Test in Excel by uploading `manifest.xml`
- [ ] Add download/export options to the form
- [ ] Add EXAP logic or advanced validation (optional future feature)
- [ ] Submit to Microsoft AppSource (if needed)
- [ ] Minify production JS/CSS (optional)

---
