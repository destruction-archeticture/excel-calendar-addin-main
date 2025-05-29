// testScript.js

import {
    saveOrUpdateEntry,
    getRowByIdentifier,
    deleteEntry
} from "./worksheetService.js";

import { updateCalendarCell } from "./calendarController.js";

// Sheet + cell to simulate
const sheetName = "FebCalendar";
const cellAddress = "F12";
const identifier = `${sheetName}!${cellAddress}`;

const entry = {
    Identifier: identifier,
    ClientName: "Demo Co.",
    TO: "Alice",
    CAD: "Bob",
    JobNumber: "JOB-TEST",
    CRM: "CRM-999",
    TestFee: "1200",
    BookingStatus: "Confirmed",
    PF: "Yes",
    Duration: "3h",
    Description: "Simulated Entry",
    TestStartTime: "10:30",
    Cell: "North",
    TestDate: new Date().toISOString(),
    DateAdded: new Date().toISOString(),
    LastModified: new Date().toISOString()
};

// Utility to ensure table + sheet exist
async function ensureEnvironment() {
    return await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        const sheetExists = sheets.items.some(s => s.name === sheetName);
        if (!sheetExists) sheets.add(sheetName);

        const dataSheet = sheets.getItemOrNullObject("DataSheet");
        await context.sync();

        if (dataSheet.isNullObject) {
            const newSheet = sheets.add("DataSheet");
            const headers = [[
                "Identifier", "ClientName", "TO", "CAD", "JobNumber", "CRM", "TestFee", "BookingStatus",
                "PF", "Duration", "Description", "TestStartTime", "Cell", "TestDate", "DateAdded", "LastModified"
            ]];
            const range = newSheet.getRange("A1:Q1");
            range.values = headers;
            newSheet.tables.add("A1:Q1", true).name = "tblDataEntries";
        }

        await context.sync();
    });
}

async function runFullSimulation() {
    console.log("🔁 Preparing environment...");
    await ensureEnvironment();

    console.log("✅ Environment ready.");
    await saveOrUpdateEntry(entry);
    console.log("✅ Entry saved to table.");

    const summary = `${entry.ClientName} - ${entry.Description}`;
    await updateCalendarCell(identifier, summary, entry.BookingStatus);
    console.log("✅ Calendar updated.");

    const loaded = await getRowByIdentifier(identifier);
    console.log("✅ Loaded from table:", loaded);

    await deleteEntry(identifier);
    console.log("✅ Entry deleted.");

    await updateCalendarCell(identifier, "", "");
    console.log("✅ Calendar cell cleared.");
}

Office.onReady(() => {
    setTimeout(runFullSimulation, 1500);
});
