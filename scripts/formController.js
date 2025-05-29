import {
    getRowByIdentifier,
    saveOrUpdateEntry,
    deleteEntry,
} from "./worksheetService.js";

import {
    getSelectedCellIdentifier,
    updateCalendarCell,
} from "./calendarController.js";

import {
    cleanField,
    validateFields,
    getFormattedDate,
} from "./utils.js";

export async function loadFormData() {
    console.log("[Form] Starting form data load...");
    const identifier = await getSelectedCellIdentifier();
    console.log("[Form] Selected cell identifier:", identifier);
    if (!identifier) return;

    const existing = await getRowByIdentifier(identifier);
    if (!existing) {
        console.log("[Form] No existing entry found for identifier.");
        return;
    }

    console.log("[Form] Existing entry found:", existing.values);

    const fields = [
        "ClientName", "TO", "CAD", "JobNumber", "CRM", "TestFee", "BookingStatus",
        "PF", "Duration", "Description", "TestStartTime", "Cell", "TestDate", "DateAdded"
    ];

    const values = existing.values;

    for (let i = 0; i < fields.length; i++) {
        const el = document.getElementById(fields[i]);
        if (el && values[i] !== undefined) {
            el.value = values[i] !== "N/A" ? values[i] : "";
        }
    }

    console.log("[Form] Form populated successfully.");
}

export async function submitFormData() {
    console.log("[Submit] Starting submission process...");
    const identifier = await getSelectedCellIdentifier();
    console.log("[Submit] Selected cell identifier:", identifier);
    if (!identifier) return;

    const required = ["BookingStatus", "TestDate", "DateAdded"];
    if (!validateFields(required)) {
        console.warn("[Submit] Required fields missing – submission aborted.");
        return;
    }

    const entry = {
        Identifier: identifier,
        ClientName: cleanField(document.getElementById("ClientName").value),
        TO: cleanField(document.getElementById("TO").value),
        CAD: cleanField(document.getElementById("CAD").value),
        JobNumber: cleanField(document.getElementById("JobNumber").value),
        CRM: cleanField(document.getElementById("CRM").value),
        TestFee: cleanField(document.getElementById("TestFee").value),
        BookingStatus: document.getElementById("BookingStatus").value,
        PF: cleanField(document.getElementById("PF").value),
        Duration: cleanField(document.getElementById("Duration").value),
        Description: cleanField(document.getElementById("Description").value),
        TestStartTime: cleanField(document.getElementById("TestStartTime").value),
        Cell: cleanField(document.getElementById("Cell").value),
        TestDate: getFormattedDate("TestDate"),
        DateAdded: getFormattedDate("DateAdded"),
        LastModified: new Date().toISOString()
    };

    console.log("[Submit] Prepared entry object:", entry);

    await saveOrUpdateEntry(entry);
    console.log("[Submit] Entry saved to worksheet");

    const summary = `${entry.ClientName} - ${entry.TO}/${entry.CAD} - ${entry.Description} - ${entry.TestStartTime} - ${entry.Duration} - ${entry.CRM}`;
    await updateCalendarCell(identifier, summary, entry.BookingStatus);
    console.log("[Submit] Calendar cell updated");
}

export async function cancelTest() {
    console.log("[Cancel] Starting test cancellation...");
    const identifier = await getSelectedCellIdentifier();
    console.log("[Cancel] Identifier to cancel:", identifier);
    if (!identifier) return;

    await deleteEntry(identifier);
    console.log("[Cancel] Entry deleted from worksheet");

    await updateCalendarCell(identifier, "", "");
    console.log("[Cancel] Calendar cell cleared");
}
