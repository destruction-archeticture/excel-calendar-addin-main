// calendarController.js

export async function getSelectedCellIdentifier() {
    console.log("[Calendar] Getting selected cell identifier...");
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = context.workbook.getSelectedRange();
        sheet.load("name");
        range.load("address");
        await context.sync();

        const identifier = `${sheet.name}!${range.address}`;
        console.log("[Calendar] Identifier resolved:", identifier);
        return identifier;
    });
}

export async function updateCalendarCell(identifier, summary, bookingStatus) {
    console.log("[Calendar] Updating cell:", identifier);
    const [sheetName, cellAddress] = identifier.split("!");
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const range = sheet.getRange(cellAddress);

        console.log("[Calendar] Writing summary to cell:", summary);
        range.values = [[summary]];

        switch (bookingStatus) {
            case "Confirmed":
                range.format.fill.color = "#c6efce"; // light green
                console.log("[Calendar] Applied green for 'Confirmed'");
                break;
            case "Provisional":
                range.format.fill.color = "#fce4d6"; // light pink
                console.log("[Calendar] Applied pink for 'Provisional'");
                break;
            default:
                range.format.fill.clear();
                console.log("[Calendar] Cleared fill color (no status)");
        }

        await context.sync();
        console.log("[Calendar] Cell update complete");
    });
}
