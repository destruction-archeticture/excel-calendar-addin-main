// worksheetService.js

const TABLE_NAME = "tblDataEntries";

export async function getRowByIdentifier(identifier) {
    console.log("[Data] Searching for identifier:", identifier);
    return await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(TABLE_NAME);
        const dataBodyRange = table.getDataBodyRange();
        dataBodyRange.load(["values", "rowIndex"]);
        table.load("columns");
        await context.sync();

        const headers = table.columns.items.map(col => col.name);
        const identifierIndex = headers.indexOf("Identifier");

        for (let i = 0; i < dataBodyRange.values.length; i++) {
            const row = dataBodyRange.values[i];
            if (row[identifierIndex] === identifier) {
                console.log(`[Data] Match found at row ${i}:`, row);
                return { rowIndex: i, values: row };
            }
        }

        console.log("[Data] No match found for identifier");
        return null;
    });
}

export async function saveOrUpdateEntry(entry) {
    console.log("[Data] Saving or updating entry:", entry.Identifier);
    return await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(TABLE_NAME);
        table.load("columns");
        await context.sync();

        const headers = table.columns.items.map(col => col.name);
        const rowData = headers.map(h => entry[h] || "");
        console.log("[Data] Constructed rowData array:", rowData);

        const existing = await getRowByIdentifier(entry.Identifier);
        if (existing) {
            console.log("[Data] Updating existing row at index:", existing.rowIndex);
            const dataRange = table.getDataBodyRange();
            const targetRow = dataRange.getRow(existing.rowIndex);
            targetRow.values = [rowData];
        } else {
            console.log("[Data] Adding new row");
            table.rows.add(null, [rowData]);
        }

        await context.sync();
        console.log("[Data] Save or update complete");
    });
}

export async function deleteEntry(identifier) {
    console.log("[Data] Deleting entry with identifier:", identifier);
    return await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(TABLE_NAME);
        const dataRange = table.getDataBodyRange();
        dataRange.load("values");
        table.load("columns");
        await context.sync();

        const headers = table.columns.items.map(col => col.name);
        const identifierIndex = headers.indexOf("Identifier");

        for (let i = 0; i < dataRange.values.length; i++) {
            const row = dataRange.values[i];
            if (row[identifierIndex] === identifier) {
                console.log("[Data] Found entry at row", i, "- deleting");
                table.rows.getItemAt(i).delete();
                break;
            }
        }

        await context.sync();
        console.log("[Data] Deletion complete");
    });
}
