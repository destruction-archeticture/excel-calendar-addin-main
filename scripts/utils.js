// utils.js

export function cleanField(value) {
    if (value === null || value === undefined || value.trim() === "") {
        console.log("[Utils] Field cleaned → output: 'N/A'");
        return "N/A";
    }
    const cleaned = value.trim();
    console.log("[Utils] Field cleaned → output:", cleaned);
    return cleaned;
}

export function validateFields(requiredIds) {
    const missing = [];

    requiredIds.forEach(id => {
        const el = document.getElementById(id);
        if (!el || el.value.trim() === "") {
            missing.push(id);
        }
    });

    if (missing.length > 0) {
        alert("The following required fields are missing:\n\n" + missing.join("\n"));
        console.warn("[Validation] Missing required fields:", missing);
        return false;
    }

    console.log("[Validation] All required fields are present");
    return true;
}

export function getFormattedDate(inputId) {
    const el = document.getElementById(inputId);
    if (!el || !el.value) {
        console.warn(`[Utils] Date input '${inputId}' is missing or empty`);
        return null;
    }

    const date = new Date(el.value);
    if (isNaN(date.getTime())) {
        console.warn(`[Utils] Invalid date in input '${inputId}'`);
        return null;
    }

    const iso = date.toISOString();
    console.log(`[Utils] Date input '${inputId}' converted to ISO:`, iso);
    return iso;
}
