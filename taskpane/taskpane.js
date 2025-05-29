import {
    loadFormData,
    submitFormData,
    cancelTest,
} from "../scripts/formController.js";

Office.onReady(() => {
    console.log("[Init] Office ready – Initializing taskpane");

    document.getElementById("submitBtn").onclick = () => {
        console.log("[Action] Submit button clicked");
        showLoading(true);
        submitFormData()
            .then(() => console.log("[Success] Form submitted successfully"))
            .catch(err => console.error("[Error] Form submission failed:", err))
            .finally(() => showLoading(false));
    };

    document.getElementById("cancelBtn").onclick = () => {
        console.log("[Action] Cancel button clicked");
        showModal("Are you sure you want to cancel this test?");
    };

    document.getElementById("modalYes").onclick = () => {
        console.log("[Action] Modal confirmed: proceeding with cancellation");
        hideModal();
        showLoading(true);
        cancelTest()
            .then(() => console.log("[Success] Test cancelled successfully"))
            .catch(err => console.error("[Error] Test cancellation failed:", err))
            .finally(() => showLoading(false));
    };

    document.getElementById("modalNo").onclick = () => {
        console.log("[Action] Modal dismissed: cancellation aborted");
        hideModal();
    };

    console.log("[Load] Loading form data from selection");
    showLoading(true);
    loadFormData()
        .then(() => console.log("[Success] Form data loaded"))
        .catch(err => console.error("[Error] Failed to load form data:", err))
        .finally(() => showLoading(false));
});

function showLoading(visible) {
    console.log(`[UI] Loading overlay: ${visible ? "shown" : "hidden"}`);
    document.getElementById("loadingOverlay").classList.toggle("hidden", !visible);
}

function showModal(message) {
    console.log(`[UI] Showing modal: "${message}"`);
    document.getElementById("modalMessage").textContent = message;
    document.getElementById("modalConfirm").classList.remove("hidden");
}

function hideModal() {
    console.log("[UI] Hiding modal");
    document.getElementById("modalConfirm").classList.add("hidden");
}
