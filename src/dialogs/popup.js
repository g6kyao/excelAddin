Office.onReady((info) => {
    // TODO1: Assign handler to the OK button.
    document.getElementById("ok-button").onclick = () => tryCatch(sendStringToParentPage);
});

// TODO2: Create the OK button handler.
function sendStringToParentPage() {
    const userName = document.getElementById("name-box").value;
    Office.context.ui.messageParent(userName);
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}