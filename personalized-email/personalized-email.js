Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("send-email-button").onclick = sendEmail;
    }
});

function sendEmail() {
    const status = document.getElementById('status');
    status.textContent = "Functionality not yet implemented.";
    // Placeholder for future functionality
    console.log("Send Email button clicked.");
}
