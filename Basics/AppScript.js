function doGet() {
    return HtmlService.createHtmlOutputFromFile('reimbursement')
        .setTitle('Ministry Expense Reimbursement Form');
}

function sendEmail(to, subject, body) {
    MailApp.sendEmail({
        to: to,
        subject: subject,
        body: body
    });
}

function processForm(formData, fileData) {
    try {
        Logger.log("Received form data: " + JSON.stringify(formData));

        var ss = SpreadsheetApp.openById("137QVV0PwLvwzG_XUJZ3lAAs4yinl5pxxUM5VMu9xI1A");
        var sheet = ss.getSheetByName("Reimbursement Data") || ss.insertSheet("Reimbursement Data");

        if (!sheet) {
            throw new Error("Sheet not found. Check the sheet name.");
        }

        if (sheet.getLastRow() === 0) {
            sheet.appendRow(["Timestamp", "Request ID", "Name", "Email", "UPI", "Ministry", "Expense Category", "Amount", "Justification", "Payment Mode", "Expense Date", "Receipt Link", "Approval Status"]);
        }

        var timestamp = new Date();
        var requestId = "REQ-" + timestamp.getTime();
        Logger.log("Generated Request ID: " + requestId);
        var folder = DriveApp.getFolderById("1FKeBJtE5ul8vbygraVxk0z_Erdkb7PjC");


        let baseRow = [timestamp, requestId, formData.name, formData.email, formData.pocUpi, formData.ministry];

        formData.expenses.forEach(expense => {
            let receiptLink = "";

            if (expense.receipt) {
                let blob = Utilities.newBlob(Utilities.base64Decode(expense.receipt.data), expense.receipt.type, expense.receipt.name);
                let file = folder.createFile(blob);
                receiptLink = file.getUrl();
            }
        
        sheet.appendRow([...baseRow, expense.category, expense.amount, expense.justification, expense.paymentMode, expense.expenseDate, receiptLink, "Pending Approval"]);
        });

        Logger.log("Data appended successfully.");
        sendEmail(formData.email, "Reimbursement Request Submitted", 
            "Your request (ID: " + requestId + ") has been received. You will be updated soon.");

        notifyTreasurer(requestId, formData.ministry);

        return { status: "success", requestId: requestId };

    } catch (error) {
        Logger.log("Error in processForm: " + error.message);
        return { status: "error", message: error.message };
    }
}

function notifyTreasurer(requestId, ministry) {
    var treasurerEmail = "tanvee.jitendra_ug2024@ashoka.edu.in"; // Treasurer's email
    var subject = "New Reimbursement Request - " + requestId;
    var body = "A new reimbursement request (Request ID: " + requestId + ") has been submitted for review.\n\n" +
               "Ministry: " + ministry + "\n\n" +
               "Please log in to the reimbursement system and review the request for approval.";

    sendEmail(treasurerEmail, subject, body);
}

// Global variable for the sheet ID
var SHEET_ID = "137QVV0PwLvwzG_XUJZ3lAAs4yinl5pxxUM5VMu9xI1A";

