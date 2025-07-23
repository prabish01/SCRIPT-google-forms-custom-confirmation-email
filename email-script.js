function onFormSubmit(e) {
  const sheet = e.source.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = e.values;
  const data = Object.fromEntries(headers.map((h, i) => [h, values[i]]));
  const email = data["Email Address"];
  const fullName = data["Full Name"];
  const location = data["Location"];
  const amount = data["Amount"];
  const repayAmount = data["Repayment Amount"];
  const repayDue = data["Repayment Due Date"];
  const paymentMethod = data["Payment Method"];
  const status = data["Status"] || "REQ";
  // Construct message
  let message = `
Dear ${fullName},
This statement is to declare that ${fullName}, who lives in ${location}, has applied for a loan of $${amount}. 
They have agreed to repay $${repayAmount} by ${repayDue} via ${paymentMethod}. 
Their loan request has been recorded with the status "${status}".
Submitted Files:
`;
  // Attach uploaded file links with custom labels
  const fileFields = [
    "Close Up Of ID (Front & Back) *",
    "Most Recent Pay-Stubs (2/Two) *",
    "Upload A Selfie While Holding Your ID.",
    "Reddit Username + ID And Todays Date On A Piece Of Paper (Make Sure It's Visible Please) *",
    "Upload A Proof Of Address. (Utility bill, internet bill etc..) *"
  ];
  for (let field of fileFields) {
    if (data[field]) {
      message += `\n• ${field}: ${data[field]}`;
    }
  }
  message += `\n\nIf any of the above information is incorrect, please contact us immediately.\n\nBest regards,\nLoan Processing Team`;
  // Send Email
  GmailApp.sendEmail(email, `Loan Application Confirmation - ${fullName}`, message);
}
