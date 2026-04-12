// ============================
// RUNS WHEN FORM IS SUBMITTED
// ============================
function onFormSubmit(e) {
  try {

    // Make sure function is triggered by actual form submission
    if (!e || !e.response) {
      throw new Error("This function must be run by a Google Form submit trigger.");
    }

    // Convert form answers into an object (question → answer)
    const responses = getResponseMap_(e);

    // Extract each field from form
    const name = (responses["Name"] || "Student").toString().trim();
    const email = (responses["Email"] || "").toString().trim();
    const locker = (responses["Locker"] || "").toString().trim();
    const timeSlot = (responses["Time Slot"] || "").toString().trim();
    const item = (responses["Item"] || "Not specified").toString().trim();

    // Make sure email exists
    if (!email) throw new Error("Missing Email field.");

    // Only allow UIndy emails
    if (!email.toLowerCase().endsWith("@uindy.edu")) {
      Logger.log("Rejected non-UIndy email: " + email);
      return; // stop script
    }

    // Validate required inputs
    if (!locker) throw new Error("Missing Locker field.");
    if (!timeSlot) throw new Error("Missing Time Slot field.");

    // Get the linked response sheet
    const sheet = getResponseSheet_();

    // Get or create important columns
    const statusCol = getOrCreateColumn_(sheet, "Status");       // stores state
    const sentAtCol = getOrCreateColumn_(sheet, "Sent At");      // stores timestamp
    const barcodeCol = getOrCreateColumn_(sheet, "Barcode Value"); // stores barcode

    // Get current row (latest submission)
    const currentRow = sheet.getLastRow();

    // Convert time slot string → start & end Date objects
    const slotInfo = parseTimeSlotForToday_(timeSlot);

    // Get current time
    const now = new Date();

    // ============================
    // CHECK DUPLICATE BOOKINGS
    // ============================
    if (isDuplicateActiveSlot_(sheet, locker, timeSlot, currentRow)) {

      // Mark as rejected
      sheet.getRange(currentRow, statusCol).setValue("REJECTED - SLOT ALREADY TAKEN");

      // Send rejection email
      GmailApp.sendEmail(
        email,
        "UIndy Library Locker Request Rejected",
        "Hi " + name + ",\n\n" +
        "Your request was rejected because " + locker + " is already taken during " + timeSlot + ".\n\n" +
        "Item: " + item + "\n\n" +
        "Please submit again.\n\n" +
        "Thanks,\nUIndy Library Rental"
      );

      return; // stop execution
    }

    // ============================
    // CHECK IF SLOT EXPIRED
    // ============================
    if (now >= slotInfo.end) {

      // Mark rejected
      sheet.getRange(currentRow, statusCol).setValue("REJECTED - TIME SLOT EXPIRED");

      // Send rejection email
      GmailApp.sendEmail(
        email,
        "UIndy Library Locker Request Rejected",
        "Hi " + name + ",\n\n" +
        "Time slot already ended.\n\n" +
        "Locker: " + locker + "\n" +
        "Time Slot: " + timeSlot + "\n" +
        "Item: " + item
      );

      return;
    }

    // ============================
    // SLOT HAS NOT STARTED YET
    // ============================
    if (now < slotInfo.start) {

      // Mark as pending
      sheet.getRange(currentRow, statusCol).setValue("PENDING - NOT STARTED");

      return; // do NOT send yet
    }

    // ============================
    // UPDATED BARCODE LOGIC
    // ============================

    // Generate time-based value AND attach locker name
    // Example result: "12345678-Locker 1"
    const barcodeText = generateBarcodeValueFromTimeSlot_(timeSlot) + "-" + locker;

    // Send barcode email
    sendBarcodeEmail_(email, name, locker, timeSlot, item, barcodeText);

    // Update sheet status
    sheet.getRange(currentRow, statusCol).setValue("SENT");

    // Save timestamp
    sheet.getRange(currentRow, sentAtCol).setValue(new Date());

    // Save barcode value
    sheet.getRange(currentRow, barcodeCol).setValue(barcodeText);

  } catch (err) {

    // Log any errors
    Logger.log("Error in onFormSubmit: " + err.message);
    throw err;
  }
}
