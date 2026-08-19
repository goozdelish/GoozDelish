const SHEET_ID = "1-qQ2W6g7bAk_nCkn_d4HwDW0dh4oMNi-6oZzlvLtelU";
const SHEET_NAME = "Sheet1";

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = JSON.parse(e.postData.contents);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Date Submitted",
        "Name",
        "How Found Us",
        "Contact Information",
        "Servings",
        "Pickup Date",
        "Pickup Time",
        "Cake Message",
        "Flavour",
        "Option",
        "Order Details",
        "Reference Photo",
        "Home Bakery Agreement"
      ]);
    }

    let photoUrl = "";

    if (data.photo && data.photo.data) {
      const folder = getOrCreateFolder_("GoozDelish Order Reference Photos");
      const bytes = Utilities.base64Decode(data.photo.data);
      const blob = Utilities.newBlob(
        bytes,
        data.photo.mimeType || "image/jpeg",
        data.photo.fileName || "reference-photo.jpg"
      );
      const file = folder.createFile(blob);
      photoUrl = file.getUrl();
    }

    sheet.appendRow([
      new Date(),
      data.name || "",
      data.foundUs || "",
      data.contact || "",
      data.servings || "",
      data.pickupDate || "",
      data.pickupTime || "",
      data.cakeMessage || "",
      data.flavour || "",
      data.option || "",
      data.details || "",
      photoUrl,
      data.agreement ? "Agreed" : ""
    ]);

    return jsonResponse_({success:true,message:"Order saved successfully."});

  } catch (error) {
    return jsonResponse_({success:false,message:error.toString()});
  }
}

function getOrCreateFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
