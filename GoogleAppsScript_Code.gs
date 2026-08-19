const SHEET_NAME = "Orders";
const PHOTO_FOLDER_NAME = "GoozDelish Order Reference Photos";

function getOrCreatePhotoFolder_() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "Submitted At", "Name", "Contact", "How Found Us",
        "Flavour", "Options", "Order Notes", "Agreement", "Photo Reference"
      ]);
    }

    let photoUrl = "";
    if (data.photoData) {
      const folder = getOrCreatePhotoFolder_();
      const bytes = Utilities.base64Decode(data.photoData);
      const blob = Utilities.newBlob(bytes, data.photoType || "image/jpeg", data.photoName || "reference-photo.jpg");
      const file = folder.createFile(blob);
      photoUrl = file.getUrl();
    }

    sheet.appendRow([
      new Date(),
      data.name || "",
      data.contact || "",
      data.source || "",
      data.flavour || "",
      data.options || "",
      data.notes || "",
      data.agreement || "",
      photoUrl
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({success:true, photoUrl:photoUrl}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({success:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
