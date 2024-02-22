function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Scripts")
    .addItem("Move Prep Sheet to Folder", "movePrepSheet")
    .addToUi();
}

// Define the column numbers for the Form Responses sheet
const responseCols = {
  "week": 0,
  "section": 1,
  "date": 2,
  "prepSheetLink": 3,
  "quality": 4,
  "folder": 5,
  "notes": 6,
}

/**
 * Function to get values from the Settings sheet
 * @param {string} key
 * @returns {string}
 */
function getSetting(key) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = SS.getSheetByName("Settings");
  const data = settingSheet.getRange(1, 1).getDataRegion().getValues().slice(1);
  let value = data.filter(x => x[0] === key)[0][1];
  return value;
}

/**
 * Prompt the user with a yes or no question
 * If yes, the function returns
 * If no, the function throws an error
 * @param {string} prompt
 * @param {string} errorMsg
 */
function confirmationPrompt(prompt, errorMsg) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(prompt, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    return;
  }
  else {
    throw new Error(errorMsg);
  }
}

/**
 * @param {string} url
 * @returns {DriveApp.File}
 */
function getDriveFileByUrl(url) {
  const id = url.match(/[-\w]{25,}/);
  return DriveApp.getFileById(id);
}

/**
 * @param {string} url
 * @returns {DriveApp.Folder}
 */
function getDriveFolderByUrl(url) {
  const id = url.match(/[-\w]{25,}/);
  return DriveApp.getFolderById(id);
}

/**
 * Try to find a child folder with the desired name
 * If it exists, return that folder
 * Otherwise, create such a folder and return it
 * @param {DriveApp.Folder} parent
 * @param {string} name
 * @returns {DriveApp.Folder}
 */
function getChildFolder(parent, name) {
  const children = parent.getFolders();
  while (children.hasNext()) {
    let folder = children.next();
    if (folder.getName() === name) {
      return folder;
    }
  }
  let newFolder = parent.createFolder(name);
  return newFolder;
}

// Function to move the prep sheet to the proper folder, after confirming with the user
function movePrepSheet() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  // Open the Form Responses sheet, making sure it is the active sheet
  const responseSheet = SS.getActiveSheet();
  if (responseSheet.getSheetName() != "Form Responses") {
    throw new Error("The 'Form Responses' sheet must be active");
  }
  // Get the active row, and check that it is not the headers
  const activeRowIdx = responseSheet.getCurrentCell().getRow();
  if (activeRowIdx === 1) {
    throw new Error("Select a row besides the header")
  }
  const activeRow = responseSheet.getRange(activeRowIdx, 1, 1, Object.keys(responseCols).length);
  const activeRowValues = activeRow.getValues()[0];

  // Check that the prep sheet quality has been reviewed
  const quality = activeRowValues[responseCols.quality];
  if (quality === "") {
    throw new Error("The prep sheet quality has not been evaluated");
  }
  // Check that the folder has not been set
  if (activeRowValues[responseCols.folder] !== "") {
    throw new Error("The prep sheet has already been moved to a folder");
  }
  // Check that the file name is correctly formatted
  const prepSheetUrl = activeRowValues[responseCols.prepSheetLink];
  const prepSheet = getDriveFileByUrl(prepSheetUrl);
  confirmationPrompt(`Is the following file name correctly formatted?\n${prepSheet.getName()}`, "Please update the file name")
  // Check that the subject is correct
  const section = activeRowValues[responseCols.section];
  const subject = section.split("(")[0].trim();
  confirmationPrompt(`Is the following subject correct?\n${subject}`, "Please update the Form Responses Raw sheet and the form to use the correct subject");
  // Confirm moving the prep sheet to a new folder
  const destination = quality === "Detailed" ? "Database" : "Reject";
  confirmationPrompt(`Would you like this prep sheet to be moved to the ${destination} Folder?`);

  // Get semester semester
  const semester = getSetting("Semester");
  // Move the prep sheet to the proper folder
  if (quality === "Detailed") {
    // Get the proper folder, creating it if needed
    const databaseFolder = getDriveFolderByUrl(getSetting("Prep Sheet Database Folder"));
    const subjectFolder = getChildFolder(databaseFolder, subject);
    const semesterFolder = getChildFolder(subjectFolder, semester);
    // Move the prep sheet to the folder
    prepSheet.moveTo(semesterFolder);

    // Look for any external links to files saved in Google Drive
    // If any are found, copy the linked document, move them to the external files folder, and replace the links
    // Open the prep sheet
    if (prepSheet.getMimeType() === MimeType.GOOGLE_DOCS) {
      const doc = DocumentApp.openByUrl(prepSheetUrl);
      const body = doc.getBody();
      const bodyText = body.editAsText();

      // Make a list to store the external urls and their locations
      const externalUrls = [];
      // Move the cursor through the document text
      let i = 0;
      while (i < bodyText.getText().length) {
        // See if the current character links to a document
        let url = bodyText.getLinkUrl(i);
        // If so, run this block
        if (url) {
          let start = i;
          let end = i;
          // Move the cursor to find the end offset of the url
          while (i < bodyText.getText().length - 1) {
            let adjacentUrl = bodyText.getLinkUrl(i + 1);
            if (adjacentUrl === url) {
              end++;
              i++;
            }
            else {
              break;
            }
          }
          // Store the url, start, and end in a list
          externalUrls.push({"url": url, "start": start, "end": end});
        }
        // Move the cursor before the next loop
        i++;
      }

      // Create a map between old document urls and new urls
      const externalUrlMap = new Map();
      // Duplicate the files found links to Google Drive or Google Docs
      for (const externalUrl of externalUrls) {
        const url = externalUrl.url;
        // Check if the url matches the document pattern
        const docPattern = /^https:\/\/(docs|drive)\.google\.com/
        const matches = docPattern.test(url);
        const copyMade = externalUrlMap.has(url);
        // If there is a match, try to make a copy
        if (matches && !copyMade) {
          try {
            // Get the external file
            const externalFile = getDriveFileByUrl(url);
            // Get the proper folder, creating it if needed
            const externalFilesFolder = getDriveFolderByUrl(getSetting("External Files Folder"));
            const externalSubjectFolder = getChildFolder(externalFilesFolder, subject)
            const externalSemesterFolder = getChildFolder(externalSubjectFolder, semester);
            // Make a copy of the external file
            const copiedFile = externalFile.makeCopy(externalFile.getName(), externalSemesterFolder);
            // Map the old url to the new url
            externalUrlMap.set(url, copiedFile.getUrl())
          }
          catch {
            // Continue to the next url if an error occurs
            continue;
          }
        }
      }
      
      // Replace the old urls in the document with the new urls
      for (const externalUrl of externalUrls) {
        const url = externalUrl.url;
        if (externalUrlMap.has(url)) {
          const start = externalUrl.start;
          const end = externalUrl.end;
          const newUrl = externalUrlMap.get(url);
          bodyText.setLinkUrl(start, end, newUrl);
        }
      }
    }
  }
  else {
    // Get the proper folder, creating it if needed
    const rejectFolder = getDriveFolderByUrl(getSetting("Reject Folder"));
    const semesterFolder = getChildFolder(rejectFolder, semester);
    // Move the prep sheet to the folder
    prepSheet.moveTo(semesterFolder);
  }

  // Set the folder cell in the spreadsheet with the destination
  responseSheet.getRange(activeRowIdx, responseCols.folder + 1).setValue(destination);
}
