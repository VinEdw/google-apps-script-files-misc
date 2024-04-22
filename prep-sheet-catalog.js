function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Scripts")
    .addItem("Update Catalog", "createCatalog")
    .addToUi();
}

/**
 * @param {string} url
 * @returns {DriveApp.Folder}
 */
function getFolderByUrl(url) {
  const id = url.match(/[-\w]{25,}/);
  return DriveApp.getFolderById(id);
}

/**
 * Function to get values from the config sheet
 * @param {string} key
 * @returns {string}
 */
function getSetting(key) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = SS.getSheetByName("Config");
  const data = settingSheet.getRange(1, 1).getDataRegion().getValues().slice(1);
  let value = data.filter(x => x[0] === key)[0][1];
  return value;
}

function createCatalog() {
  // Get the database folder
  const databaseUrl = getSetting("database");
  const databaseFolder = getFolderByUrl(databaseUrl);
  // Create a variable to store all the desired file data
  const catalogData = [];
  // Loop through each subject folder in the database (except the rejects and external files)
  const subjectFolders = databaseFolder.getFolders();
  while (subjectFolders.hasNext()) {
    let subjectFolder = subjectFolders.next();
    let subject = subjectFolder.getName();
    if (["Z-External Files", "Z-Rejects"].includes(subject)) {
      continue;
    }
    const semesterFolders = subjectFolder.getFolders();
    while (semesterFolders.hasNext()) {
      let semesterFolder = semesterFolders.next();
      let semester = semesterFolder.getName();
      let files = semesterFolder.getFiles();
      while (files.hasNext()) {
        let file = files.next();
        let title = file.getName();
        let url = file.getUrl()
        catalogData.push([subject, semester, title, url]);
      }
    }
  }
  // Sort the data
  catalogData.sort((a, b) => {
    let cmpSubject = a[0].localeCompare(b[0]);
    let cmpSemester = a[1].localeCompare(b[1]);
    let cmpTitle = a[2].localeCompare(b[2]);
    return cmpSubject || cmpSemester || cmpTitle;
  });

  // Write the catalog data to the spreadsheet
  const SS = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = SS.getSheetByName("Sheet1");
  // Clear the old data
  sheet.getRange(1, 1)
    .getDataRegion()
    .offset(1, 0)
    .clearContent();
  // Write the new data
  sheet.getRange(2, 1)
    .offset(0, 0, catalogData.length, catalogData[0].length)
    .setValues(catalogData);
  // Update the file name with the current date
  let today = new Date(Date.now());
  let fname = `Prep Sheet Database Catalog (Last Updated: ${today.getMonth()+1}/${today.getDate()}/${today.getFullYear()})`;
  SS.rename(fname);
}
