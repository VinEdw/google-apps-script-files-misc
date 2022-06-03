function onOpen(e) {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Create sheets by course', 'organizeResults')
    .addToUi();
}

function organizeResults() {
  // access the active spreadsheet
  const mySS = SpreadsheetApp.getActiveSpreadsheet();
  // make a list of the names of all the current sheets
  const sheetArr = mySS.getSheets();
  const sheetNamesArr = sheetArr.map(element => element.getName());

  // make a list of all the unique form responses (the class options) found in column D of the 'Form Responses 1' sheet
  const formResponses = mySS.getSheetByName('Form Responses 1');
  const classOptionsWithDuplicates = formResponses.getRange("D2").getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues().slice(1).sort();
  const classOptions = [];
  for (let i = 0; i < classOptionsWithDuplicates.length; i++) {
    let name = classOptionsWithDuplicates[i][0];
    if (!classOptions.includes(name)) {
      classOptions.push(name);
    }
  }

  // log the number of class options and the names of each one for reference 
  console.log(classOptions.length, classOptions);

  // for each course that does already have a sheet, add one with the appropriate query in cell A1
  let addCount = 0;
  for (let i = 0; i < classOptions.length; i++) {
    let name = classOptions[i] // the name of the course
    let eqName = name.replaceAll(/"/g, '""'); // that name with " replaced with "" for use in the cell equation/formula

    let splitName = name.split(/\s+/);
    let shortName = (splitName.slice(0, 5).join(' ') + ' - Tutor: ' + splitName.slice(-2).join(' ')).slice(0, 100); // that name shortened to only have key information

    // let shortName = name.slice(0, 100) // that name shortened to be 100 characters or less (so that it may be used as a sheet name)
    console.log(name);
    if (sheetNamesArr.includes(shortName)) {
      continue;
    }
    console.log('adding sheet')
    addCount++;
    const sheet = mySS.insertSheet(shortName, i + 1);
    sheet.getRange("A1").setFormula(`QUERY('Form Responses 1'!A:N, "SELECT * WHERE D = '${eqName}' ORDER BY F DESC LABEL B 'Name', C 'Email', D 'Course', F 'Online or In Person' ",1)`); // add the query formula
  }
  console.log(addCount + ' sheets added');
}