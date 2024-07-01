// SECTION: Key Spreadsheet Column Numbers

const studentCols = {
  courseName: 0,
  studentId: 1,
  studentName: 2,
  totalHours: 3,
  totalSessions: 4,
  grade: 5,
}

const courseCols = {
  dataAdded: 0,
  tutorName: 1,
  category: 2,
  courseName: 3,
  attendanceSheet: 4,
  attendanceSheetPage: 5,
  timecardSheet: 6,
  gradeSheet: 7,
  sessionHours: 10,
}

const attendanceCols = {
  studentName: 0,
  totalHours: 1,
  totalSessions: 2,
}

const timecardCols = {
  section: 2,
  duty: 3,
  hours: 5,
}

const gradeCols = {
  studentId: 1,
  studentName: 2,
  gradeLetter: 5,
}

// SECTION: Menus

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Scripts")
    .addItem("Get Course Data", "getCourseData")
    .addToUi();
}

// SECTION: Main

function getCourseData() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  // Open the course sheet, making sure it is the active sheet
  const courseSheet = SS.getActiveSheet();
  if (courseSheet.getSheetName() != "Courses") {
    throw new Error("The 'Courses' sheet must be active");
  }
  // Get the active row, and check that it is not the headers
  const activeRowIdx = courseSheet.getCurrentCell().getRow();
  if (activeRowIdx === 1) {
    throw new Error("Select a row besides the header")
  }
  const activeRow = courseSheet.getRange(activeRowIdx, 1, 1, Object.keys(courseCols).length);
  const activeRowValues = activeRow.getValues()[0];

  // Get the total number of session hours offered
  const courseName = activeRowValues[courseCols.courseName];
  const crn = courseName.match(/\((\d+)\)/)[1];
  // Open the timecard sheet
  const timecardUrl = activeRowValues[courseCols.timecardSheet];
  const timecardSS = SpreadsheetApp.openByUrl(timecardUrl);
  const timecardSheet = timecardSS.getSheets()[0];
  const timecardRange = timecardSheet.getDataRange();
  const timecardValues = timecardRange.getValues();
  // Filter for entries with the section and that were a group session
  // Then add up all the hours
  const totalHours = timecardValues
    .filter(x => x[timecardCols.section].includes(crn))
    .filter(x => x[timecardCols.duty].toLowerCase().includes("group session"))
    .filter(x => !x[timecardCols.duty].toLowerCase().includes("no show"))
    .map(x => x[timecardCols.hours])
    .reduce((a, b) => a + b, 0);
  // Write the value to the spreadsheet
  courseSheet.getRange(activeRowIdx, courseCols.sessionHours + 1)
    .setValue(totalHours);

  // Open the grade spreadsheet
  const gradeUrl = activeRowValues[courseCols.gradeSheet];
  const gradeSS = SpreadsheetApp.openByUrl(gradeUrl);
  const gradeRows = gradeSS.getActiveSheet()
    .getDataRange()
    .offset(1, 0)
    .getValues();

  // Open the attendance sheet
  const attendanceUrl = activeRowValues[courseCols.attendanceSheet];
  const attendanceSS = SpreadsheetApp.openByUrl(attendanceUrl);
  const attendanceValues = attendanceSS.getSheets()
    [activeRowValues[courseCols.attendanceSheetPage]]
    .getDataRange()
    .offset(1, 0)
    .getValues();

  // Start looping through the rows in the grade sheet
  const data = [];
  for (const row of gradeRows) {

    // Extract the needed info from the grade row
    const studentId = row[gradeCols.studentId];
    const studentName = row[gradeCols.studentName];
    const gradeLetter = row[gradeCols.gradeLetter];

    // If the grade letter is 'EW' or '', skip the student
    if (["EW", ""].includes(gradeLetter)) {
      continue;
    }

    // Convert the grade letter to a number
    const gradeMap = new Map();
    gradeMap.set("A", 4);
    gradeMap.set("B", 3);
    gradeMap.set("C", 2);
    gradeMap.set("D", 1);
    gradeMap.set("F", 0);
    gradeMap.set("W", "W");
    const gradeNumber = gradeMap.get(gradeLetter);

    // Get the total hours & total number of session
    let totalHours = 0;
    let totalSessions = 0;
    const filteredRows = attendanceValues.filter(x => x[attendanceCols.studentName].includes(studentId));
    if (filteredRows.length !== 0) {
      totalHours = filteredRows[0][attendanceCols.totalHours];
      totalSessions = filteredRows[0][attendanceCols.totalSessions];
    }

    // Set the values for the student sheet
    const studentData = ["", "", "", "", "", ""];
    studentData[studentCols.courseName] = courseName;
    studentData[studentCols.studentId] = studentId;
    studentData[studentCols.studentName] = studentName;
    studentData[studentCols.grade] = gradeNumber;
    studentData[studentCols.totalHours] = totalHours;
    studentData[studentCols.totalSessions] = totalSessions;
    // Add the student data to the list
    data.push(studentData);
  }

  // Write the data to the spreadsheet
  const studentSheet = SS.getSheetByName("Students");
  // First delete any existing data
  // They come in chunks, so you just need to find the first row and the total count
  const courseValues = studentSheet.getDataRange()
    .getValues()
    .map(x => x[studentCols.courseName]);
  const startingRow = 1 + courseValues.indexOf(courseName);
  const courseCount = courseValues.filter(x => x === courseName).length;
  if (startingRow !== 0) {
    studentSheet.deleteRows(startingRow, courseCount);
  }
  // Then, write the new data at the bottom of the table
  const studentRange = studentSheet.getRange(studentSheet.getLastRow() + 1, 1, data.length, data[0].length);
  studentRange.setValues(data);
}
