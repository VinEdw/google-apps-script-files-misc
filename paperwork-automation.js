// SECTION: Key Spreadsheet Column Numbers

const tutorCols = {
  name: 1,
  email: 2,
  driveFolder: 3,
  paperworkDoc: 4,
  attendanceSheet: 5,
  timecardSheet: 6,
}
const professorCols = {
  name: 0,
  email: 1,
}
const courseCols = {
  tutorType: 0,
  subject: 1,
  groupSessionCRN: 2,
  courseName: 3,
  courseCRN: 4,
  days: 5,
  times: 6,
  location: 7,
  professor: 8,
  tutor: 9,
  lectureHours: 10,
  sessionHours: 11,
  prepHours: 12,
  observationHours: 13,
  trainingHours: 14,
  totalHours: 15,
  assignmentLetter: 16,
  attendanceSheet: 17,
}

// SECTION: Menus

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Scripts")
    .addItem("Create Tutor Paperwork", "paperworkPrompt")
    .addSeparator()
    .addItem("Email Tutor Paperwork Doc", "emailPaperworkPrompt")
    .addItem("Email Professor & Tutor Assignment Letter", "emailAssignmentPrompt")
    .addToUi();
}

// SECTION: User Prompts

function tutorPrompt() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt("Enter a tutor name:", ui.ButtonSet.OK_CANCEL)
  const button = result.getSelectedButton();
  const name = result.getResponseText().trim();
  if (button === ui.Button.OK) {
    const tutor = getTutor(name);
    return tutor;
  }
}

function paperworkPrompt() {
  const tutor = tutorPrompt();
  if (tutor) {
    createPaperwork(tutor);
  }
}

function emailPaperworkPrompt() {
  const tutor = tutorPrompt();
  if (tutor) {
    emailPaperwork(tutor);
  }
}

function emailAssignmentPrompt() {
  const tutor = tutorPrompt();
  if (tutor) {
    emailAssignment(tutor);
  }
}

// SECTION: Drive Helper Functions

/**
 * @param {string} url
 * @returns {DriveApp.Folder}
 */
function getFolderByUrl(url) {
  const id = url.match(/[-\w]{25,}/);
  return DriveApp.getFolderById(id);
}

/**
 * @param {DriveApp.File} file
 * @returns {DriveApp.Folder}
 */
function getParentFolder(file) {
  const parents = file.getParents();
  let folder = parents.next();
  return folder;
}

/**
 * Return the folder with the given name inside the parent folder
 * If create is true, create a folder with that name if it does not exist
 * @param {DriveApp.Folder} parent
 * @param {string} name
 * @param {boolean} create
 * @returns {DriveApp.Folder}
 */
function getChildFolder(parent, name, create = false) {
  const children = parent.getFolders();
  while (children.hasNext()) {
    let folder = children.next();
    if (folder.getName() === name) {
      return folder;
    }
  }
  if (create) {
    let newFolder = parent.createFolder(name);
    return newFolder;
  }
}

/**
 * @param {DriveApp.Folder} folder
 * @param {RegExp} regex
 * @returns {DriveApp.File}
 */
function getChildFileRegex(folder, regex) {
  const children = folder.getFiles();
  while (children.hasNext()) {
    let file = children.next();
    let name = file.getName();
    if (regex.test(name)) {
      return file;
    }
  }
}

/**
 * @param {DriveApp.File} file
 * @returns {boolean}
 */
function allowAnyoneViewFile(file) {
  try {
    // Set the sharing permission (anyone with link can view)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return true;
  }
  catch {
    return false;
  }
}

// SECTION: Form Helper Functions

/**
 * @param {DriveApp.File} formFile
 * @param {FormApp.Form} form
 * @param {DriveApp.Form} folder
 * @returns {string}
 */
function createLinkedSheet(formFile, form, folder) {
  // Create a new spreadsheet
  const linkedSS = SpreadsheetApp.create(formFile.getName() + " (Responses)");
  // Move the spreadsheet to the desired folder and make the form connection
  const file = DriveApp.getFileById(linkedSS.getId())
  file.moveTo(folder);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, linkedSS.getId());
  allowAnyoneViewFile(file);
  // Delete the empty default sheet
  let emptySheet = linkedSS.getSheetByName("Sheet1");
  linkedSS.deleteSheet(emptySheet);
  // Return the spreadsheet url
  return linkedSS.getUrl();
}

/**
 * @param {FormApp.ListItem} listItem
 * @returns {FormApp.ListItem}
 */
function createWeekDropdown(listItem) {
  // Some helper functions
  const pad2Digits = num => ("0" + num.toString()).slice(-2);
  const addDays = (date, days) => {
    let newDate = new Date(date.valueOf());
    newDate.setDate(newDate.getDate() + days);
    return newDate;
  }
  const dateStr = date => `${pad2Digits(date.getMonth() + 1)}/${pad2Digits(date.getDate())}`;

  // Parse the instructions in the config sheet to determine the starting date for week 0 and the number of semester weeks
  let mon = getSetting("Week0Monday");
  let totalWeeks = getSetting("TotalWeeks");

  // Create an array to store the new values
  const newValues = [];
  for (let i = 0; i <= totalWeeks; i++) {
    // Create variables for the Sunday and Friday of the week
    let sun = addDays(mon, 6);
    let fri = addDays(mon, 4);
    // Create a string representation of the week option and add it to the list
    let str = `Week ${pad2Digits(i)}: ${dateStr(mon)} - ${dateStr(sun)} (Submission due Friday, ${dateStr(fri)})`;
    newValues.push(str);
    // Add 7 days to the current Monday to get the next week
    mon = addDays(mon, 7);
  }

  // Set the choices on the listItem
  listItem.setChoiceValues(newValues);
  return listItem;
}

// SECTION: Doc Helper Functions

/**
 * @param {DocumentApp.Body} body
 * @param {string} pattern
 * @param {string} msg
 */
function replaceDocText(body, pattern, msg) {
  // Edit the document as text
  const text = body.editAsText();
  // Replace the pattern with the desired message
  text.replaceText(pattern, msg);
}

/**
 * @param {DocumentApp.Body} body
 * @param {string} pattern
 * @param {string} msg
 * @param {string} url
 */
function injectLink(body, pattern, msg, url) {
  // Edit the document as text
  const text = body.editAsText();
  // Find where the pattern starts
  let startIndex = text.getText().search(pattern);
  // Replace the pattern with the desired message
  text.replaceText(pattern, msg);
  // Turn the message into a link with the desired url
  text.setLinkUrl(startIndex, startIndex + msg.length - 1, url);
}

// SECTION: Spreadsheet Data Parsing

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

class Person {
  /**
   * @param {string} name
   * @param {string} email
   */
  constructor(name, email) {
    this.name = name;
    this.email = email;
  }
}

class Tutor extends Person {
  /**
   * @param {string} name
   * @param {string} email
   * @param {Array.<Course>} courses
   */
  constructor(name, email, courses) {
    super(name, email);
    this.courses = courses;
  }

  /**
   * @param {string} tutorType
   * @returns {boolean}
   */
  isTutorType(tutorType) {
    return this.courses
      .map(x => x.tutorType)
      .some(x => x === tutorType);
  }

  /**
   * @returns {string}
   */
  getSubject() {
    const stemCount = this.courses.filter(x => x.subject === "STEM").length;
    const humCount = this.courses.filter(x => x.subject === "HUM").length;
    return stemCount >= humCount ? "STEM" : "HUM";
  }

  /**
   * @returns {Arr.<string>}
   */
  getProfessorNames() {
    const professorNames = [];
    for (const course of this.courses) {
      let professorName = course.professor.name;
      if (professorNames.includes(professorName)) {
        continue;
      }
      professorNames.push(professorName);
    }
    return professorNames;
  }
}

class Professor extends Person {
  /**
   * @param {string} name
   * @param {string} email
   */
  constructor(name, email) {
    super(name, email);
  }
}

class Course {
  /**
   * @param {Object} param
   * @param {string} param.tutorType
   * @param {string} param.subject
   * @param {string} param.name
   * @param {number} param.courseCRN
   * @param {number} param.groupSessionCRN
   * @param {string} param.days
   * @param {string} param.times
   * @param {string} param.location
   * @param {Professor} param.professor
   * @param {string} param.lectureHours
   * @param {string} param.sessionHours
   * @param {string} param.prepHours
   * @param {string} param.observationHours
   * @param {string} param.trainingHours
   * @param {string} param.totalHours
   * @param {SpreadsheetApp.Range} param.assignmentLetterCell
   * @param {SpreadsheetApp.Range} param.attendanceSheetCell
   */
  constructor({tutorType, subject, name, courseCRN, groupSessionCRN, days, times, location, professor, lectureHours, sessionHours, prepHours, observationHours, trainingHours, totalHours, assignmentLetterCell, attendanceSheetCell} = {}) {
    this.tutorType = tutorType;
    this.subject = subject;
    this.name = name;
    this.courseCRN = courseCRN;
    this.groupSessionCRN = groupSessionCRN;
    this.days = days;
    this.times = times;
    this.location = location;
    this.professor = professor;
    this.lectureHours = lectureHours;
    this.sessionHours = sessionHours;
    this.prepHours = prepHours;
    this.observationHours = observationHours;
    this.trainingHours = trainingHours;
    this.totalHours = totalHours;
    this.assignmentLetterCell = assignmentLetterCell;
    this.attendanceSheetCell = attendanceSheetCell;
  }
}

function getProfessor(name) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();

  const professorRow = SS.getSheetByName("Professors")
    .getRange(1, 1)
    .getDataRegion()
    .getValues()
    .filter(x => x[professorCols.name] === name);
  let email = professorRow[0][professorCols.email];
  let professor = new Professor(name, email);
  return professor;
}

/**
 * @param {string} name
 * @returns {Tutor}
 */
function getTutor(name) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();

  // Get the spreadsheet row of the tutor with the input name
  const tutorRows = SS.getSheetByName("Tutors")
    .getRange(1, 1)
    .getDataRegion()
    .getValues()
    .filter(x => x[tutorCols.name] === name);
  if (tutorRows.length === 0) {
    throw new Error("Tutor name not found")
  }
  else if (tutorRows.length > 1) {
    throw new Error("Tutor name listed multiple times in 'Tutors' sheet")
  }
  const tutorRow = tutorRows[0]
  let email = tutorRow[tutorCols.email];

  // Get the courses that this tutor is assigned to
  const courseSheet = SS.getSheetByName("Courses");
  const courses = courseSheet.getRange(1, 1)
    .getDataRegion()
    .getValues()
    .map((x, idx) => [x, idx])
    .filter(x => x[0][courseCols.tutor] === name)
    .map(x => {
      return new Course({
        tutorType: x[0][courseCols.tutorType],
        subject: x[0][courseCols.subject],
        name: x[0][courseCols.courseName],
        courseCRN: x[0][courseCols.courseCRN],
        groupSessionCRN: x[0][courseCols.groupSessionCRN],
        days: x[0][courseCols.days],
        times: x[0][courseCols.times],
        location: x[0][courseCols.location],
        professor: getProfessor(x[0][courseCols.professor]),
        lectureHours: x[0][courseCols.lectureHours],
        sessionHours: x[0][courseCols.sessionHours],
        prepHours: x[0][courseCols.prepHours],
        observationHours: x[0][courseCols.observationHours],
        trainingHours: x[0][courseCols.trainingHours],
        totalHours: x[0][courseCols.totalHours],
        assignmentLetterCell: courseSheet.getRange(x[1] + 1, courseCols.assignmentLetter + 1),
        attendanceSheetCell: courseSheet.getRange(x[1] + 1, courseCols.attendanceSheet + 1),
      })
    })

  // Return a tutor object
  let tutor = new Tutor(name, email, courses);
  return tutor;
}

// SECTION: Paperwork Creation

/**
 * @param {Tutor} tutor
 */
function createPaperwork(tutor) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const fileSS = DriveApp.getFileById(SS.getId());
  const parentFolder = getParentFolder(fileSS);

  // Create a folder for the tutor
  const paperworkFolder = getChildFolder(parentFolder, "Paperwork Submissions", true);
  const subjectFolder = getChildFolder(paperworkFolder, tutor.getSubject(), true);
  const tutorFolderName = `${tutor.name}`;
  const tutorFolder = subjectFolder.createFolder(tutorFolderName);

  // Start duplicating and tailoring the files from the templates folder
  const templateFolder = getFolderByUrl(getSetting("TemplateFolder"));
  if (!templateFolder) {
    throw new Error("'Templates' folder not found")
  }

  // Create the time record form and its linked spreadsheet
  const timeRecordLinks = createTimeRecord(tutor, tutorFolder, templateFolder);

  // Create the attendance form and its linked spreadsheet
  const attendanceFormLinks = createAttendanceForm(tutor, tutorFolder, templateFolder);

  // Create the availability survey
  const availabilitySurveyLinks = createAvailabilitySurvey(tutor, tutorFolder, templateFolder);

  // Create the assignment letters
  const assignmentLetterLinks = createAssignmentLetters(tutor, tutorFolder, templateFolder, availabilitySurveyLinks, attendanceFormLinks);

  // Create the paperwork submission links doc
  const paperworkLink = createPaperworkDoc(tutor, tutorFolder, templateFolder, timeRecordLinks, attendanceFormLinks, availabilitySurveyLinks, assignmentLetterLinks);

  // Write the links to the spreadsheet
  // If there are multiple attendance sheets, use the folder; otherwise use the link
  let attendanceSheetUrl;
  if (attendanceFormLinks.sheets.length > 1) {
    attendanceSheetUrl = attendanceFormLinks.folder;
  }
  else {
    attendanceSheetUrl = attendanceFormLinks.sheets[0];
  }
  updateTutorLinks(tutor, tutorFolder.getUrl(), paperworkLink, attendanceSheetUrl, timeRecordLinks.sheet);
}

/**
 * @param {Tutor} tutor
 * @param {string} driveFolder
 * @param {string} paperworkDoc
 * @param {string} attendanceSheet
 * @param {string} timecardSheet
 */
function updateTutorLinks(tutor, driveFolder, paperworkDoc, attendanceSheet, timecardSheet) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const tutorSheet = SS.getSheetByName("Tutors");
  const table = tutorSheet.getRange(1, 1).getDataRegion();
  const tableValues = table.getValues();
  let targetRowIdx;
  for (const [i, row] of tableValues.entries()) {
    const name = row[tutorCols.name];
    if (tutor.name === name) {
      targetRowIdx = i;
      break;
    }
  }
  table.offset(targetRowIdx, 3, 1, 4)
    .setValues([[driveFolder, paperworkDoc, attendanceSheet, timecardSheet]]);
}

/**
 * @param {Tutor} tutor
 * @param {DriveApp.Folder} tutorFolder
 * @param {DriveApp.Folder} templateFolder
 */
function createTimeRecord(tutor, tutorFolder, templateFolder) {
  // Store any links created in this object
  const links = {};
  // Find the time record template and duplicate it
  const template = getChildFileRegex(templateFolder, /Time Record/);
  const file = template.makeCopy(
    template.getName()
      .replaceAll("{semester}", getSetting("Semester"))
      .replaceAll("{tutorName}", tutor.name),
    tutorFolder
  );
  const form = FormApp.openByUrl(file.getUrl());
  links.form = form.getPublishedUrl();

  // Set the new title to use the tutor name
  let title = form.getTitle();
  let newTitle = title
    .replaceAll("{semester}", getSetting("Semester"))
    .replaceAll("{tutorName}", tutor.name);
  form.setTitle(newTitle);

  // Save the item of the 'Total number of hours' question
  let hoursQuestion;
  // Modify certain questions as needed
  const items = form.getItems();
  for (let item of items) {
    let title = item.getTitle();
    // Set the tutor's sections
    if (title === "Select section") {
      let sectionSelect = item.asCheckboxItem();
      let choices = sectionSelect.getChoices();
      for (const course of tutor.courses) {
        const courseStr = `${course.name} (${course.courseCRN}) ${course.professor.name} ${course.days} ${course.times}`;
        let choice = sectionSelect.createChoice(courseStr);
        choices.push(choice);
      }
      sectionSelect.setChoices(choices);
    }
    else if (title === "Duty performed") {
      let dutySelect = item.asMultipleChoiceItem();
      if (!tutor.isTutorType("SI")) {
        let choices = dutySelect.getChoices();
        const siDuties = ["Prep", "Observation", "Observation Debrief"];
        choices = choices.filter(x => !siDuties.includes(x.getValue()));
        dutySelect.setChoices(choices);
      }
    }
    else if (title === "Week") {
      let weekSelect = item.asListItem();
      createWeekDropdown(weekSelect);
    }
    else if (title === "Total number of hours") {
      hoursQuestion = item.asTextItem();
    }
  }

  // Move the hours question to index 4
  // The question will be moved back to its original spot after the sheet is made
  // This is done to adjust the column order in the linked spreadsheet
  let hoursQuestionIdx = hoursQuestion.getIndex();
  let temporaryIdx = 5;
  form.moveItem(hoursQuestionIdx, temporaryIdx);
  // Create a linked spreadsheet and save the url
  links.sheet = createLinkedSheet(file, form, tutorFolder)
  // Move the hours question back to its original position
  form.moveItem(temporaryIdx, hoursQuestionIdx);
  // Add week and month summary sheets & formulas to the spreadsheet
  const timeRecordSS = SpreadsheetApp.openByUrl(links.sheet);
  const dayFormula = `=QUERY('Form Responses 1'!A:J, "SELECT E, SUM(F) WHERE E IS NOT NULL GROUP BY E LABEL E 'Date', SUM(F) 'Hours'", 1)`;
  const weekFormula = `=QUERY('Form Responses 1'!A:J, "SELECT B, SUM(F) WHERE B IS NOT NULL GROUP BY B LABEL B 'Week', SUM(F) 'Hours'", 1)`;
  const monthFormula = `=QUERY('Form Responses 1'!A:J, "SELECT MONTH(E)+1, SUM(F) WHERE E IS NOT NULL GROUP BY MONTH(E)+1 LABEL MONTH(E)+1 'Month', SUM(F) 'Hours'", 1)`;
  timeRecordSS.insertSheet("Breakdown by Day")
    .getRange(1, 1)
    .setFormula(dayFormula);
  timeRecordSS.insertSheet("Breakdown by Week")
    .getRange(1, 1)
    .setFormula(weekFormula);
  timeRecordSS.insertSheet("Breakdown by Month")
    .getRange(1, 1)
    .setFormula(monthFormula);
  
  // Return the links
  return links;
}

/**
 * @param {Tutor} tutor
 * @param {DriveApp.Folder} tutorFolder
 * @param {DriveApp.Folder} templateFolder
 */
function createAttendanceForm(tutor, tutorFolder, templateFolder) {
  // Store any links created in this object
  const links = {"forms": [], "sheets": []};
  // Put the attendance sheets in a folder
  const attendanceFolder = getChildFolder(tutorFolder, "Attendance Sheets", true);
  links.folder = attendanceFolder.getUrl();
  // Create an attendance form and sheet for each professor the tutor works with
  for (let professorName of tutor.getProfessorNames()) {
    const courses = tutor.courses.filter(x => x.professor.name === professorName);
    // Find the attendance form template and duplicate it
    const template = getChildFileRegex(templateFolder, /Student Attendance/);
    const file = template.makeCopy(
      template.getName()
        .replaceAll("{semester}", getSetting("Semester"))
        .replaceAll("{tutorName}", tutor.name)
        .replaceAll("{professorName}", professorName)
        .replaceAll("{courseCRNs}", courses.map(x => x.courseCRN).join("/")),
      tutorFolder
    );
    const form = FormApp.openByUrl(file.getUrl());
    links.forms.push(form.getPublishedUrl());

    // Set the new title to use the tutor name
    let title = form.getTitle();
    let newTitle = title
      .replaceAll("{semester}", getSetting("Semester"))
      .replaceAll("{tutorName}", tutor.name)
      .replaceAll("{professorName}", professorName)
      .replaceAll("{courseCRNs}", courses.map(x => x.courseCRN).join("/"));
    form.setTitle(newTitle);

    // Set the week select question
    for (let item of form.getItems()) {
      let title = item.getTitle();
      if (title === "Week") {
        let weekSelect = item.asListItem();
        createWeekDropdown(weekSelect);
        break;
      }
    }

    // For each course, create a question
    // Save the index of the questions
    const courseIdxs = [];
    for (const course of courses) {
      const courseStr = `${course.name} (${course.courseCRN}) ${course.professor.name} ${course.days} ${course.times}`;
      const questionStr = courseStr + " - Select all students who attended the group session:";
      const item = form.addCheckboxItem();
      item.setTitle(questionStr);
      form.moveItem(item.getIndex(), form.getItems().length - 2)
      courseIdxs.push(item.getIndex());
    }

    // Create a linked spreadsheet and save the url
    let sheetUrl = createLinkedSheet(file, form, attendanceFolder);
    links.sheets.push(sheetUrl);
    // Add attendance summary sheets with formulas to the spreadsheet
    const attendanceSS = SpreadsheetApp.openByUrl(sheetUrl);
    const responseSheet = attendanceSS.getSheetByName("Form Responses 1");
    for (let i = 0; i < courses.length; i++) {
      const course = courses[i];
      // Add the attendance sheet url to the schedule sheet
      course.attendanceSheetCell.setValue(sheetUrl);
      // Use the position of the question in the form to form to infer the column
      const courseIdx = courseIdxs[i];
      const courseCol = courseIdx + 2;
      // Get the column letter(s) where student names are entered
      const courseColLetter = responseSheet
        .getRange(1, courseCol)
        .getA1Notation()
        .replace(/\d+/, "");
      const courseColA1 = `'Form Responses 1'!$${courseColLetter}$2:$${courseColLetter}`
      const sheetName = `${course.name} ${course.professor.name} (${course.courseCRN})`;
      const sheet = attendanceSS.insertSheet(sheetName);
      const formulaCount = 500;
      // Increase the width of the first column
      sheet.setColumnWidth(1, 300);
      // Set the column headers
      const columnHeaders = ["Student", "Total Hours", "Total Sessions", "Stayed ≥ 0.5 hr", "Stayed ≥ 1.0 hr", "Stayed ≥ 1.5 hr", "Stayed ≥ 2.0 hr"];
      sheet.getRange(1, 1, 1, columnHeaders.length)
        .setValues([columnHeaders]);
      // Set the student list formula
      const studentFormula = `=SORT(UNIQUE(FLATTEN(IFERROR(ARRAYFORMULA(SPLIT(INDIRECT("${courseColA1}"), ", ", FALSE))))), 1, TRUE)`;
      sheet.getRange(2, 1)
        .setFormula(studentFormula);
      // Set the total hour calculating formula, setting the number format to 2 decimal places
      const hourFormula = `=IF(ISBLANK(A2), "", ARRAYFORMULA(SUM(FILTER(24*TIMEVALUE(INDIRECT("'Form Responses 1'!$F$2:$F")-INDIRECT("'Form Responses 1'!$E$2:$E")), FIND(A2, INDIRECT("${courseColA1}"))))))`;
      sheet.getRange(2, 2)
        .setFormula(hourFormula)
        .setNumberFormat("0.00")
        .copyTo(sheet.getRange(2, 2, formulaCount));
      // Set the total session count formula
      const sessionFormula = `=IF(ISBLANK(A2), "", ARRAYFORMULA(SUM(COUNT(FILTER(INDIRECT("'Form Responses 1'!$A$2:$A"), FIND(A2, INDIRECT("${courseColA1}")))))))`;
      sheet.getRange(2, 3)
        .setFormula(sessionFormula)
        .copyTo(sheet.getRange(2, 3, formulaCount));
      // Set the formulas to count the sessions attended for longer than a certain threshold
      for (let j = 0; j < 4; j++) {
        const threshold = 0.5 * (j + 1);
        const col = 4 + j;
        const thresholdFormula = `=IF(ISBLANK(A2), "", ARRAYFORMULA(SUM(COUNT(FILTER(INDIRECT("'Form Responses 1'!$A$2:$A"), FIND(A2, INDIRECT("${courseColA1}")), ARRAYFORMULA(24*TIMEVALUE(INDIRECT("'Form Responses 1'!$F$2:$F")-INDIRECT("'Form Responses 1'!$E$2:$E"))) >= ${threshold})))))`;
        sheet.getRange(2, col)
          .setFormula(thresholdFormula)
          .copyTo(sheet.getRange(2, col, formulaCount));
      }
    }
  }
  
  // Return the links
  return links;
}

/**
 * @param {Tutor} tutor
 * @param {DriveApp.Folder} tutorFolder
 * @param {DriveApp.Folder} templateFolder
 */
function createAvailabilitySurvey(tutor, tutorFolder, templateFolder) {
  // Store any links created in this object
  const links = {};
  // Find the availability survey form template and duplicate it
  const template = getChildFileRegex(templateFolder, /Student Availability Survey/);
  const file = template.makeCopy(
    template.getName()
      .replaceAll("{semester}", getSetting("Semester"))
      .replaceAll("{tutorName}", tutor.name),
    tutorFolder
  );
  const form = FormApp.openByUrl(file.getUrl());
  links.viewForm = form.getPublishedUrl();
  links.editForm = form.getEditUrl();

  // Set the new title to use the tutor name
  let title = form.getTitle();
  let newTitle = title.replaceAll("{tutorName}", tutor.name);
  form.setTitle(newTitle);
  // Set the form description to use the tutor name and CRNs
  let description = form.getDescription();
  let newDescription = description
    .replaceAll("{semester}", getSetting("Semester"))
    .replaceAll("{tutorName}", tutor.name)
    .replaceAll("{CRNKey}", tutor.courses.map(x => `- ${x.groupSessionCRN} → ${x.name} (${x.courseCRN}) ${x.professor.name} ${x.days} ${x.times}`).join("\n"));
  form.setDescription(newDescription);

  // Set the section select question
  for (let item of form.getItems()) {
    let title = item.getTitle();
    if (title === "Section") {
      let sectionSelect = item.asMultipleChoiceItem();
      let sections = tutor.courses
        .map(x => `${x.name} (${x.courseCRN}) ${x.professor.name} ${x.days} ${x.times}`);
      sectionSelect.setChoiceValues(sections);
      break;
    }
  }

  // Add the tutor as an editor for the form
  // If debug mode is not enabled
  if (!getSetting("Debug")) {
    file.addEditor(tutor.email);
  }

  // Return the links
  return links;
}

/**
 * @param {Tutor} tutor
 * @param {DriveApp.Folder} tutorFolder
 * @param {DriveApp.Folder} templateFolder
 */
function createPaperworkDoc(tutor, tutorFolder, templateFolder, timeRecordLinks, attendanceFormLinks, availabilitySurveyLinks, assignmentLetterLinks) {
  // Find the paperwork doc template and duplicate it
  const template = getChildFileRegex(templateFolder, /Paperwork Submission Links/);
  const file = template.makeCopy(
    template.getName()
      .replaceAll("{semester}", getSetting("Semester"))
      .replaceAll("{tutorName}", tutor.name),
    tutorFolder
  );
  const doc = DocumentApp.openByUrl(file.getUrl());

  // Get the document body
  const body = doc.getBody();

  // Define the url replacements to make
  injectLink(body, "{timecardForm}", "Timecard Form", timeRecordLinks.form);
  injectLink(body, "{timecardSheet}", "Timecard Sheet", timeRecordLinks.sheet);
  injectLink(body, "{studentAvailabilityForm}", "Student Availability Form", availabilitySurveyLinks.editForm);
  injectLink(body, "{roomRequestForm}", "Room Request Form", getSetting("RoomRequestForm"));
  injectLink(body, "{roomRequestSheet}", "Room Request Sheet", getSetting("RoomRequestSheet"));
  injectLink(body, "{endOfSemesterSurvey}", "End of Semester Survey", getSetting("EndOfSemesterSurvey"));
  replaceDocText(body, "{lastWeeks}", `${getSetting("TotalWeeks") - 2}-${getSetting("TotalWeeks") - 1}`);

  // If the tutor is an SI, inject the SI resource doc
  // Otherwise, delete the row in the document
  if (tutor.isTutorType("SI")) {
    injectLink(body, "{SIResourceDoc}", "SI Resources and Links", getSetting("SIResourceDoc"));
  }
  else {
    body.findText("{SIResourceDoc}")
      .getElement()
      .getParent()
      .getParent()
      .getParent()
      .removeFromParent();
  }

  // If there is only one attendance form+sheet, then simply make one replacement
  // Otherwise, make a link for each professor
  if (attendanceFormLinks.forms.length === 1) {
    injectLink(body, "{attendanceForm}", "Attendance Form", attendanceFormLinks.forms[0]);
    injectLink(body, "{attendanceSheet}", "Attendance Sheet", attendanceFormLinks.sheets[0]);
  }
  else {
    const professorNames = tutor.getProfessorNames();
    replaceDocText(body, "{attendanceForm}", `Attendance Forms:\n${professorNames.map(x => `{AF ${x}}`).join("\n")}`);
    replaceDocText(body, "{attendanceSheet}", `Attendance Sheets:\n${professorNames.map(x => `{AS ${x}}`).join("\n")}`);
    for (let i = 0; i < attendanceFormLinks.forms.length; i++) {
      let formUrl = attendanceFormLinks.forms[i];
      let sheetUrl = attendanceFormLinks.sheets[i];
      let professorName = professorNames[i];
      let formPattern = `{AF ${professorName}}`;
      let sheetPattern = `{AS ${professorName}}`;
      injectLink(body, formPattern, professorName, formUrl);
      injectLink(body, sheetPattern, professorName, sheetUrl);
    }
  }

  // If there is only one assignment letter, then simply make one replacement
  // Otherwise, make a link for each professor
  if (assignmentLetterLinks.length === 1) {
    injectLink(body, "{assignmentLetter}", "Assignment Letter", assignmentLetterLinks[0]);
  }
  else {
    const professorNames = tutor.getProfessorNames();
    replaceDocText(body, "{assignmentLetter}", `Assignment Letters:\n${professorNames.map(x => `{AL ${x}}`).join("\n")}`);
    for (let i = 0; i < assignmentLetterLinks.length; i++) {
      let url = assignmentLetterLinks[i];
      let professorName = professorNames[i];
      let pattern = `{AL ${professorName}}`;
      injectLink(body, pattern, professorName, url);
    }
  }

  allowAnyoneViewFile(file);

  return file.getUrl();
}

/**
 * @param {Tutor} tutor
 * @param {DriveApp.Folder} tutorFolder
 * @param {DriveApp.Folder} templateFolder
 */
function createAssignmentLetters(tutor, tutorFolder, templateFolder, availabilitySurveyLinks, attendanceFormLinks) {
  // Make a list of each assignment letter url
  const links = [];
  // Create an assignment letter for each professor
  for (let [i, professorName] of tutor.getProfessorNames().entries()) {
    const courses = tutor.courses.filter(x => x.professor.name === professorName);
    // Find the assignment letter doc template and duplicate it
    const assignmentLetterTemplate = getChildFileRegex(templateFolder, /Assignment Letter/);
    const assignmentLetter = assignmentLetterTemplate.makeCopy(
      assignmentLetterTemplate.getName()
        .replaceAll("{semester}", getSetting("Semester"))
        .replaceAll("{tutorName}", tutor.name)
        .replaceAll("{professorName}", professorName)
        .replaceAll("{tutorType}", courses[0].tutorType),
      tutorFolder
    );
    const assignmentLetterDoc = DocumentApp.openByUrl(assignmentLetter.getUrl());

    // Get the document body
    const body = assignmentLetterDoc.getBody();

    // Make the text and url replacements
    replaceDocText(body, "{professorName}", professorName);
    replaceDocText(body, "{tutorType}", courses[0].tutorType);
    replaceDocText(body, "{tutorName}", tutor.name);
    replaceDocText(body, "{groupSessionCRN}", courses[0].groupSessionCRN);
    injectLink(body, "{availabilitySurvey}", "availability survey", availabilitySurveyLinks.viewForm);
    replaceDocText(body, "{lectureHours}", courses[0].lectureHours);
    replaceDocText(body, "{sessionHours}", courses[0].sessionHours);
    replaceDocText(body, "{prepHours}", courses[0].prepHours);
    replaceDocText(body, "{observationHours}", courses[0].observationHours);
    replaceDocText(body, "{trainingHours}", courses[0].trainingHours);
    replaceDocText(body, "{totalHours}", courses[0].totalHours);
    injectLink(body, "{attendanceSheet}", "attendance sheet", attendanceFormLinks.sheets[i]);

    // Add extra rows in the course table if needed, then make the text replacements
    let courseTable = body.getTables()[0];
    for (let i = 0; i < courses.length - 1; i++) {
      let templateRow = courseTable.getRow(0).copy();
      courseTable.appendTableRow(templateRow);
    }
    for (let i = 0; i < courses.length; i++) {
      let course = courses[i];
      let row = courseTable.getRow(i);
      replaceDocText(row, "{courseName}", course.name);
      replaceDocText(row, "{courseCRN}", course.courseCRN);
      replaceDocText(row, "{days}", course.days);
      replaceDocText(row, "{times}", course.times);
      replaceDocText(row, "{location}", course.location);
      course.assignmentLetterCell.setValue(assignmentLetter.getUrl());
    }

    links.push(assignmentLetter.getUrl());
    
    allowAnyoneViewFile(assignmentLetter);
  }

  return links;
}

// SECTION: Email Sending Functions

/**
 * @param {Tutor} tutor
 */
function emailPaperwork(tutor) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const tutorSheet = SS.getSheetByName("Tutors");
  // Get the paperwork doc url
  const tutorRow = tutorSheet.getRange(1, 1)
    .getDataRegion()
    .getValues()
    .filter(x => x[tutorCols.name] === tutor.name)[0];
  const paperworkDoc = tutorRow[tutorCols.paperworkDoc];
  // Set the subject
  const subject = `${getSetting("Semester")} Paperwork Submission Links - ${tutor.name}`;
  // Set the body
  const templateFolder = getFolderByUrl(getSetting("TemplateFolder"));
  const emailTemplate = getChildFileRegex(templateFolder, /tutor-paperwork-email/)
    .getBlob()
    .getDataAsString();
  const bodyHtml = emailTemplate
    .replaceAll("{tutorName}", tutor.name)
    .replaceAll("{paperworkDoc}", paperworkDoc);
  // Create the email draft
  const draft = GmailApp.createDraft(tutor.email,
    subject,
    "",
    {cc: getSetting("StaffEmails"), htmlBody: bodyHtml});

  // Send the email draft if debug mode is not enabled
  if (!getSetting("Debug")) {
    draft.send();
  }
}

/**
 * @param {Tutor} tutor
 */
function emailAssignment(tutor) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const courseSheet = SS.getSheetByName("Courses");
  for (const profName of tutor.getProfessorNames()) {
    const professor = getProfessor(profName);
    // Get the assignment letter url
    const courseRow = courseSheet.getRange(1, 1)
      .getDataRegion()
      .getValues()
      .filter(x => x[courseCols.tutor] === tutor.name)
      .filter(x => x[courseCols.professor] === profName)[0];
    const assignmentLetter = courseRow[courseCols.assignmentLetter];
    // Set the subject
    const subject = `${getSetting("Semester")} ASAC Group Tutor Assignment Letter - ${profName}`;
    // Set the body
    const templateFolder = getFolderByUrl(getSetting("TemplateFolder"));
    const emailTemplate = getChildFileRegex(templateFolder, /assignment-letter-email/)
      .getBlob()
      .getDataAsString();
    const bodyHtml = emailTemplate
      .replaceAll("{professorName}", profName)
      .replaceAll("{tutorName}", tutor.name)
      .replaceAll("{assignmentLetter}", assignmentLetter);
    const draft = GmailApp.createDraft(professor.email + "," + tutor.email,
      subject,
      "",
      {cc: getSetting("StaffEmails"), htmlBody: bodyHtml});
    // Send the email draft if debug mode is not enabled
    if (!getSetting("Debug")) {
      draft.send();
    }
  }
}
