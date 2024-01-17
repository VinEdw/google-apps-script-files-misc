function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Scripts")
    .addItem("Create Tutor Paperwork", "paperworkPrompt")
    .addToUi();
}

function paperworkPrompt() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt("Enter a tutor name:", ui.ButtonSet.OK_CANCEL)
  const button = result.getSelectedButton();
  const name = result.getResponseText();
  if (button === ui.Button.OK) {
    const tutor = getTutor(name);
    createPaperwork(tutor);
  }
}

// Define some key column numbers in the spreadsheet
const tutorCols = {
  name: 0,
  email: 1,
  driveFolder: 2,
  paperworkDoc: 3,
  attendanceSheet: 4,
  timecardSheet: 5,
}
const professorCols = {
  name: 0,
  email: 1,
}
const courseCols = {
  tutorType: 0,
  subject: 1,
  courseName: 2,
  courseCRN: 3,
  groupSessionCRN: 4,
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
  assignmentLetter: 16
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
   * @param {string} tutorType
   * @param {string} subject
   * @param {string} name
   * @param {number} courseCRN
   * @param {number} groupSessionCRN
   * @param {string} days
   * @param {string} times
   * @param {string} location
   * @param {Professor} professor
   * @param {string} lectureHours
   * @param {string} sessionHours
   * @param {string} prepHours
   * @param {string} observationHours
   * @param {string} trainingHours
   * @param {string} totalHours
   * @param {SpreadsheetApp.Range} assignmentLetterCell
   */
  constructor(tutorType, subject, name, courseCRN, groupSessionCRN, days, times, location, professor, lectureHours, sessionHours, prepHours, observationHours, trainingHours, totalHours, assignmentLetterCell) {
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
  const tutorRow = SS.getSheetByName("Tutors")
    .getRange(1, 1)
    .getDataRegion()
    .getValues()
    .filter(x => x[tutorCols.name] === name);
  let email = tutorRow[0][tutorCols.email];

  // Get the courses that this tutor is assigned to
  const courseSheet = SS.getSheetByName("Courses");
  const courses = courseSheet.getRange(1, 1)
    .getDataRegion()
    .getValues()
    .map((x, idx) => [x, idx])
    .filter(x => x[0][courseCols.tutor] === name)
    .map(x => {
      return new Course(
        x[0][courseCols.tutorType],
        x[0][courseCols.subject],
        x[0][courseCols.courseName],
        x[0][courseCols.courseCRN],
        x[0][courseCols.groupSessionCRN],
        x[0][courseCols.days],
        x[0][courseCols.times],
        x[0][courseCols.location],
        getProfessor(x[0][courseCols.professor]),
        x[0][courseCols.lectureHours],
        x[0][courseCols.sessionHours],
        x[0][courseCols.prepHours],
        x[0][courseCols.observationHours],
        x[0][courseCols.trainingHours],
        x[0][courseCols.totalHours],
        courseSheet.getRange(x[1] + 1, courseCols.assignmentLetter + 1)
      )
    })

  // Return a tutor object
  let tutor = new Tutor(name, email, courses);
  return tutor;
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

/**
 * @param {Tutor} tutor
 */
function createPaperwork(tutor) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const fileSS = DriveApp.getFileById(SS.getId());
  const parentFolder = getParentFolder(fileSS);

  // Create a folder for the tutor
  const paperworkFolder = getChildFolder(parentFolder, "Paperwork Submissions");
  const subjectFolder = getChildFolder(paperworkFolder, tutor.getSubject());
  const tutorFolderName = `${tutor.name}`;
  const tutorFolder = subjectFolder.createFolder(tutorFolderName);

  // Start duplicating and tailoring the files from the templates folder
  const templateFolder = getChildFolder(parentFolder, "Templates");

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
  updateTutorLinks(tutor, tutorFolder.getUrl(), paperworkLink, attendanceFormLinks.sheet, timeRecordLinks.sheet);
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
  for (const row of tableValues) {
    const name = row[tutorCols.name];
    if (tutor.name === name) {
      row[tutorCols.driveFolder] = driveFolder;
      row[tutorCols.paperworkDoc] = paperworkDoc;
      row[tutorCols.attendanceSheet] = attendanceSheet;
      row[tutorCols.timecardSheet] = timecardSheet;
      break;
    }
  }
  table.setValues(tableValues);
}

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

  // Parse the instructions in the first option to determine the starting date for week 0 and the number of semester weeks
  const input = JSON.parse(listItem.getChoices()[0].getValue());

  // Create a variable for that monday
  let mon = new Date(input.week0Year, input.week0Month - 1, input.week0Day);
  // Create an array to store the new values
  const newValues = [];
  for (let i = 0; i <= input.totalWeeks; i++) {
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

/**
 * @param {Tutor} tutor
 * @param {DriveApp.Folder} tutorFolder
 * @param {DriveApp.Folder} templateFolder
 */
function createTimeRecord(tutor, tutorFolder, templateFolder) {
  // Store any links created in this object
  const links = {};
  // Find the time record template and duplicate it
  const timeRecordTemplate = getChildFileRegex(templateFolder, /Time Record/);
  const timeRecord = timeRecordTemplate.makeCopy(
    timeRecordTemplate.getName().replaceAll("{tutorName}", tutor.name),
    tutorFolder
  );
  const timeRecordForm = FormApp.openByUrl(timeRecord.getUrl());
  links.form = timeRecordForm.getPublishedUrl();

  // Set the new title to use the tutor name
  let title = timeRecordForm.getTitle();
  let newTitle = title.replaceAll("{tutorName}", tutor.name);
  timeRecordForm.setTitle(newTitle);

  // Modify certain questions as needed
  const items = timeRecordForm.getItems();
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
      let dutySelect = item.asCheckboxItem();
      if (!tutor.isTutorType("SI")) {
        let choices = dutySelect.getChoices();
        const siDuties = ["Prep", "Observation", "Observation Debrief"];
        choices = choices.filter(x => !siDuties.includes(x.getValue()));
        dutySelect.setChoices(choices);
      }
    }
    else if (title === "Week - Select the week") {
      let weekSelect = item.asListItem();
      createWeekDropdown(weekSelect);
    }
  }

  // Create a linked spreadsheet and save the url
  links.sheet = createLinkedSheet(timeRecord, timeRecordForm, tutorFolder)
  
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
  const links = {};
  // Find the attendance form template and duplicate it
  const attendanceTemplate = getChildFileRegex(templateFolder, /Student Attendance/);
  const attendance = attendanceTemplate.makeCopy(
    attendanceTemplate.getName()
      .replaceAll("{tutorName}", tutor.name)
      .replaceAll("{courseCRNs}", tutor.courses.map(x => x.courseCRN).join("/")),
    tutorFolder
  );
  const attendanceForm = FormApp.openByUrl(attendance.getUrl());
  links.form = attendanceForm.getPublishedUrl();

  // Set the new title to use the tutor name
  let title = attendanceForm.getTitle();
  let newTitle = title
    .replaceAll("{tutorName}", tutor.name)
    .replaceAll("{courseCRNs}", tutor.courses.map(x => x.courseCRN).join("/"));
  attendanceForm.setTitle(newTitle);

  // Set the week select question
  for (let item of attendanceForm.getItems()) {
    let title = item.getTitle();
    if (title === "Week - Select the week") {
      let weekSelect = item.asListItem();
      createWeekDropdown(weekSelect);
      break;
    }
  }

  // For each course, create a question
  for (const course of tutor.courses) {
    const courseStr = `${course.name} (${course.courseCRN}) ${course.professor.name} ${course.days} ${course.times}`;
    const questionStr = courseStr + " - Select all students who attended the group session:";
    const item = attendanceForm.addCheckboxItem();
    item.setTitle(questionStr);
    attendanceForm.moveItem(item.getIndex(), attendanceForm.getItems().length - 2)
  }

  // Create a linked spreadsheet and save the url
  links.sheet = createLinkedSheet(attendance, attendanceForm, tutorFolder);
  
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
  const surveyTemplate = getChildFileRegex(templateFolder, /Student Availability Survey/);
  const survey = surveyTemplate.makeCopy(
    surveyTemplate.getName().replaceAll("{tutorName}", tutor.name),
    tutorFolder
  );
  const surveyForm = FormApp.openByUrl(survey.getUrl());
  links.form = surveyForm.getPublishedUrl();

  // Set the new title to use the tutor name
  let title = surveyForm.getTitle();
  let newTitle = title.replaceAll("{tutorName}", tutor.name);
  surveyForm.setTitle(newTitle);
  // Set the form description to use the tutor name and CRNs
  let description = surveyForm.getDescription();
  let newDescription = description
    .replaceAll("{tutorName}", tutor.name)
    .replaceAll("{CRNKey}", tutor.courses.map(x => `- ${x.groupSessionCRN} → ${x.name} (${x.courseCRN}) ${x.professor.name} ${x.days} ${x.times}`).join("\n"));
  surveyForm.setDescription(newDescription);

  // Add the tutor as an editor for the form
  survey.addEditor(tutor.email);

  // Return the links
  return links;
}

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

/**
 * @param {Tutor} tutor
 * @param {DriveApp.Folder} tutorFolder
 * @param {DriveApp.Folder} templateFolder
 */
function createPaperworkDoc(tutor, tutorFolder, templateFolder, timeRecordLinks, attendanceFormLinks, availabilitySurveyLinks, assignmentLetterLinks) {
  // Find the paperwork doc template and duplicate it
  const paperworkTemplate = getChildFileRegex(templateFolder, /Paperwork Submission Links/);
  const paperwork = paperworkTemplate.makeCopy(
    paperworkTemplate.getName().replaceAll("{tutorName}", tutor.name),
    tutorFolder
  );
  const paperworkDoc = DocumentApp.openByUrl(paperwork.getUrl());

  // Get the document body
  const body = paperworkDoc.getBody();

  // Define the url replacements to make
  injectLink(body, "{attendanceForm}", "Attendance Form", attendanceFormLinks.form);
  injectLink(body, "{attendanceSheet}", "Attendance Sheet", attendanceFormLinks.sheet);
  injectLink(body, "{timecardForm}", "Timecard Form", timeRecordLinks.form);
  injectLink(body, "{timecardSheet}", "Timecard Sheet", timeRecordLinks.sheet);
  injectLink(body, "{studentAvailabilityForm}", "Student Availability Form", availabilitySurveyLinks.form);

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

  allowAnyoneViewFile(paperwork);

  return paperwork.getUrl();
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
  for (let professorName of tutor.getProfessorNames()) {
    const courses = tutor.courses.filter(x => x.professor.name === professorName);
    // Find the assignment letter doc template and duplicate it
    const assignmentLetterTemplate = getChildFileRegex(templateFolder, /Assignment Letter/);
    const assignmentLetter = assignmentLetterTemplate.makeCopy(
      assignmentLetterTemplate.getName()
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
    injectLink(body, "{availabilitySurvey}", "availability survey", availabilitySurveyLinks.form);
    replaceDocText(body, "{lectureHours}", courses[0].lectureHours);
    replaceDocText(body, "{sessionHours}", courses[0].sessionHours);
    replaceDocText(body, "{prepHours}", courses[0].prepHours);
    replaceDocText(body, "{observationHours}", courses[0].observationHours);
    replaceDocText(body, "{trainingHours}", courses[0].trainingHours);
    replaceDocText(body, "{totalHours}", courses[0].totalHours);
    injectLink(body, "{attendanceSheet}", "attendance sheet", attendanceFormLinks.sheet);

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
