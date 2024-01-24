function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Scheduling Tools')
  .addItem('Autocomplete schedule', 'createSchedule')
  .addSeparator()
  .addItem('ID to name converter', 'idToName')
  .addItem('Name to ID converter', 'nameToId')
  .addToUi();
}

function getBrothers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const brothersSheet = ss.getActiveSheet();
  const brothersRange = brothersSheet.getRange("A15").getDataRegion();
  const brothersArr = brothersRange.getValues().slice(2);
  const brothersList = [];

  const splitCountDay = (str) => {
    if (str === '') {
      str = '0b';
    }
    if (!str.match("^[0-9]+[stb]$")) {
      return undefined;
    }
    let num = parseInt(str.slice(0, -1), 10);
    let letter = str.slice(-1);
    let obj = {
      goal: num,
      tuesday: letter === 't' || letter === 'b',
      saturday: letter === 's' || letter === 'b',
    }
    return obj;
  }

  class Brother {
    constructor(id, lastName, firstName, name, shortName, host, sound, stage, lMicrophone, rMicrophone, zAttendant, fAttendant, sAttendant, totalGoal) {
      this.id = id;
      this.lastName = lastName;
      this.firstName = firstName;
      this.name = name;
      this.shortName = shortName;
      
      this.host = splitCountDay(host);
      this.sound = splitCountDay(sound);
      this.stage = splitCountDay(stage);
      this.lMicrophone = splitCountDay(lMicrophone);
      this.rMicrophone = splitCountDay(rMicrophone);
      this.zAttendant = splitCountDay(zAttendant);
      this.fAttendant = splitCountDay(fAttendant);
      this.sAttendant = splitCountDay(sAttendant);

      this.totalGoal = totalGoal;
    }
  }

  for (let i = 0; i < brothersArr.length; i++) {
    const row = brothersArr[i];
    const bro = new Brother(...row);
    brothersList.push(bro);
  }

  return {
    list: brothersList,
    roles: ['host', 'sound', 'stage', 'lMicrophone', 'rMicrophone', 'zAttendant', 'fAttendant', 'fAttendant', 'sAttendant', 'sAttendant'],
    getBrotherByName(name) {
      for (let i = 0; i < this.list.length; i++) {
        let bro = this.list[i];
        let broName = bro.name;
        if (broName === name) {
          return bro;
        }
      }
      return undefined;
    },
    getBrotherByID(id) {
      for (let i = 0; i < this.list.length; i++) {
        let bro = this.list[i];
        let broId = bro.id;
        if (broId === id) {
          return bro;
        }
      }
      return undefined;
    },
  };
}

function getScheduleSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getActiveSheet();

  const brotherCollection = getBrothers();
  
  const dates = scheduleSheet.getRange('Dates').getValues();

  const parts = scheduleSheet.getRange('Parts').getValues().map(row => row.filter(val => val != '').map(name => {
    const bro = brotherCollection.getBrotherByName(name);
    return bro && bro.id;
  }));

  const scheduleRange = scheduleSheet.getRange('Schedule');
  const scheduleArr = scheduleRange.getValues();

  const previousMeeting = scheduleSheet.getRange('PreviousMeeting').getValues().flat().map(name => {
      const bro = brotherCollection.getBrotherByName(name);
      return bro && bro.id;
    });

  return {ss, scheduleSheet, scheduleRange, scheduleArr, parts, dates, previousMeeting, brotherCollection};
}

function createSchedule() {
  let {ss, scheduleSheet, scheduleRange, scheduleArr, parts, dates, previousMeeting, brotherCollection} = getScheduleSheetData();
  let c = 0;

  function copyArrStructure(arr) {
    const newArr = [];
    for (let i = 0; i < arr.length; i++) {
      let element = arr[i];
      if (Array.isArray(element)) {
        newArr.push(copyArrStructure(element));
      }
      else {
        newArr.push(element);
      }
    }
    return newArr;
  }

  function getRow(y) {
    return scheduleArr[y]
  }
  function getColumn(x) {
    const newArr = [];
    for (let i = 0; i < scheduleArr.length; i++) {
      for (let j = 0; j < scheduleArr[i].length; j++) {
        const val = scheduleArr[i][j];
        if (j === x) {
          newArr.push(val);
        }
      }
    }
    return newArr;
  }
  function getCount(id, role, includePossibilities) {
    let count = 0;
    for (let j = 0; j < scheduleArr[0].length; j++) {
      let assignment = brotherCollection.roles[j];
      if (assignment === role) {
        for (const value of getColumn(j)) {
          if (value === id) {
            count++;
          }
          else if (includePossibilities && Array.isArray(value) && value.includes(id)) {
            count++;
          }
        }
      }
    }
    return count;
  }
  function checkAllGoals() {
    const roleSet = new Set(brotherCollection.roles);
    for (const bro of brotherCollection.list) {
      for (const role of roleSet) {
        const count = getCount(bro.id, role, true);
        const goal = bro[role].goal;
        if (count < goal) {
          // printSchedule()
          // console.log(scheduleArr)
          return false;
        }
      }
    }
    return true;
  }
  function getOccurrenceCount(id, arr) {
    let count = 0;
    for (let i = 0; i < arr.length; i++) {
      for (let j = 0; j < arr[i].length; j++) {
        let value = arr[i][j];
        if (value === id) {
          count++;
        }
      }
    }
    return count;
  }
  function printSchedule() {
    console.log(scheduleArr.map(element => element.map(item => Array.isArray(item) ? '' : item)));
  }

  function scheduleTo3D() {
    const idList = brotherCollection.list/*.sort((a, b) => {
      let partDiff = getOccurrenceCount(a.id, parts) - getOccurrenceCount(b.id, parts);
      if (partDiff) { return partDiff; }
      let goalDiff = b.totalGoal - a.totalGoal;
      if (goalDiff) { return goalDiff; }
      return a.id - b.id;
    })*/.map(bro => bro.id);
    for (let i = 0; i < scheduleArr.length; i++) {
      for (let j = 0; j < scheduleArr[i].length; j++) {
        const cell = scheduleArr[i][j];
        if (cell === '') {
          scheduleArr[i][j] = [...idList];
        }
      }
    }
  }

  function conflictChecker(y, x, n) {
    if (y === 0 && previousMeeting.includes(n)) {
      return false;
    }
    if (y !== 0 && scheduleArr[y-1].includes(n)) {
      return false;
    }
    if (y !== (scheduleArr.length-1) && scheduleArr[y+1].includes(n)) {
      return false;
    }
    if (scheduleArr[y].includes(n)) {
      return false;
    }

    if (parts[y].includes(n)) {
      return false;
    }

    const bro = brotherCollection.getBrotherByID(n);
    if (!bro) {
      return false;
    }
    const role = brotherCollection.roles[x];

    if (getCount(n, role, false) >= bro[role].goal) {
      return false;
    }

    // if (role === 'fAttendant' || role === 'sAttendant') {
    //   if (y !== 0 && y !== 1 && scheduleArr[y-2].includes(n)) {
    //     return false;
    //   }
    //   if (y !== (scheduleArr.length-1) && y !== (scheduleArr.length-2) && scheduleArr[y+2].includes(n)) {
    //     return false;
    //   }
    // }

    const day = dates[y][0];
    if (day === 'Tue') {
      if (!bro[role].tuesday) {
        return false;
      }
    }
    if (day === 'Sat') {
      if (!bro[role].saturday) {
        return false;
      }
    }

    return true;
  }

  function checkAllConflicts() {
    for (let i = 0; i < scheduleArr.length; i++) {
      for (let j = 0; j < scheduleArr[i].length; j++) {
        const cell = scheduleArr[i][j];
        if (typeof cell !== 'number') {
          continue;
        }
        scheduleArr[i][j] = '';
        const result = conflictChecker(i, j, cell);
        scheduleArr[i][j] = cell;
        if (!result) {
          console.log('Row:', i, 'Column:', j);
          console.log(cell, brotherCollection.getBrotherByID(cell).name);
          return false;
        }
      }
    }
    return true;
  }

  function checkFinished() {
    for (let i = 0; i < scheduleArr.length; i++) {
      for (let j = 0; j < scheduleArr[i].length; j++) {
        if (Array.isArray(scheduleArr[i][j])) {
          return false;
        }
      }
    }
    if (!checkAllConflicts()) {
      return false;
    }
    return true;
  }

  function extrapolateSchedule() {
    let updating = true;
    while (updating) {
      updating = false;
      for (let i = 0; i < scheduleArr.length; i++) {
        for (let j = 0; j < scheduleArr[i].length; j++) {
          if (Array.isArray(scheduleArr[i][j])) {
            const newArr = [];
            for (let k = 0; k < scheduleArr[i][j].length; k++) {
              let num = scheduleArr[i][j][k];
              if (conflictChecker(i, j, num)) {
                newArr.push(num);
              }
            }
            scheduleArr[i][j] = newArr;
          }
        }
      }
      for (let i = 0; i < scheduleArr.length; i++) {
        for (let j = 0; j < scheduleArr[i].length; j++) {
          if (Array.isArray(scheduleArr[i][j])) {
            if (scheduleArr[i][j].length === 0) {
              return false;
            }
            if (scheduleArr[i][j].length === 1) {
              let num = scheduleArr[i][j][0];
              if (!conflictChecker(i, j, num)) {
                return false;
              }
              scheduleArr[i][j] = num;
              updating = true;
            }
          }
        }
      }
    }
    return true;
  }

  function solve() {
    outerLoop:
    for (let j = 0; j < scheduleArr[0].length; j++) {
      for (let i = 0; i < scheduleArr.length; i++) {
        if (Array.isArray(scheduleArr[i][j])) {
          for (let num of scheduleArr[i][j]) {
            if (conflictChecker(i, j, num)) {
              const backup = copyArrStructure(scheduleArr);
              scheduleArr[i][j] = num;
              c++;

              if (c % 10000 === 0) {
                console.log(c);
                console.log(scheduleArr);
                printSchedule();
              }

              if (extrapolateSchedule() /*&& checkAllGoals()*/) {
                solve();
                if (checkFinished()) {
                  break outerLoop;
                }
              }
              else {
                // console.log(scheduleArr);
                // printSchedule();
              }
              scheduleArr = backup;
            }
          }
          return;
        }
      }
    }
  }

  console.log(scheduleSheet.getName());
  scheduleTo3D();
  if (!extrapolateSchedule() || !checkAllConflicts()) {
    console.log(scheduleArr);
    console.log('Initial state is unsolveable');
    return;
  }
  console.log(scheduleArr);
  printSchedule();
  solve();

  if (checkFinished()) {
    console.log('Finished', c);
    console.log(scheduleArr);
    scheduleRange.setValues(scheduleArr);
    console.log("Brothers who did not reach their schedule goal:")
    for (const brother of brotherCollection.list) {
      let id = brother.id;
      let name = brother.name;
      let goal = brother.totalGoal;
      let count = getOccurrenceCount(id, scheduleArr);
      if (count < goal) {
        console.log(`${id} ${name} ${count} < ${goal}`);
      }
    }
  }
  else {
    console.log('Could not be solved');
    console.log(scheduleArr);
  }
}

function idToName() {
  const {ss, scheduleSheet, scheduleRange, scheduleArr, parts, dates, previousMeeting, brotherCollection} = getScheduleSheetData();
  for (let i = 0; i < scheduleArr.length; i++) {
    for (let j = 0; j < scheduleArr[i].length; j++) {
      let cell = scheduleArr[i][j];
      let bro = brotherCollection.getBrotherByID(cell);
      if (bro) {
        scheduleArr[i][j] = bro.name;
      }
    }
  }
  scheduleRange.setValues(scheduleArr);
}

function nameToId() {
  const {ss, scheduleSheet, scheduleRange, scheduleArr, parts, dates, previousMeeting, brotherCollection} = getScheduleSheetData();
  for (let i = 0; i < scheduleArr.length; i++) {
    for (let j = 0; j < scheduleArr[i].length; j++) {
      let cell = scheduleArr[i][j];
      let bro = brotherCollection.getBrotherByName(cell);
      if (bro) {
        scheduleArr[i][j] = bro.id;
      }
    }
  }
  scheduleRange.setValues(scheduleArr);
}
