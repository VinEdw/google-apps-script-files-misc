// This makes our custom menu items.
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Solving Tools')
  .addItem('Propose Guess', 'guessingFunction')
  .addItem('Clear Board','clearBoard')
  .addToUi();
}

//This makes the buttons work
function onEdit(e) {
  let mySS = SpreadsheetApp.getActiveSpreadsheet();
  let range = e.range;
  let sheet = range.getSheet();

  let proposeGuessButton = sheet.getRange('ProposeGuessButton');
  let clearBoardButton = sheet.getRange('ClearBoardButton');
  let messageCell = sheet.getRange('MessageCell');
  
  if (range.getA1Notation() === proposeGuessButton.getA1Notation() && range.isChecked()) {
    messageCell.setValue('Loading...');
    guessingFunction();
    range.uncheck();
  }
  else if (range.getA1Notation() === clearBoardButton.getA1Notation() && range.isChecked()) {
    clearBoard();
    range.uncheck();
  }

}

function guessingFunction() {
  //variables to reference parts of the spreadsheet
  const mySS = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = mySS.getActiveSheet();

  let guessesRange = activeSheet.getRange('Guesses'); //this is the range on the spreadsheet where the guess attempts are stored
  let guesses = guessesRange.getValues(); //this is an array with the values in the guesses range
  let solutionCell = activeSheet.getRange('Solution'); //this is the cell where the solution will go if found
  let numOfColors = activeSheet.getRange('NumOfColors').getValue();
  let numOfSlots = activeSheet.getRange('NumOfSlots').getValue();
  let duplicates = activeSheet.getRange('Duplicates').getValue();
  let messageCell = activeSheet.getRange('MessageCell');
  
  if(numOfSlots > numOfColors && !duplicates) {
    messageCell.setValue('Error: \n The current settings are incompatible. \n Program terminated. \n Try changing "Duplicates" to TRUE.');
    return;
  }

  //this function helps create pSet and gSet
  function setCreator(dataSet, slot, num) {
    if(slot > 0) {
      for(let i = 1; i <= numOfColors; i++) {
        let val = num + Math.pow(10,slot)*i;
        setCreator(dataSet, slot-1, val);
      }
    }
    else if (slot === 0) {
      for(let i = 1; i <= numOfColors; i++) {
        let val = num + i;
        //console.log(val);
        dataSet.push(val);
      }
    }
  }

  let pSet = []; //this is an array with all the possible combinations of the colors and slots (6 colors, 4 slots *usually)
  if (numOfColors === 10) {
    for (let i = 0; i < Math.pow(10,numOfSlots); i++) {
      pSet.push(i);
    }
  }
  else {
    setCreator(pSet, numOfSlots-1, 0);
  }
  

  //this removes permutations that contain duplicate peg colors from pSet
  if(!duplicates) {
    for (let i = 0; i < pSet.length; i++) {
      let guess = pSet[i];
      let guessArr = [];
      for(let j = 0; j < numOfSlots; j++) {
        guessArr.unshift(getDigit(guess, j));
      }

      for (let k = 0; k < guessArr.length; k++) {
        let a = guessArr.filter(num => num === guessArr[k]).length; 
        if (a > 1) {
          pSet.splice(i,1);
          i--;
          break;
        }
      }
    }
  }

  let gSet = []; //this is an array with all the unused guesses (it currently matches pSet) w/ a score slot on each
  if (numOfColors === 10) {
    for (let i = 0; i < Math.pow(10,numOfSlots); i++) {
      gSet.push(i);
    }
  }
  else {
    setCreator(gSet,numOfSlots-1,0);
  }
  
  for (let i = 0; i < gSet.length; i++) {
    gSet[i] = [gSet[i], 0];
  }
  console.log(pSet.length);
  

  //this function returns the value of the digit in a particular position of a number
  function getDigit(num, pos) {
    let a = Math.pow(10, pos);
    let b = a * 10;
    let c = Math.trunc(num / a);
    let d = Math.trunc(num / b) * 10;
    let e = c - d;
    return e;
  }
  //this function returns a string of B's and W's (black and white pegs in Mastermind) for a given guess and solution
  function resultOfGuess(guess, solution) {
    let val = '';
    let guessArr = [];
    for(let i = 0; i < numOfSlots; i++) {
      guessArr.unshift(getDigit(guess, i));
    }
    let solutionArr = [];
    for (let i = 0; i < numOfSlots; i++) {
      solutionArr.unshift(getDigit(solution, i));
    }

    for (let i = 1; i <= numOfColors; i++) {
      let a = guessArr.filter(num => num === i).length;
      let b = solutionArr.filter(num => num === i).length;
      let c = null;

      if (b !== 0) {
        if (a >= b) {
          c = b;
        }
        else {
          c = a;
        }
      }
      for (let j = 0; j < c; j++) {
        val = val + 'W';
      }
    }

    for (let i = 0; i < numOfSlots; i++) {
      if (guessArr[i] === solutionArr[i]) {
        val = 'B' + val.slice(0, val.length - 1);
      }
    }

    return val;
  }

  //this function removes solutions that are impossible from the input set based on the input guess
  function removeImpossibleSolutions(guess, result, possibilities) {
    let dataSet = [];
    for(let i = 0; i < possibilities.length; i++) {
      let val = possibilities[i];
      dataSet.push(val);
    }

    for (let i = 0; i < dataSet.length; i++) {
      if (resultOfGuess(guess, dataSet[i]) !== result) {
        dataSet.splice(i,1);
        i--;
      }
    }
    return dataSet;
  }

  //this removes all the impossible possibilities from pSet based on the current guesses made 
  for (let row = 0; row < guesses.length; row++) {
    let guess = guesses[row][0];
    let result = guesses[row][1].toUpperCase();
    guesses[row][1] = result;

    if (guess === '') {
      break;
    }
    let gPos = gSet.findIndex(element => element = guess);
    if (gPos >= 0) {
      gSet.splice(gPos, 1);
    }
    
    pSet = removeImpossibleSolutions(guess, result, pSet);

    guesses[row][2] = pSet.length;
  }

  console.log('Stage 1');
  console.log('Num of Unused Guesses: ' + gSet.length);
  console.log('Num of Possibilities: ' + pSet.length);
  //console.log(pSet);
  //console.log(gSet);
  console.log(guesses);

  // cutOff = gSet.findIndex(element => element[0] >= 123456)
  // console.log(cutOff)
  // gSet.splice(cutOff);
  // console.log(gSet.length);

  if (pSet.length === 0) {
    messageCell.setValue('Error:\n   The guess results are inconsistent with each other. \n Program terminated. \n Check for typos.');
  }
  else if (pSet.length === 1) {
    for (let row = 0; row < guesses.length; row++) {
      if (guesses[row][0] === '') {
        guesses[row][0] = pSet[0];
        break;
      }
    }
    guessesRange.setValues(guesses);
    solutionCell.setValue(pSet[0]);
    messageCell.setValue('Solution found! \n' + pSet[0]);
  }
  else {
  //give a score to each possible guess; this score is the maximum size of pSet after this guess is made
  for (let i = 0; i < gSet.length; i++) {
    let guess = gSet[i][0];
    console.log(guess)
    let resultArr = [];
    //console.log('Guess: ' + guess);
    for (let j = 0; j < pSet.length; j++) {
      let result = resultOfGuess(guess, pSet[j]);
      // console.log(result);
      if (!resultArr.includes(result)) {
        resultArr.push(result);

        let score = removeImpossibleSolutions(guess, result, pSet).length;
        //console.log(score);
        if (score > gSet[i][1]) {
          // console.log('replace score');
          gSet[i][1] = score;
        }
      }
      
    }
    //console.log('final scores');
    //console.log(gSet[i][1]);
  }
  
  //sort gSet in order of best score to worst; if scores match, put preference to the one that is a part of pSet; then, put preference to the lower number
  gSet.sort(
    (a, b) => {
      if (a[1] < b[1]) {
        return -1; 
      }
      else if (a[1] > b[1]) {
        return 1;
      }
      else {
        let aIndex = pSet.indexOf(a[0]);
        let bIndex = pSet.indexOf(b[0]);
        if (aIndex !== -1 && bIndex === -1) {
          return -1;
        }
        else if (aIndex === -1 && bIndex !== -1) {
          return 1; 
        }
        else {
          return a[0] - b[0]; 
        }
      }
    }
  )

  for (let row = 0; row < guesses.length; row++) {
    if (guesses[row][0] === '') {
      guesses[row][0] = gSet[0][0];
      guesses[row][2] = 'Max: ' + gSet[0][1];
      break;
    }
  }

  guessesRange.setValues(guesses);
  messageCell.setValue('An optimal guess has been determined. \n' + gSet[0][0]);

  console.log('Stage 2');
  console.log('Num of Possibilities: ' + pSet.length);
  console.log(gSet);

  }
}

function clearBoard() {
  //variables to reference parts of the spreadsheet
  const mySS = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = mySS.getActiveSheet();

  activeSheet.getRange('Guesses').clearContent();
  activeSheet.getRange('Solution').clearContent();
  activeSheet.getRange('MessageCell').clearContent();
}