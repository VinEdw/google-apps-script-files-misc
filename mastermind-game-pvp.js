//This function makes all the buttons work and has all their functionality
function onEdit(e) {
  let mySS = SpreadsheetApp.getActiveSpreadsheet();
  let range = e.range;
  let sheet = range.getSheet();
  
  //section for the buttons in the Settings sheet
  if(sheet.getSheetName() === 'Settings') {
    let newGameButton = sheet.getRange('NewGame');
    let messageCell = sheet.getRange('AlertCell');
    let resetScoreboard = sheet.getRange('ResetScoreboard');

    //section for the newGameButton
    if (range.getA1Notation() === newGameButton.getA1Notation()) {
      if(range.isChecked()) {
        messageCell.setValue('New Game');
        let eachSolutionRange = sheet.getRange('EachSolution');
        if (sheet.getRange('NumOfPossibilities').getValue() === '#NUM!') {
          messageCell.setValue('Error: The current settings are incompatible. Try changing "Duplicates" to TRUE.');
        }

        else {
          for(let i = 1; i <= 4; i++) {
            let playerSheetName = 'Player ' + i;
            let playerSheet = mySS.getSheetByName(playerSheetName);

            playerSheet.getRangeList(['Guesses','MessageCell']).clearContent();
            playerSheet.getRange('Solution').setFontColor('#f9cb9c');
            playerSheet.getRange('ColoredLight').setBackground('red');
            playerSheet.getRange('ReadyCheckBox').uncheck();

          }

          eachSolutionRange.clearContent();
          let eachSolution = eachSolutionRange.getValues();
          let setManual = sheet.getRange('ManualSolutions').getValue();
          if (!setManual) {
            let sequence = randomSequence();
            for (let row = 0; row < eachSolution.length; row++) {
              eachSolution[row][0] = sequence;
            }
            
            eachSolutionRange.setValues(eachSolution);
          }

          messageCell.clearContent();
        }

        range.uncheck();
      }
    }
    //section for the resetScoreboard button
    else if (range.getA1Notation() === resetScoreboard.getA1Notation()) {
      if (range.isChecked()) {
        messageCell.setValue('Reset Scoreboard');
        let scoreboardRange = mySS.getSheetByName('Player 1').getRange('Scoreboard');
        let scoreboard = scoreboardRange.getValues();
        
        for (let i = 0; i < scoreboard[0].length; i++) {
          scoreboard[0][i] = 0; 
        }
        scoreboardRange.setValues(scoreboard);
        messageCell.clearContent();
        range.uncheck();
      }
    }
  }
  
  //section for the non Settings sheets; the Player sheets
  else if (sheet.getSheetName() !== 'Settings') {
    let gCheck = sheet.getRange('CheckButton');
    let readyCheckBox = sheet.getRange('ReadyCheckBox');
    let messageCell = sheet.getRange('MessageCell')
    

    //section for the guess checking button
    if (range.getA1Notation() === gCheck.getA1Notation()) {
      if (range.isChecked()) {
        if (sheet.getRange('ColoredLight').getBackground() !== '#ff0000') {
          messageCell.setValue('Checking');

          let guessesRange = sheet.getRange('Guesses');
          let guesses = guessesRange.getValues();

          for (let i = 0; i < guesses.length; i++) {
            if (i === guesses.length -1 || guesses[i + 1][0] === '') {
              let guess = guesses[i][0];
              let solution = sheet.getRange('Solution').getValue();
              let result = resultOfGuess(guess, solution);
              guesses[i][1] = result;
              guessesRange.setValues(guesses);
              messageCell.clearContent();

              if (guess === solution) {
                sheet.getRange('Solution').setFontColor('black');
                let name = sheet.getRange('PlayerName').getValue();

                if(sheet.getRange('ColoredLight').getBackground() === '#00ff00') {
                  let numOfPlayers = mySS.getSheetByName('Settings').getRange('numOfPlayers').getValue();
                  for (let i = 1; i <= numOfPlayers; i++) {
                    let playerSheetName = 'Player ' + i;
                    let playerSheet = mySS.getSheetByName(playerSheetName);

                    playerSheet.getRange('ColoredLight').setBackground('yellow');
                  }
                  for (let i = 1; i <= numOfPlayers; i++) {
                    let playerSheetName = 'Player ' + i;
                    let playerSheet = mySS.getSheetByName(playerSheetName);

                    playerSheet.getRange('MessageCell').setValue(name + ' won!');
                  }

                  if (mySS.getSheetByName('Settings').getRange('ScoreboardToggle').getValue()) {
                    let scoreboardRange = mySS.getSheetByName('Player 1').getRange('Scoreboard');
                    let scoreboard = scoreboardRange.getValues();
                    let slotNum = Number(sheet.getSheetName()[7]) - 1;

                    scoreboard[0][slotNum]++;

                    scoreboardRange.setValues(scoreboard);
                  }

                }
                else {
                  messageCell.setValue('Solution found!');
                }
              }

              break;
            }
          }

        }
        range.uncheck();
      }

    }
    //section for the readyCheckBox button
    else if (range.getA1Notation() === readyCheckBox.getA1Notation()) {
      if (range.isChecked()) {
        let numOfPlayers = mySS.getSheetByName('Settings').getRange('numOfPlayers').getValue();
        let numOfSlots = mySS.getSheetByName('Settings').getRange('NumOfSlots').getValue();
        let numOfColors = mySS.getSheetByName('Settings').getRange('NumOfColors').getValue();
        let duplicates = mySS.getSheetByName('Settings').getRange('Duplicates').getValue();
        let eachSolution = mySS.getSheetByName('Settings').getRange('EachSolution').getValues();
        let allReady = true; 

        //This checks if all the active players have checked their ready button
        for(let i = 1; i <= numOfPlayers; i++) {
          let playerSheetName = 'Player ' + i;
          let playerSheet = mySS.getSheetByName(playerSheetName);

          let ready = playerSheet.getRange('ReadyCheckBox').getValue();
          if (!ready || playerSheet.getRange('ColoredLight').getBackground() !== '#ff0000') {
            allReady = false;
            break;
          }
        }

        //This tells all the active players that the game is doing its pre checks
        if (allReady) {
          for (let i = 1; i <= numOfPlayers; i++) {
            let playerSheetName = 'Player ' + i;
            let playerSheet = mySS.getSheetByName(playerSheetName);

            playerSheet.getRange('MessageCell').setValue('Loading...');
          }
        }

        //This checks if the solutions in the setting sheet are of the proper length
        if (allReady) {
          for (let i = 0; i < numOfPlayers; i++) {
            let solution = eachSolution[i][0];

            if(!Number.isInteger(solution)) {
              allReady = false;
              for (let i = 1; i <= numOfPlayers; i++) {
                let playerSheetName = 'Player ' + i;
                let playerSheet = mySS.getSheetByName(playerSheetName);
                playerSheet.getRange('ReadyCheckBox').uncheck();
                playerSheet.getRange('MessageCell').setValue('Error: At least one of the solutions is not an integer.');
              }
              break;
            }
          }
        }

        //This checks if the solutions in the settings sheet are of the proper length
        if (allReady) {
          for (let i = 0; i < numOfPlayers; i++) {
            let solution = eachSolution[i][0];

            if (numOfColors === 10) {
              if (`${solution}`.length > numOfSlots) {
                allReady = false;
                break;
              }
            }
            else if (`${solution}`.length !== numOfSlots) {
              allReady = false;
              break;
            }
          }

          if (!allReady) {
            for (let i = 1; i <= numOfPlayers; i++) {
              let playerSheetName = 'Player ' + i;
              let playerSheet = mySS.getSheetByName(playerSheetName);
              playerSheet.getRange('ReadyCheckBox').uncheck();
              playerSheet.getRange('MessageCell').setValue('Error: At least one of the solutions is not of the proper length.');
            }
          }
        }

        //This checks if the solution uses any numbers/colors that are not in the set range
        if (allReady) {
          let colorList = [];
          for (let i = 1; i <= numOfColors; i++) {
            if (i === 10) {
              colorList.push(0);
            }
            else {
              colorList.push(i);
            }
          }

          for (let i = 0; i < numOfPlayers; i++) {
            let solution = eachSolution[i][0];

            for (let j = 0; j < numOfSlots; j++) {
              let color = getDigit(solution,j);
              if (!colorList.includes(color)) {
                allReady = false;
                for (let i = 1; i <= numOfPlayers; i++) {
                  let playerSheetName = 'Player ' + i;
                  let playerSheet = mySS.getSheetByName(playerSheetName);
                  playerSheet.getRange('ReadyCheckBox').uncheck();
                  playerSheet.getRange('MessageCell').setValue('Error: At least one of the solutions has a digit not in the set available.');
                }
                break;
              }
            }
            if(!allReady) {
              break;
            }
          }
        }

        //This checks if the solution has duplicates if duplicates are turned off
        if (allReady && !duplicates) {
          for (let i = 0; i < numOfPlayers; i++) {
            let solution = eachSolution[i][0];
            let solutionArr = [];
            for (let j = 0; j < numOfSlots; j++) {
              solutionArr.unshift(getDigit(solution, j));
            }

            for (let k = 0; k < solutionArr.length; k++) {
              let a = solutionArr.filter(num => num === solutionArr[k]).length;
              if (a > 1) {
                allReady = false;
                for (let i = 1; i <= numOfPlayers; i++) {
                  let playerSheetName = 'Player ' + i;
                  let playerSheet = mySS.getSheetByName(playerSheetName);
                  playerSheet.getRange('ReadyCheckBox').uncheck();
                  playerSheet.getRange('MessageCell').setValue('Error: At least one of the solutions has duplicate digits even though duplicates are off.');
                }
                break;
              }
            }
            if (!allReady) {
              break;
            }
          }
        }

        //Change the lights to green and say go after all the checks for valid solutions have been performed
        if (allReady) {
          for (let i = 1; i <= numOfPlayers; i++) {
            let playerSheetName = 'Player ' + i;
            let playerSheet = mySS.getSheetByName(playerSheetName);

            playerSheet.getRange('ColoredLight').setBackground('#00ff00');
          }
          for (let i = 1; i <= numOfPlayers; i++) {
            let playerSheetName = 'Player ' + i;
            let playerSheet = mySS.getSheetByName(playerSheetName);

            playerSheet.getRange('MessageCell').setValue(`Go!\nPlayers: ${numOfPlayers}`);
          }
        }
      }
    }
  }
}

function getDigit(num, pos) {
    let a = Math.pow(10, pos);
    let b = a * 10;
    let c = Math.trunc(num / a);
    let d = Math.trunc(num / b) * 10;
    let e = c - d;
    return e;
}

function resultOfGuess(guess, solution) {
  let mySS = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = mySS.getSheetByName('Settings');
  let numOfSlots = settingsSheet.getRange('NumOfSlots').getValue();
  let numOfColors = settingsSheet.getRange('NumOfColors').getValue();
  
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
    let colorNum = i;
    if (colorNum === 10) {
      colorNum = 0;
    }
    let a = guessArr.filter(num => num === colorNum).length;
    let b = solutionArr.filter(num => num === colorNum).length;
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
  if (val === '') {
    val = '-';
  }

  return val;
}

function randomSequence() {
  let mySS = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = mySS.getSheetByName('Settings');
  let numOfSlots = settingsSheet.getRange('NumOfSlots').getValue();
  let numOfColors = settingsSheet.getRange('NumOfColors').getValue();
  let duplicates = settingsSheet.getRange('Duplicates').getValue();

  let num = 0;
  if (duplicates) {
    for (let i = 0; i < numOfSlots; i++) {
      let k = Math.floor(Math.random() * numOfColors) + 1;
      if (k === 10) {
        k = 0;
      }
      num += Math.pow(10, i) * k;
    }
  }
  else {
    let colorList = [];
    for (let i = 1; i <= numOfColors; i++) {
      if (i === 10) {
        colorList.push(0);
      }
      else {
        colorList.push(i);
      }
    }

    for (let i = 0; i < numOfSlots; i++) {
      let a = Math.floor(Math.random() * colorList.length);
      let b = colorList[a];
      colorList.splice(a,1);
      num += Math.pow(10, i) * b;
    }
  }
  
  return num;
}