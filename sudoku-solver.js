//this function makes the buttons work
function onEdit(e) {
  let range = e.range;
  let sheet = range.getSheet();

  let solveButton = sheet.getRange('SolveButton');
  let clearButton = sheet.getRange('ClearButton');

  if (range.getA1Notation() === solveButton.getA1Notation() && range.isChecked()) {
    backtrackingSolver();
    range.uncheck();
  }
  else if (range.getA1Notation() === clearButton.getA1Notation() && range.isChecked()) {
    clearBoard();
    range.uncheck();
  }
}

// This makes our custom menu items.
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Solving Tools')
  .addItem('Solve', 'backtrackingSolver')
  .addItem('Clear', 'clearBoard')
  .addToUi();
}

//this function clears the board
function clearBoard() {
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('Board').setFontWeight('normal').clearContent();
}

function backtrackingSolver() {
  let mySS = SpreadsheetApp.getActiveSpreadsheet();
  let activeSheet = mySS.getActiveSheet();
  let boardRange = activeSheet.getRange('Board');
  let board = boardRange.getValues();
  let numOfRows = board.length;
  let numOfColumns = board[0].length;
  let messageCell = activeSheet.getRange('MessageCell');
  messageCell.setValue('Loading...');
  
  //check if the board is solvable
  if (!checkBoard(board)) {
    console.log('Error: Board is not solvable');
    messageCell.setValue('Error: Board is not solvable');
    return;
  }

  let optionsArr = boardRange.getValues();
  for (let r = 0; r < numOfRows; r++) {
    for (let c = 0; c < numOfRows; c++) {
      if (optionsArr[r][c] === '') {
        possibilitiesArr = [];
        for (let i = 1; i <= numOfRows; i++) {
          optionsArr[r][c] = i;
          if (checkBoard(optionsArr)) {
            possibilitiesArr.push(i);
          }
        }
        if (possibilitiesArr.length === 0) {
          console.log('Error: Board is not solvable');
          messageCell.setValue('Error: Board is not solvable');
          return;
        }
        optionsArr[r][c] = possibilitiesArr;
      }
    }
  }

  //this function plays sudoku without guess and check to try making the optionsArr smaller (limit the possibilities)
  function limitOptionsArr() {
    let updating = false;
    do {
      updating = false;
      //console.log(board);
      //check for cells than only have one possible number they can accept
      if (!updating) {
        for (let r = 0; r < numOfRows; r++) {
          for (let c = 0; c < numOfColumns; c++) {
            if (Array.isArray(optionsArr[r][c]) && optionsArr[r][c].length === 1) {
              optionsArr[r][c] = optionsArr[r][c][0];
              board[r][c] = optionsArr[r][c];
              //console.log(board[r][c] + ' at ' + 'R: ' + r + ' C: ' + c);
              //console.log('via length 1');
              updating = true;
            } 
          }
        }
      }
      
      //check for cells where they are the only one in a row that can accept a certain number
      if (!updating) {
        for (let r = 0; r < numOfRows; r++) {
          let possibilitiesArr = [];
          for (let c = 0; c < numOfColumns; c++) {
            if (Array.isArray(optionsArr[r][c])) {
              possibilitiesArr.push(optionsArr[r][c]);
            }
          }
          for (let i = 1; i <= numOfRows; i++) {
            let count = possibilitiesArr.flat().filter(x => x === i).length;
            if (count === 1) {
              for (let c = 0; c < numOfColumns; c++) {
                if (Array.isArray(optionsArr[r][c]) && optionsArr[r][c].includes(i)) {
                  optionsArr[r][c] = i;
                  board[r][c] = i;
                  //console.log(i + ' at ' + 'R: ' + r + ' C: ' + c);
                  //console.log('via row requirement')
                  updating = true;
                  break;
                }
              }
            }
          }
        }
      }

      //check for cells where they are the only one in a column that can accept a certain number
      if (!updating) {
        for (let c = 0; c < numOfColumns; c++) {
          let possibilitiesArr = [];
          for (let r = 0; r < numOfRows; r++) {
            if (Array.isArray(optionsArr[r][c])) {
              possibilitiesArr.push(optionsArr[r][c]);
            }
          }
          for (let i = 1; i <= numOfRows; i++) {
            let count = possibilitiesArr.flat().filter(x => x === i).length;
            if (count === 1) {
              for (let r = 0; r < numOfRows; r++) {
                if (Array.isArray(optionsArr[r][c]) && optionsArr[r][c].includes(i)) {
                  optionsArr[r][c] = i;
                  board[r][c] = i;
                  //console.log(i + ' at ' + 'R: ' + r + ' C: ' + c);
                  //console.log('via column requirement');
                  updating = true;
                  break;
                }
              }
            }
          }
        }
      }

      //check for cells where they are the only one in a mini box that can accept a certain number
      if (!updating) {
        for (let sr = 0; sr < numOfRows; sr += 3) {
          for (let sc = 0; sc < optionsArr[sr].length; sc += 3) {
            let possibilitiesArr = [];
            for (let r = sr; r < sr + 3; r++) {
              for (let c = sc; c < sc + 3; c++) {
                if (Array.isArray(optionsArr[r][c])) {
                  possibilitiesArr.push(optionsArr[r][c]);
                }
              }
            }
            for (let i = 1; i <= numOfRows; i++) {
              let count = possibilitiesArr.flat().filter(x => x === i).length;
              if (count === 1) {
                let leave = false;
                for (let r = sr; r < sr + 3; r++) {
                  if (leave) {
                    break;
                  }
                  for (let c = sc; c < sc + 3; c++) {
                    if (Array.isArray(optionsArr[r][c]) && optionsArr[r][c].includes(i)) {
                      optionsArr[r][c] = i;
                      board[r][c] = i;
                      //console.log(i + ' at ' + 'R: ' + r + ' C: ' + c);
                      //console.log('via box requirement');
                      updating = true;
                      leave = true;
                      break;
                    }
                  }
                }
              }
            }
          }
        }
      }

      //if something changed about the board, update the rest of the optionsArr
      if (updating) {
        for (let r = 0; r < numOfRows; r++) {
          for (let c = 0; c < numOfColumns; c++) {
            if (Array.isArray(optionsArr[r][c])) {
              for (let i = optionsArr[r][c].length - 1; i >= 0; i--) {
                board[r][c] = optionsArr[r][c][i];
                if (!(checkRow(r, board) && checkColumn(c, board) && checkBox(r, c, board))) {
                  optionsArr[r][c].splice(i,1);
                }
              }
              board[r][c] = '';
            }
          }
        }
      }
      
    } while (updating);
    console.log('New board');
    console.log(board);
  }

  console.log('Initial optionsArr');
  console.log(optionsArr);
  limitOptionsArr();
  console.log('New optionsArr');
  console.log(optionsArr);

  //check if the board is solvable
  if (!checkBoard(board)) {
    console.log('Error: Board is not solvable');
    messageCell.setValue('Error: Board is not solvable');
    return;
  }

  boldDataInBoard();

  function jumpForward(pos, options) {
    let row = pos[0];
    let column = pos[1];
    do {
      row++;
      if (row === options.length) {
        row = 0; 
        column++;
        if (column === options[row].length) {
          return [row, column];
        }
      }
    } while(!Array.isArray(options[row][column]));
    return [row, column];
  }

  function jumpBackward(pos, options) {
    let row = pos[0];
    let column = pos[1];
    do {
      row--;
      if (row === -1) {
        row = options.length - 1;
        column--;
        if (column === -1) {
          return [row, column];
        }
      }
    } while (!Array.isArray(options[row][column]));
    return [row, column];
  }

  function backtrack(pos, options) {
    let row = pos[0];
    let column = pos[1];

    if (column === -1 || column === options[row].length) {
      //console.log(board);
      return undefined;
    }

    if (Array.isArray(options[row][column])) {
      if (board[row][column] === '') {
        board[row][column] = options[row][column][0];
        if (checkRow(row, board) && checkColumn(column, board) && checkBox(row, column, board)) {
          let nPos = jumpForward(pos, options);
          return nPos;
        }
        else {
          return pos;
        }
      }
      else {
        let val = board[row][column];
        let i = options[row][column].indexOf(val);

        if (i === options[row][column].length - 1) {
          board[row][column] = '';
          let nPos = jumpBackward(pos, options);
          return nPos;
        }
        else {
          board[row][column] = options[row][column][i+1];
          if (checkRow(row, board) && checkColumn(column, board) && checkBox(row, column, board)) {
            let nPos = jumpForward(pos, options);
            return nPos;
          }
          else {
            return pos;
          }
        }
      }
    }
    else {
      let nPos = jumpForward(pos, options);
      return nPos;
    }

  }

  let pos = [0, 0];
  let finished = false;
  let numOfLoops = 0;
  while (!finished) {
    pos = backtrack(pos, optionsArr);
    numOfLoops++;
    //console.log(board);
    //console.log(pos);
    if (pos === undefined) {
      finished = true;
    }
  }

  console.log('Final solved board');
  console.log(board);
  console.log(numOfLoops);
  boardRange.setValues(board);

  messageCell.clearContent();
}

//checks if the input board has any problems in a given row
function checkRow(row, board) {
  let numOfColumns = board[row].length;
  let rowNumbers = [];
  for (let c = 0; c < numOfColumns; c++) {
    let val = board[row][c];
    if (Number.isInteger(val) && val !== 0) {
      if (rowNumbers.includes(val)) {
        return false;
      }
      else {
        rowNumbers.push(val);
      }
    }
  }
  return true;
}

//checks if the input board has any problems in a given column
function checkColumn(column, board) {
  let numOfRows = board.length;
  let columnNumbers = [];
  for (let r = 0; r < numOfRows; r++) {
    let val = board[r][column];
    if (Number.isInteger(val) && val !== 0) {
      if (columnNumbers.includes(val)) {
        return false;
      }
      else {
        columnNumbers.push(val);
      }
    }
  }
  return true;
}

//checks if the input board has any problems in a given 3x3 box
function checkBox(row, column, board) {
  while (!(row === 0 || row === 3 || row === 6)) {
    row--;
  }
  while (!(column === 0 || column === 3 || column === 6)) {
    column--;
  }

  let boxNumbers = [];
  for (let r = row; r < row + 3; r++) {
    for (let c = column; c < column + 3; c++) {
      let val = board[r][c];
      if (Number.isInteger(val) && val !== 0) {
        if (boxNumbers.includes(val)) {
          return false;
        }
        else {
          boxNumbers.push(val);
        }
      }
    }
  }
  return true;
}

//this function checks if the 9x9 sudoku board input has no mistakes; returns true or false
function checkBoard(board) {
  let numOfRows = board.length;
  let numOfColumns = board[0].length;
  //check all the rows for no duplicates within a row
  for (let r = 0; r < numOfRows; r++) {
    if (!checkRow(r, board)) {
      return false;
    }
  }

  //check all the columns for no duplicates within a column
  for (let c = 0; c < numOfColumns; c++) {
    if (!checkColumn(c, board)) {
      return false;
    }
  }
  
  //check all the mini boxes for no duplicates within the section
  for (let sr = 0; sr < numOfRows; sr += 3) {
    for (let sc = 0; sc < numOfColumns; sc += 3) {
      if (!checkBox(sr, sc, board)) {
        return false;
      }
    }
  }
  return true;
}

function boldDataInBoard() {
  let mySS = SpreadsheetApp.getActiveSpreadsheet();
  let activeSheet = mySS.getActiveSheet();
  let boardRange = activeSheet.getRange('Board')
  let board = boardRange.getValues();

  let startCell = boardRange.getCell(1,1);
  let startRow = startCell.getRow()
  let startColumn = startCell.getColumn();
  //console.log(startRow);
  //console.log(startColumn);  

  let boldArr = boardRange.getValues();

  for (let r = 0; r < board.length; r++) {
    for (let c = 0; c < board[r].length; c++) {
      if (board[r][c]) {
        boldArr[r][c] = 'bold';
      }
      else {
        boldArr[r][c] = 'normal'
      }
    }
  }

  boardRange.setFontWeights(boldArr);
}

function solverV2() {
  let mySS = SpreadsheetApp.getActiveSpreadsheet();
  let activeSheet = mySS.getActiveSheet();
  let boardRange = activeSheet.getRange('Board');
  let board = boardRange.getValues();
  let numOfRows = board.length;
  let numOfColumns = board[0].length;
  //let messageCell = activeSheet.getRange('MessageCell');
  //messageCell.setValue('Loading...');


  function backtrack() {
    for(let r = 0; r < numOfRows; r++) {
      for(let c = 0; c < numOfColumns; c++) {
        if(board[r][c] === '') {
          for(let n = 1; n <= 9; n++) {
            board[r][c] = n;
            if (checkRow(r, board) && checkColumn(c, board) && checkBox(r, c, board)) {
              backtrack()
              board[r][c] = '';
              //console.log(board);
            }
          }
          board[r][c] = '';
          return;
        }
      }
    }
    console.log(board)
  }

  backtrack();
}