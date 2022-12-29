function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
function forEachRangeCell(thisSheet, matchString, f) {
  const range = thisSheet.getDataRange();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  for (let i = 1; i <= numCols; i++) {
    for (let j = 1; j <= numRows; j++) {
      const cell = range.getCell(j, i);
      const cellLoc = columnToLetter(i) + String(j);
      if (cell.getFormula().toLowerCase().includes(matchString.toLowerCase())) {
        f(j, i, cellLoc);
      }
    }
  }
}
function setGlobalVar(thisKey,thisValue) {
  PropertiesService.getScriptProperties().setProperty(thisKey, thisValue);
}
function getGlobalVar(thisKey) {
  return PropertiesService.getScriptProperties().getProperty(thisKey);
}
function A1Print(ss, x, loc='A1') {
  ss.getRange(loc).setValue(JSON.stringify(x));
}

/**
 • Ensures there's room for any function output
 *
 * @param {Array} Filter_Function input the function that needs to always have room for an output.
 * @return moves cells below to make room
 * @customfunction
*/
function MakeRoomFor(Filter_Function) {
  var loc = SpreadsheetApp.getActive().getActiveSheet().getActiveCell().getA1Notation();
  setGlobalVar('makeRoomSize'+loc, JSON.stringify([Filter_Function[0].length, Filter_Function.length]));
  return Filter_Function;
}

function onEdit(e) {
  const ss = e.range.getSheet();
  //A1Print(ss,'running: MakeRoomFor_onEdit');
  forEachRangeCell(ss, 'MakeRoomFor(', (r, c, loc) => {
    //A1Print(ss,'found:', 'B1');
    const cell = ss.getRange(loc);
    if (!cell.getValue().toLowerCase().includes("#ref")) {
      return;
    }
    
    const makeRoomSizeTXT = getGlobalVar('makeRoomSize'+loc);
    //A1Print(ss, 'funcRunning');
    const makeRoomSize = JSON.parse(makeRoomSizeTXT);
    makeRoomSize[0] = makeRoomSize[0];

    //find out how many cells are blank
    let blankCellsArray = Array(makeRoomSize[0]).fill(0);
    for (let x=0; x < makeRoomSize[0]; x++) {
      for (let i=0; i < makeRoomSize[1]; i++) {
        blankCellsArray[x]++;
        if (!cell.offset((i+1),x).isBlank()) {
          //cell.offset(i,x).setComment("lastBlank");
          break;
        }
      }
    }
    let blankCells = blankCellsArray.reduce((a, b) => Math.min(a, b));
    //A1Print(ss,blankCellsArray);

    //sub actual height from blank cells num
    const toShift = makeRoomSize[1] - blankCells;
    
    //shift that many cells
    ss.getRange(r+1, c, (ss.getLastRow()-r)+1, makeRoomSize[0]).moveTo(ss.getRange(r+1+toShift, c));
    //Logger.log(String(c) + ', ' + String(r+1) + ', ' + String(20) + ', ' + String(makeRoomSize[0]));

    //Comment debug data
    //cell.setComment(makeRoomSizeTXT + ";  " + loc + ";  blank:" + JSON.stringify(blankCellsArray) + ";\nmustMove:" + toShift);
  });
}
