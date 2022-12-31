function cellA1ToIndex(e,r){r=0==(r=r||0)?0:1;var n=e.match(/(^[A-Z]+)|([0-9]+$)/gm);if(2!=n.length)throw new Error("Invalid cell reference");e=n[0];return{row:rowA1ToIndex(n[1],r),col:colA1ToIndex(e,r)}}function colA1ToIndex(e,r){if("string"!=typeof e||2<e.length)throw new Error("Expected column label.");r=0==(r=r||0)?0:1;var n="A".charCodeAt(0),o=e.charCodeAt(e.length-1)-n;return 2==e.length&&(o+=26*(e.charCodeAt(0)-n+1)),o+r}function rowA1ToIndex(e,r){return e-1+(r=0==(r=r||0)?0:1)}
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
function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}
Array.prototype.indexOf2d = function(item) {
  for(var k = 0; k < this.length; k++){
    if(JSON.stringify(this[k]) == JSON.stringify(item)){
      return k;
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
function toast(s) {
  SpreadsheetApp.getActive().toast(JSON.stringify(s));
}

function onEdit(e) {
  const ss = e.source.getActiveSheet();
  const APfunctions = JSON.parse(getGlobalVar('ApplyThroughFunctions') || '{}');
  A1Print(ss, APfunctions);

  // check if the edited cell is on one of these pages
  let found = false;
  let matches = Array();
  for (const [key, rel] of Object.entries(APfunctions)) {
    if (rel.HEREpageName == ss.getName()) {
      found = true;
      matches.push([key, rel]);
    }
  }
  if (found == false) { return; }
  //A1Print(ss,matches);

  for (let i = 0; i < matches.length; i++) {
    const m = matches[i];
    // check if the edited cell is in one of the tracked coloums
    if (m[1].HEREcolumnToListenTo.split(',').includes( columnToLetter( e.range.getColumn() - 1 ) ) ) {
      //toast('found right col and page. And: ' + e.value);
    } else { continue; }


    // check if that function cell is still an ApplyThrough() function
    if (ss.getRange(m[1].ThisCellAddress).getFormula().toLowerCase().includes('applythrough(')) {
    } else {
      delete APfunctions[m[0]];
      setGlobalVar('ApplyThroughFunctions', JSON.stringify(APfunctions));
      continue;
    }
    if (e.range.getRow() < ss.getRange(m[1].ThisCellAddress).getRow()) { continue; }
    e.range.setValue(undefined);




    // if all that's true, do the applying through
    const orgss = SpreadsheetApp.getActive().getSheetByName(m[1].THEREpageName);
    SpreadsheetApp.flush();
    const buddyLoc1 = m[1].HEREcolumnToRefrenceTo + String(e.range.getRow());
    //toast(buddyLoc1);
    let waiting = true;
    while (waiting) {
      if (ss.getRange(m[1].ThisCellAddress).getValue() == "#ERROR") {
        SpreadsheetApp.flush();
      } else {
        waiting = false;
      }
    }
    const buddyVal = ss.getRange(buddyLoc1).getValue();
    //toast(buddyVal);

    // get the ref col
    const refC = orgss.getRange(m[1].THEREcolumnToRefrence + ':' + m[1].THEREcolumnToRefrence).getValues();
    //toast( refC, 'A2');

    // find buddyVal in that
    const buddyLoc2 = refC.indexOf2d([buddyVal]) + 1;
    //toast( buddyLoc2, 'A3');

    // get corresponding col to edit
    const corCol = columnToLetter((e.range.getColumn() - m[1].ThisCellCoords.col) + letterToColumn(m[1].THEREfirstColOfData) - 1);
    //toast(corCol);

    // set corresponding columnToChange value
    orgss.getRange(corCol + String(buddyLoc2)).setValue(e.value);

    // dance!
  }
}

function ApplyThrough(ThisCellAddress, HEREcolumnToListenTo, HEREcolumnToRefrenceTo, THEREpageName, THEREfirstColOfData, Filter_Function) {
  const HEREpageName = SpreadsheetApp.getActiveSheet().getName();
  const ThisCellCoords = cellA1ToIndex(ThisCellAddress);
  const ColLetterList = HEREcolumnToListenTo.split(',').map((el) => {
      return columnToLetter(parseInt(el) + ThisCellCoords.col);
  });
  
  const currentFunctions = JSON.parse(getGlobalVar('ApplyThroughFunctions') || '{}');

  currentFunctions["'" + HEREpageName + "'!" + ThisCellAddress] = {
      HEREpageName : HEREpageName,
      ThisCellAddress : ThisCellAddress,
      ThisCellCoords : ThisCellCoords,
      HEREcolumnToListenTo : ColLetterList.join(','),
      HEREcolumnToRefrenceTo : columnToLetter(parseInt(HEREcolumnToRefrenceTo) + ThisCellCoords.col + 1),
      THEREpageName : THEREpageName,
      THEREfirstColOfData: THEREfirstColOfData,
      THEREcolumnToRefrence : columnToLetter(letterToColumn(THEREfirstColOfData) + parseInt(HEREcolumnToRefrenceTo))
  }
  setGlobalVar('ApplyThroughFunctions', JSON.stringify(currentFunctions));

  //return JSON.stringify(currentFunctions);
  return Filter_Function;
}
