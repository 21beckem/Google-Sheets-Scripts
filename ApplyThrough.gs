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
  //A1Print(ss, APfunctions);

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
    if (m[1].HEREcolumnToListenTo == columnToLetter(e.range.getColumn())) {
      //toast('found right col and page. And: ' + e.value);
    } else { continue; }


    // check if that function cell is still an ApplyThrough() function
    if (ss.getRange(m[1].ThisCellAddress).getFormula().toLowerCase().includes('applythrough(')) {
    } else {
      delete APfunctions[m[0]];
      setGlobalVar('ApplyThroughFunctions', JSON.stringify(APfunctions));
      continue;
    }



    // if all that's true, do the applying through
    const toChangeTo = e.value;
    e.range.setValue(undefined);
    const orgss = SpreadsheetApp.getActive().getSheetByName(m[1].THEREpageName);
    const buddyLoc1 = m[1].HEREcolumnToRefrenceTo + String(e.range.getRow());
    const buddyVal = ss.getRange(buddyLoc1).getValue();
    toast(buddyVal);

    // get the ref col

    // find buddyVal in that

    // set corresponding columnToChange value

    // dance!
  }
}

function ApplyThrough(ThisCellAddress, HEREcolumnToListenTo, HEREcolumnToRefrenceTo, THEREpageName, THEREcolumnToChange, THEREcolumnToRefrence, Filter_Function) {
  const HEREpageName = SpreadsheetApp.getActiveSheet().getName();
  
  const currentFunctions = JSON.parse(getGlobalVar('ApplyThroughFunctions') || '{}');

  currentFunctions["'" + HEREpageName + "'!" + ThisCellAddress] = {
    HEREpageName : HEREpageName,
    ThisCellAddress : ThisCellAddress,
    HEREcolumnToListenTo : HEREcolumnToListenTo,
    HEREcolumnToRefrenceTo : HEREcolumnToRefrenceTo,
    THEREpageName : THEREpageName,
    THEREcolumnToChange : THEREcolumnToChange,
    THEREcolumnToRefrence : THEREcolumnToRefrence
  }
  setGlobalVar('ApplyThroughFunctions', JSON.stringify(currentFunctions));

  //return JSON.stringify(currentFunctions);
  return Filter_Function;
}
