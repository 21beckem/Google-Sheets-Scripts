// helpful functions
function forSpecificFormulaCells(thisSheet, matchString, f) {
  const allForms = thisSheet.getDataRange().getFormulas();
  //A1Print(thisSheet, allForms);
  const numRows = allForms.length;
  const numCols = allForms[0].length;

  for (let i = 0; i < numCols; i++) {
    for (let j = 0; j < numRows; j++) {
      if (allForms[j][i].toLowerCase().includes(matchString.toLowerCase())) {
        //A1Print(thisSheet, allForms[j][i]);
        const cellLoc = columnToLetter(i+1) + String(j+1);
        const cVal = String(thisSheet.getRange(cellLoc).getValue()).toLowerCase();
        //A1Print(thisSheet, cVal);
        if (!(cVal.includes('#error') || cVal.includes("#ref"))) {
          //A1Print(thisSheet, "scripting", 'C1');
          f(j+1, i+1, cellLoc);
        }

      }
    }
  }
}
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
 • Finds changes in the child formula and shifts another column to stay matching
 *
 * @param {Array} Filter_Function input the function that needs to always have room for an output.
 * @param {String} ColumnToTrack input AS TEXT the name of the column you want to TRACK.
 * @param {String} ColumnToMatch input AS TEXT the name of the column you want to move to match the previous parameter's column.
 * @return moves cells below to make room
 * @customfunction
*/
function TrackAndMatch(Filter_Function, ColumnToTrack, ColumnToMatch) {
  if (typeof ColumnToTrack != 'string') {
    throw 'The second parameter must be a string!   For example: =TrackAndMatch(FILTER( ... ), "C", "D")';
  }
  if (typeof ColumnToMatch != 'string') {
    throw 'The third parameter must be a string!   For example: =TrackAndMatch(FILTER( ... ), "C", "D")';
  }
  var loc = SpreadsheetApp.getActive().getActiveSheet().getActiveCell().getA1Notation();
  setGlobalVar('NEW_trackingData'+loc, JSON.stringify(Filter_Function));
  setGlobalVar('trackingColumn'+loc, ColumnToTrack);
  setGlobalVar('matchingColumn'+loc, ColumnToMatch);
  //return JSON.parse(getGlobalVar('NEW_trackingData'));
  return Filter_Function;
}

function onEdit(e) {
  const ss = e.range.getSheet();
  //A1Print(ss, 'TrackAndMatch_onEdit Running', 'B1');
  //var allForms = ss.getDataRange().getFormulas();

  forSpecificFormulaCells(ss, 'TrackAndMatch(', (r, c, loc) => {
    //A1Print(ss, 'TrackAndMatch_onEdit Running', 'B1');
    //account for the first time this ever runs:
    if (getGlobalVar('OLD_trackingData'+loc) == null) { setGlobalVar('OLD_trackingData'+loc, getGlobalVar('NEW_trackingData'+loc)); }
    const OLD_trackingData = JSON.parse(getGlobalVar('OLD_trackingData'+loc));
    const NEW_trackingData = JSON.parse(getGlobalVar('NEW_trackingData'+loc));
    const ColumnToTrack = getGlobalVar('trackingColumn'+loc);
    const ColumnToMatch = getGlobalVar('matchingColumn'+loc);
    const w = (letterToColumn(ColumnToMatch.slice(-1)) - letterToColumn(ColumnToMatch.charAt(0))) + 1;
    // only after I already got all vars
    setGlobalVar('OLD_trackingData'+loc, getGlobalVar('NEW_trackingData'+loc));

    //straight off the bat:
    //A1Print(ss, OLD_trackingData);
    //A1Print(ss, NEW_trackingData, 'B1');
    if (JSON.stringify(OLD_trackingData) == JSON.stringify(NEW_trackingData)) {
      //A1Print(ss, 'same!', 'C1');
      return;
    }

    //A1Print(ss,new Date().toTimeString());
    // --- get tracked and to-match current data on the sheet
    let columnData = ss.getRange(ColumnToTrack + ':' + ColumnToTrack).getValues();
    columnData = columnData.slice(r-1);
    const NEW_trackingColumn = JSON.parse(JSON.stringify(columnData.slice(0, NEW_trackingData.length)));
    //A1Print(ss,NEW_trackingColumn);
    if (ColumnToMatch.includes(':')) {
      columnData = ss.getRange(ColumnToMatch).getValues();
    } else {
      columnData = ss.getRange(ColumnToMatch + ':' + ColumnToMatch).getValues();
    }
    columnData = columnData.slice(r-1);
    const OLD_matchingColumn = columnData.slice(0, OLD_trackingData.length);
    //A1Print(ss,OLD_matchingColumn, 'B1');

    // --- get the index of the data we're tracking
    let foundCol = -1;
    for (let i = 0; i < OLD_trackingData[0].length; i++) {
      if (columnToLetter(c + i) == ColumnToTrack) {
        foundCol = i;
        break;
      }
    }
    if (foundCol < 0) {
      return;
      //aka die.
      //their column isn't within the data we have
    }

    // --- get old tracked data
    let oldthing = Array();
    for (let i = 0; i < OLD_trackingData.length; i++) {
      oldthing.push(OLD_trackingData[i][foundCol]);
    }
    const CURRENT_trackingColumn = oldthing;
    
    
    
    
    //A1Print(ss, 'still running');

    //A1Print(ss, OLD_matchingColumn);
    // --- create empty array
    let whereTheyGo = Array(NEW_trackingColumn.length).fill( Array(w).fill("") );
    //A1Print(ss, OLD_matchingColumn, 'B1');
    //A1Print(ss, whereTheyGo, 'B1');

    
    // --- "pair" data points together
    for (let i = 0; i < OLD_matchingColumn.length; i++) {
      const matcher = OLD_matchingColumn[i];
      const friend = CURRENT_trackingColumn[i];

      // now just find where the friend went :)
      if(friend == NEW_trackingColumn[i]) {
        whereTheyGo[i] = matcher;
        continue;
      }
      if(friend == NEW_trackingColumn[i+1]) {
        whereTheyGo[i+1] = matcher;
        continue;
      }
      if(friend == NEW_trackingColumn[i-1]) {
        whereTheyGo[i-1] = matcher;
        continue;
      }
      for (let i = 0; i < NEW_trackingColumn.length; i++) {
        if (friend == NEW_trackingColumn[i]) {
          whereTheyGo[i] = matcher;
        }
      }
    }
    //A1Print(ss,whereTheyGo, 'B1');


    // --- convert to 2D array
    /*let output = Array();
    for (let i = 0; i < whereTheyGo.length; i++) {
      output.push([whereTheyGo[i]]);
    }*/
    output = whereTheyGo;

    // --- apply new array
    //A1Print(ss, width);
    //A1Print(ss, output, 'B1');
    ss.getRange(r,letterToColumn(ColumnToMatch.charAt(0)), whereTheyGo.length,w).setValues(output);

    //ss.getRange(loc).setNote('you have my function!');
    
  });
}
