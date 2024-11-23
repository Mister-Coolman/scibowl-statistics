function convertSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("Testing");
  const maxColumn = ws.getLastColumn();
  const maxRow = ws.getLastRow();
  const allHeaders = ws.getRange(1, 1, 1, maxColumn).getValues()[0];
  const allData = ws.getRange(2, 1, maxRow - 1, maxColumn).getValues();
  var teamA = [];
  var teamB = [];
  console.log(allHeaders)
  console.log(allData)
  var team = false;
  allHeaders.forEach(header => {
    console.log(header)
    if (header == "number" || header == "category") {
      return;
    }
    if (String(header).includes("Bonus")) {
      team = true;
      return;
    }
    if (!team) {
      teamA.push(header);
    } else {
      teamB.push(header);
    }
  });
  console.log(teamA, teamB)
  sheetName = insertNames(teamA, teamB);
  var ns = ss.getSheetByName(sheetName);
  var rowStart = 4;
  allData.forEach(row => {
    ns.getRange(rowStart, 3).setValue(String(row[1]));
    rowStart++;
  });
  for (var i = 0; i < allData.length; i++) {
    for (var j = 0; j < teamA.length; j++) {
      var dataPoint = convertData(allData[i][j + 2]);
      ns.getRange(i + 4, j + 4).setValue(dataPoint);
    }
    ns.getRange(i + 4, 9).setValue(convertBonus(allData[i][2 + teamA.length]))
    for (var j = 0; j < teamB.length; j++) {
      var dataPoint = convertData(allData[i][j + 3 + teamA.length]);
      ns.getRange(i + 4, j + 10).setValue(dataPoint);
    }
    ns.getRange(i + 4, 15).setValue(convertBonus(allData[i][allData[i].length - 1]));
  }
  ns.setColumnWidths(3, 12, 50);
  ns.setColumnWidths(2, 1, 40);
  ns.setColumnWidth(9, 45);
  ns.setColumnWidth(15, 45);
}

function insertNames(teamA, teamB) {
  var date = new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var month = date.getMonth();
  month++;
  var sheetName = "Practice " + month + "/" + date.getDate() + " " + date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds();
  ss.insertSheet(sheetName);
  var ns = ss.getSheetByName(sheetName);
  var copyRange = ss.getSheetByName("Blank Scoresheet").getRange("A1:Q84");
  copyRange.copyTo(ns.getRange("A1:Q84")) 
  var rowStart = 3;
  var colStart = 4;
  teamA.forEach(name => {
    var number = rowStart - 2;
    ns.getRange("Q" + rowStart).setValue("A" + String(number));
    ns.getRange("P" + rowStart).setValue(name);
    ns.getRange(3, colStart).setValue("A" + String(number));
    rowStart++;
    colStart++;
  });
  colStart = 10;
  rowStart = 8;
  teamB.forEach(name => {
    var number = rowStart - 7;
    ns.getRange("Q" + rowStart).setValue("B" + String(number));
    ns.getRange("P" + rowStart).setValue(name);
    ns.getRange(3, colStart).setValue("B" + String(number));
    rowStart++;
    colStart++;
  });
  return sheetName;
}

function convertData(num) {
  if (num == "C") {
    return 4;
  }
  if (num == "P") {
    return -4;
  }
  if (num == "I") {
    return 0;
  }
  return "";
}

function convertBonus(num) {
  if (num == "C") {
    return 10;
  }
  if (num == "I") {
    return 0;
  }
  return "";
}
