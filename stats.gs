function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuItems = [
    {name: 'Update Summary', functionName: 'updateSummary'},
    {name: 'Aesthetics', functionName: 'aesthetics'}
  ];
  ss.addMenu('Statistics', menuItems);
};

function updateSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheetName = "Output";
  var summarySheet = ss.getSheetByName(summarySheetName);
  const toaverage = [0,0,0,1,1,1,1,1,1,1,0,0,0,1,0];
  var rows = ["C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q"];
  const summaryrows = ["B","C","D","E","F","G","H","I","J","K","L","M","N","O"];
  const categories = ["Total Games", "Total Questions","Total Points","Average PPG","X-Risk/Energy","Math","Chem","Earth/Space","Bio","Physics", "Correct Interrupts", "Total Negs", "Conv %", "Avg Buzzes", "Avg Buzz % Correct"];
  
  summarySheet.clear();

  summarySheet.getRange(1, 2, 1, categories.length).setValues([categories]);

  // Collect all test names
  var sheets = ss.getSheets();
  var testnames = [];
  // More efficient name collection, less foolproof 
  // var rostersheet = ss.getSheetByName("Rosters");
  // var rosterData = rostersheet.getRange(1, 1, rostersheet.getLastRow(), 1).getValues();
  // testnames = rosterData.flat().filter(Boolean);

  // Less efficient name collection, more foolproof
  sheets.forEach(sheet => {
    var sheetName = sheet.getName()
    if (sheetName.indexOf("Practice") === 0) {
      names = sheet.getRange("P3:P12").getValues();
      names.forEach(name => {
        if (!(testnames.includes(name[0]))) {
          testnames.push(name[0])
        }
      })
    }
  });

  testnames = testnames.flat().filter(n => n).sort();
  var totaldata = {};
  testnames.forEach(name => {
    totaldata[name] = new Array(rows.length).fill(0);
  });

  // Iterate through each sheet and update totaldata
  sheets.forEach(sheet => {
    var sheetName = sheet.getName();
    if (sheetName.indexOf("Practice") === 0) {
      // var inputRange = sheet.getRange("B68:P78").getValues();
      // inputRange.forEach(row => {
      //   var name = row[0];
      //   if (totaldata.hasOwnProperty(name)) {
      //     for (var j = 0; j < totaldata[name].length; j++) {
      //       totaldata[name][j] += row[j + 1]; // Skip the name column
      //     }
      //   }
      // });
      var inputRange = sheet.getRange("B46:R55").getValues();
      var totalQs = sheet.getRange("Q44").getValue();
      inputRange.forEach(row => {
        var name = row[0];
        if (totaldata.hasOwnProperty(name)) {
          // Total Games
          totaldata[name][0] += Math.round((row[15] / totalQs) * 1e3) / 1e3;
          // Total Questions
          totaldata[name][1] += row[15];
          // Total Points
          totaldata[name][2] += row[13];
          // PPG
          totaldata[name][3] += row[13];
          // Categories
          for (var j = 2; j < 8; j++) {
            totaldata[name][j + 2] += row[j];
          }
          // Correct Interrupts
          totaldata[name][10] += row[16];
          // Negs
          totaldata[name][11] += row[11];
          // Conversion %
          totaldata[name][12] += row[14];
          // Avg Buzzes
          totaldata[name][13] += row[8];
          // Buzz % Correct
          totaldata[name][14] += row[9];
          //
        }
      })
    }
  });
  // Update summary sheet
  var rowIndex = 2;
  testnames.forEach(name => {
    var dataRow = [name];
    for (var j = 0; j < rows.length; j++) {
      dataRow.push(toaverage[j] === 1 ? (totaldata[name][j] / totaldata[name][0]).toFixed(3) : totaldata[name][j]);
    }
    if (dataRow[15] != 0) {
      dataRow[13] = ((dataRow[13] / dataRow[15]) * 100).toFixed(3);
    } else {dataRow[13] = 0;}
    dataRow[dataRow.length - 1] = (dataRow[dataRow.length - 1] / (dataRow[1] * dataRow[dataRow.length - 2])) * 100;
    if (dataRow[dataRow.length - 2] == 0) {dataRow[dataRow.length - 1] = 0;} else {
      dataRow[dataRow.length - 1] = dataRow[dataRow.length - 1].toFixed(3);
    }
    if (dataRow[1] == 0) {
      for (var i = 2; i < dataRow.length; i++) {dataRow[i] = 0;}
    }
    summarySheet.getRange(rowIndex, 1, 1, dataRow.length).setValues([dataRow]);
    rowIndex++;
  });

  // for (var i = 5; i < 12; i++) {
  //   outputSortedStats(totaldata, toaverage, summaryrows, i);
  // }
  // outputSortedStats(totaldata, toaverage, summaryrows, 12);
}



function readSummary(categoryvalue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("Output");
  var ppgstats = {};
  var row = 2;
  const catvalue = Number(categoryvalue);
  while (!ws.getRange(row, 1).isBlank()) {
    tmpvalue = Number(parseFloat(ws.getRange(row, catvalue).getValue()).toFixed(3));
    if (ppgstats[tmpvalue] == null) {
      ppgstats[tmpvalue] = [];
    }
    ppgstats[tmpvalue][ppgstats[tmpvalue].length] = ws.getRange(row, 1).getValue();
    row++;
  }
  return ppgstats;
}

function outputSortedStats(totaldata, toaverage, rows, categoryvalue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ppgstats = readSummary(categoryvalue)
  const outputsheetnames = ["PPG Output", "Energy Output", "Math Output", "Chemistry Output", "Earth/Space Output", "Bio Output", "Physics Output", "Negs Output"]
  const categories = ["Total Games", "Total Questions","Total Points","Average PPG","Energy","Math","Chem","Earth/Space","Bio","Physics","Total Negs", "Conv %", "Avg Buzzes", "Avg Buzz % Correct"];
  var outputSummarySheet = ss.getSheetByName(outputsheetnames[categoryvalue - 5]);

  var ppgkeys = Object.keys(ppgstats);
  ppgkeys.sort(function(a, b){return a-b});

  outputSummarySheet.clear();
  outputSummarySheet.getRange(1, 2, 1, categories.length).setValues([categories]);
  var rowIndex = 2;
  var sortednames = [];
  ppgkeys.forEach(ppgkey => {
    ppgstats[ppgkey].reverse();
    sortednames = sortednames.concat(ppgstats[ppgkey]);
  });
  sortednames = sortednames.reverse();
  sortednames.forEach(name => {
    var dataRow = [name];
    for (var j = 0; j < rows.length; j++) {
      dataRow.push(toaverage[j] == 1 ? (totaldata[name][j] / totaldata[name][0]).toFixed(3) : (totaldata[name][j]));
    }
    if (dataRow[14] != 0) {
      dataRow[12] = ((dataRow[12] / dataRow[14]) * 100).toFixed(3);
    } else {dataRow[12] = 0;}
    dataRow[dataRow.length - 1] = (dataRow[dataRow.length - 1] / (dataRow[1] * dataRow[dataRow.length - 2])) * 100;
    if (dataRow[dataRow.length - 2] == 0) {dataRow[dataRow.length - 1] = 0;} else {
      dataRow[dataRow.length - 1] = dataRow[dataRow.length - 1].toFixed(3);
    }
    if (dataRow[1] == 0) {
      for (var i = 2; i < dataRow.length; i++) {dataRow[i] = 0;}
    }
    outputSummarySheet.getRange(rowIndex, 1, 1, dataRow.length).setValues([dataRow]);
    rowIndex++;
  })
};

function aesthetics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(sheet => {
    if (sheet.getName().indexOf("RR") === 0) {
      sheet.getRange(1, 1, 3, 1).merge().setVerticalAlignment("top");
      sheet.getRange(4, 1, 22, 1).merge().setVerticalAlignment("top");
      sheet.getRange(4, 16, 22, 1).merge().setVerticalAlignment("top");
      sheet.getRange(26, 1, 3, 1).merge().setVerticalAlignment("top");
      sheet.getRange(29, 1, 9, 1).merge().setVerticalAlignment("top");
      sheet.getRange(40, 1, 3, 1).merge().setVerticalAlignment("top");
      sheet.getRange(43, 1, 39, 1).merge().setVerticalAlignment("top");
      sheet.getRange(26, 7, 4, 2).merge().setVerticalAlignment("top").setHorizontalAlignment("center");
      sheet.getRange(26, 9, 4, 7).merge().setVerticalAlignment("top");
      sheet.setColumnWidths(4, 12, 50);
      sheet.setColumnWidth(1, 100);
      sheet.setColumnWidth(3, 100);
      sheet.setColumnWidth(2, 50);
      sheet.setColumnWidth(9, 75);
      sheet.setColumnWidth(15, 75);
    }
  })
}
