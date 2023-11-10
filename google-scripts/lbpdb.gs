function createLBPAppA() {
  lbpLabExport()
  formatLBPAppA()
  exportPDF()
}

function lbpLabExport() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()

  let lbpDB = ss.getActiveSheet();
  let dbLR = lbpDB.getLastRow();
  let dbLC = lbpDB.getLastColumn();
  let labRes = ss.getSheets()[ss.getSheets().length - 1];
  let labLR = labRes.getLastRow();
  let labLC = labRes.getLastColumn();

  let lbpHeaders = lbpDB.getRange(1, 1, 1, dbLC).getValues();
  for (let i of lbpHeaders[0]) {
    if (i === 'roomEquivalent') var room = lbpDB.getRange(2, lbpHeaders[0].indexOf(i) + 1, dbLR - 1).getValues();
    if (i === 'Building Component') var compo = lbpDB.getRange(2, lbpHeaders[0].indexOf(i) + 1, dbLR - 1).getValues();
    if (i === 'Paint Color') var color = lbpDB.getRange(2, lbpHeaders[0].indexOf(i) + 1, dbLR - 1).getValues();
    if (i === 'Substrate') var sub = lbpDB.getRange(2, lbpHeaders[0].indexOf(i) + 1, dbLR - 1).getValues();
    if (i === 'Sample ID' || i === 'SampleID') var lbpSIDs = lbpDB.getRange(2, lbpHeaders[0].indexOf(i) + 1, dbLR - 1).getValues();
  }
  
  // create sample descriptions using color + sub + compo
  // push into new array
  let sampleDesc = []
  for (let i = 0; i < color.length; i++) {
    sampleDesc.push([`${color[i]} ${sub[i]} ${compo[i]}`])
  }

  // get headers from the lab report, then iterate to find sid, result, and result unit columns
  // get create arrays for each column
  let labHeaders = labRes.getRange(1, 1, 1, labLC).getValues()
  for (let i of labHeaders[0]) {
    if (i === 'SampleID') var labSIDs = labRes.getRange(2, labHeaders[0].indexOf(i) + 1, labLR - 1).getValues();
    if (i === 'Result') var res = labRes.getRange(2, labHeaders[0].indexOf(i) + 1, labLR - 1).getValues();
    if (i === 'ResultUnits') var units = labRes.getRange(2, labHeaders[0].indexOf(i) + 1, labLR - 1).getValues();
  }

  // combine res and units
  // push into new array
  let totLead = []
  for (let i = 0; i < res.length; i++) {
    totLead.push([`${res[i]} ${units[i]}`])
  } 

  let leadAppA = [];
  let lbp = [];
  let lcp = [];
  let nd = [];
  for (let x of labSIDs) {
    if (units[labSIDs.indexOf(x)][0] === 'wt%') { // sort sids based on lead reporting rules
      if (res[labSIDs.indexOf(x)] >= 0.5) { // based on the reporting units
        lbp.push(x[0]);
      }  else if (res[labSIDs.indexOf(x)] > 0) {
        lcp.push(x[0]);
      } else nd.push(x[0])
    }
    if (units[labSIDs.indexOf(x)][0] === 'mg/cm2') {
      if (res[labSIDs.indexOf(x)] >= 1) {
        lbp.push(x[0]);
      }  else if (res[labSIDs.indexOf(x)] > 0) {
        lcp.push(x[0]);
      } else nd.push(x[0])
    }
    for (let y of lbpSIDs) {
      if (x[0] === y[0]) {
        leadAppA.push([
          `${lbpSIDs[lbpSIDs.indexOf(y)]}`, `${sampleDesc[lbpSIDs.indexOf(y)]}`, `${room[lbpSIDs.indexOf(y)]}`, `${totLead[labSIDs.indexOf(x)]}`
        ])
      }
    }
  }

  // initialize new app a sheet
  let lbpAppA = ss.insertSheet("LBP App A");
  lbpAppA.getRange(1, 1, leadAppA.length, 4).setValues(leadAppA);
  let lbpAALR = lbpAppA.getLastRow();
  let lbpAASIDs = lbpAppA.getRange(1, 1, lbpAALR).getValues();
  let aaSIDs = [];
  
  let columnNames = [['Sample ID', 'Sample Description', 'Sample Location', 'Total Lead']]
  lbpAppA.insertRowBefore(1)
  lbpAppA.getRange('A1:D1').setValues(columnNames);
  // this is step to make the rest of the steps easier
  for (let i of lbpAASIDs) {
    aaSIDs.push(i[0])
  }
  
  // find the intersecting values between lbp and sids
  let lbpInt = lbp.filter(sid => aaSIDs.includes(sid));
  for (let i of lbpInt) {
    let rowNum = aaSIDs.indexOf(i);
    lbpAppA.getRange(rowNum + 2, 1, 1, 4).setFontColor('red');
    lbpAppA.getRange(rowNum + 2, 1, 1, 4).setFontWeight('bold');
  }

  // find the intersecting values between lcp and sids
  // find index to find row
  let lcpInt = lcp.filter(sid => aaSIDs.includes(sid));
  for (let i of lcpInt) {
    let rowNum = aaSIDs.indexOf(i);
    lbpAppA.getRange(rowNum + 2, 1, 1, 4).setFontColor('orange');
  }
}

// Add header, format column widths and text alignment
function formatLBPAppA() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let lbpAppA = ss.getSheetByName('LBP App A');
  let lbpAALC = lbpAppA.getLastColumn(); 
  let lbpAALR = lbpAppA.getLastRow();  

  let cover = ss.getSheets()[0];
  let build = cover.getRange(7, 2).getValue();
  let survDate = cover.getRange(9, 2).getValue();

  // format resize columns and set aligment
  for (let i = 1; i <= lbpAALC; i++) {
    if (i === 4) {
      lbpAppA.getRange(2, i, lbpAALR).setHorizontalAlignment("right");
      lbpAppA.getRange(2, i, lbpAALR).setVerticalAlignment("middle");
    } else {
      lbpAppA.getRange(2, i, lbpAALR).setHorizontalAlignment("center");
      lbpAppA.getRange(2, i, lbpAALR).setVerticalAlignment("middle");
    }
  }

  // entire sheet format
  let sheet = lbpAppA.getRange(1, 1, lbpAALR, lbpAALC);
  sheet.setBorder(true, true, true, true, true, true);
  sheet.setFontFamily("Arial");
  sheet.setFontSize(11);
  lbpAppA.setColumnWidth(1, 225); lbpAppA.setColumnWidth(2, 500); lbpAppA.setColumnWidth(3, 300); lbpAppA.setColumnWidth(4, 150);

  let headers = [['Laboratory Lead Results', '', '', ''],
  [`${build}`, '', '', 'Lead Survey Report'],
  [`Singapore Area Coordinator`, '', '', `Survey Date: ${survDate}`],
  ['Sembawang, Singapore', '', '', ''],
  ['', '', '', '']];
  lbpAppA.insertRowsBefore(1, 5)
  lbpAppA.getRange("A1:D5").setValues(headers);

  lbpAppA.getRange(1, 1, 1, lbpAALC).mergeAcross().setHorizontalAlignment("center").setFontSize(14);
  lbpAppA.getRange(2, 1, 3).setHorizontalAlignment("left").setFontSize(11).setWrap(true).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  lbpAppA.getRange(2, lbpAALC, 3).setHorizontalAlignment("right").setFontSize(11).setWrap(true).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

  let dataHeader = lbpAppA.getRange(6, 1, 1, lbpAALC);
  dataHeader.setHorizontalAlignment("center")
  dataHeader.setBackground("Gainsboro");
  dataHeader.setFontWeight("bold");

  lbpAppA.setFrozenRows(6);
}

// exports app a as pdf, prompts user with download when finished
function exportPDF() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let lbpAppA = ss.getSheetByName('LBP App A');
  let ssURL = ss.getUrl().slice(0,-5); // https://docs.google.com/spreadsheets/d/<KEY>
  let gid = lbpAppA.getSheetId(); // gid
  let expURL = `${ssURL}/export?format=xlsx&gid=${gid}`; // https://docs.google.com/spreadsheets/d/<KEY>/export?format=xlsx&gid=<GID>
  var htmlOutput = HtmlService
    .createHtmlOutput(`<a href="${expURL}" >Click to download</a>`)
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download below');
}
