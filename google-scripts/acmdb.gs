/* --- LAST UPDATED: 07/07/2023 --- */
let ss;

// setting global constants for row and col numbers so it 
// state global variables for the sample DB
let sampleDB;
let sdbLastRow;
let sdbLastCol;

// state global variables for lab data
let dataSet;
let dsLastRow;
let dsLastCol;

// state global variables for the dbc
let dbc;
let dbcLastRow;
let dbcLastCol;

// state global variables for AppA
let appa;
let appaLastRow;
let appaLastCol;

// function for when the user wants to create an app a
function createAppA() {
  let ui = SpreadsheetApp.getUi();
  ss = SpreadsheetApp.getActiveSpreadsheet();
  
  sampleDB = ss.getActiveSheet()
  sdbLastRow = sampleDB.getLastRow();
  sdbLastCol = sampleDB.getLastColumn();

  dataSet = ss.getSheetByName("DATA");
  dsLastRow = dataSet.getLastRow();
  dsLastCol = dataSet.getLastColumn();

  // check if dbc sheet exists, if not create it
  // and set its values
  if (!ss.getSheetByName("dbc")) {
    dbc = ss.insertSheet("dbc");
    dbcLastRow = dbc.getLastRow();
    dbcLastCol = dbc.getLastColumn();
  } else {
    dbc = ss.getSheetByName("dbc");
    dbcLastRow = dbc.getLastRow();
    dbcLastCol = dbc.getLastColumn();
  }

  // check if AppA sheet exists, if not create it
  // and set its values
  if (!ss.getSheetByName("AppA")) {
    appa = ss.insertSheet("AppA");
    appaLastRow = appa.getLastRow();
    appaLastCol = appa.getLastColumn();
  } else {
    appa = ss.getSheetByName("AppA");
    appaLastRow = appa.getLastRow();
    appaLastCol = appa.getLastColumn();
  }

  // create location variable, then prompt user for project location
  let location = '';
  let response = ui.prompt('Project Location', '(e.g. Japan, USA, etc. - Can be left blank if not Japan)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) location = response.getResponseText().toUpperCase();
  else return;

  // call app a creation script with prompted user location
  try {
    createDBC();
    formatLabData();
    createLabExport();
    handleAsbestos();
    createCopyAppA();
    formatExport();
    colorAsbestos(location);
    handleAssumed(location);
    appafinalFormat();
    updateCoversheet();
    deleteExtraSheets();
    exportAppA()
  } catch(err) {
    var htmlOutput = HtmlService
      .createHtmlOutput(`<a>${err}</a>`)
      .setWidth(300) //optional
      .setHeight(100); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Script Error Output');
  }
}

// function for when the user wants to update the sample db
function updateDB() {
  let ui = SpreadsheetApp.getUi();
  ss = SpreadsheetApp.getActiveSpreadsheet();
  
  sampleDB = ss.getActiveSheet()
  sdbLastRow = sampleDB.getLastRow();
  sdbLastCol = sampleDB.getLastColumn();

  dataSet = ss.getSheetByName("DATA");
  dsLastRow = dataSet.getLastRow();
  dsLastCol = dataSet.getLastColumn();

  // check if dbc sheet exists, if not create it
  // and set its values
  if (!ss.getSheetByName("dbc")) {
    dbc = ss.insertSheet("dbc");
    dbcLastRow = dbc.getLastRow();
    dbcLastCol = dbc.getLastColumn();
  } else {
    dbc = ss.getSheetByName("dbc");
    dbcLastRow = dbc.getLastRow();
    dbcLastCol = dbc.getLastColumn();
  }

  // create location variable, then prompt user for project location
  let location = '';
  let response = ui.prompt('Project Location', '(e.g. Japan, USA, etc. - Can be left blank if not Japan)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) location = response.getResponseText().toUpperCase();
  else return;

  // call sample db updating script with prompted user location
  try {
    formatLabData();
    createDBC();
    addDBVals();
    formatDB(location);
    deleteExtraSheets();
  }
  catch(err) {
    var htmlOutput = HtmlService
      .createHtmlOutput(`<a>${err}</a>`)
      .setWidth(300) //optional
      .setHeight(100); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Script Error Output');
  }
}

function createDBC() {
  // get the header names from the sample DB
  // loop through it to find the needed columns for the dbc
  let headers = sampleDB.getRange(1, 1, 1, sdbLastCol).getValues();
  
  
  for (let i = 0; i < headers[0].length; i++) {
    if (headers[0][i].trim() == 'Homogeneous Material Number' || headers[0][i].trim() == 'Homogenous Material Number') var hoMatNum = sampleDB.getRange(1, i+1, sdbLastRow);
    if (headers[0][i].trim() == 'Sample ID') var sampID = sampleDB.getRange(1, i+1, sdbLastRow);
    if (headers[0][i].trim() == 'Material Type') var matType = sampleDB.getRange(1, i+1, sdbLastRow);
    if (headers[0][i].trim() == 'Material Description') var matDesc = sampleDB.getRange(1, i+1, sdbLastRow);
    if (headers[0][i].trim() == 'Friable') var friable = sampleDB.getRange(1, i+1, sdbLastRow);
    if (headers[0][i].trim() == 'Condition') var cond = sampleDB.getRange(1, i+1, sdbLastRow);
    if (headers[0][i].trim() == 'Room/Area 1' || headers[0][i].trim() == 'Room/Area') var rA = sampleDB.getRange(1, i+1, sdbLastRow);
  }
  
  // insert values from sample db to their respective ranges in the dbc
  if (hoMatNum !== undefined) {
    hoMatNum.copyValuesToRange(dbc, 1, 1, 1, sdbLastRow);
  } else {
    throw new Error('Can\'t find "Homogeneous Material Number" column header. Check it is written exactly as: "Homogeneous Material Number"')
  }
  if (sampID !== undefined) {
    sampID.copyValuesToRange(dbc, 2, 2, 1, sdbLastRow);
  } else {
    throw new Error('Can\'t find "Sample ID" column header. Check it is written exactly as: "Sample ID"')
  }
  if (matType !== undefined) {
    matType.copyValuesToRange(dbc, 3, 3, 1, sdbLastRow);
  } else {
    throw new Error('Can\'t find "Material Type" column header. Check it is written exactly as: "Material Type"')
  }
  if (matDesc !== undefined) {
    matDesc.copyValuesToRange(dbc, 4, 4, 1, sdbLastRow);
  } else {
    throw new Error('Can\'t find "Material Description" column header. Check it is written exactly as: "Material Description"')
  }
  if (friable !== undefined) {
    friable.copyValuesToRange(dbc, 5, 5, 1, sdbLastRow);
  } else {
    throw new Error('Can\'t find "Friable" column header. Check it is written exactly as: "Friable"')
  }
  if (cond !== undefined) {
    cond.copyValuesToRange(dbc, 6, 6, 1, sdbLastRow);
  } else {
    throw new Error('Can\'t find "Condition" column header. Check it is written exactly as: "Condition"')
  }
  if (rA !== undefined) {
    rA.copyValuesToRange(dbc, 7, 7, 1, sdbLastRow);
  } else {
    throw new Error('Can\'t find "Room/Area" column header. Check it is written exactly as: "Room/Area 1" or "Room/Area"')
  }

  // update the last row and column for the dbc
  dbcLastRow = dbc.getLastRow();
  dbcLastCol = dbc.getLastColumn();
}

function formatLabData() {
  // get the values for all the fibrous material percentages
  let cell = dataSet.getRange(2, 11, dsLastRow - 1).getValues();
  let fibGla = dataSet.getRange(2, 12, dsLastRow - 1).getValues();
  let synth = dataSet.getRange(2, 13, dsLastRow - 1).getValues();
  let talc = dataSet.getRange(2, 14, dsLastRow - 1).getValues();
  let wolla = dataSet.getRange(2, 15, dsLastRow - 1).getValues();

  // get the values for all layer numbers
  let layerNum = dataSet.getRange(2, 17, dsLastRow - 1).getValues();

  // get the values for all the percentage of asbestos
  let asb1Quant = dataSet.getRange(2, 21, dsLastRow - 1).getValues();
  let asb2Quant = dataSet.getRange(2, 23, dsLastRow - 1).getValues();
  let asb3Quant = dataSet.getRange(2, 25, dsLastRow - 1).getValues();
  
  // convert all NDs to 0s, all Trace to 0.99(as a place holder), and all other values to numbers instead of strings
  for (let i = 0; i < cell.length; i++) {
    if (cell[i][0] == 'Trace') {
      cell[i][0] = 0.99;
    }
    else cell[i][0] = Number(cell[i][0]);
  }
  for (let i = 0; i < fibGla.length; i++) {
    if (fibGla[i][0] == 'Trace') {
      fibGla[i][0] = 0.99;
    }
    else fibGla[i][0] = Number(fibGla[i][0]);
  }
  for (let i = 0; i < synth.length; i++) {
    if (synth[i][0] == 'Trace') {
      synth[i][0] = 0.99;
    }
    else synth[i][0] = Number(synth[i][0]);
  }
  for (let i = 0; i < talc.length; i++) {
    if (talc[i][0] == 'Trace') {
      talc[i][0] = 0.99;
    }
    else talc[i][0] = Number(talc[i][0]);
  }
  for (let i = 0; i < wolla.length; i++) {
    if (wolla[i][0] == 'Trace') {
      wolla[i][0] = 0.99;
    }
    else wolla[i][0] = Number(wolla[i][0]);
  }

  for (let i = 0; i < layerNum.length; i++) {
    layerNum[i][0] = Number(layerNum[i][0]);
  }

  for (let i = 0; i < asb1Quant.length; i++) {
    if (asb1Quant[i][0] == 'ND') {
      asb1Quant[i][0] = 0;
    } else if (asb1Quant[i][0] == 'Trace') {
      asb1Quant[i][0] = 0.99;
    }
    else asb1Quant[i][0] = Number(asb1Quant[i][0]);
  }
  for (let i = 0; i < asb2Quant.length; i++) {
    if (asb2Quant[i][0] == '') {
      asb2Quant[i][0] = 0;
    } else if (asb2Quant[i][0] == 'Trace') {
      asb2Quant[i][0] = 0.99;
    }
    else asb2Quant[i][0] = Number(asb2Quant[i][0]);
  }
  for (let i = 0; i < asb3Quant.length; i++) {
    if (asb3Quant[i][0] == '') {
      asb3Quant[i][0] = 0;
    } else if (asb3Quant[i][0] == 'Trace') {
      asb3Quant[i][0] = 0.99;
    }
    else asb3Quant[i][0] = Number(asb3Quant[i][0]);
  }
  // insert newly formatted data back to the lab data set
  dataSet.getRange(2, 11, dsLastRow - 1).setValues(cell);
  dataSet.getRange(2, 12, dsLastRow - 1).setValues(fibGla);
  dataSet.getRange(2, 13, dsLastRow - 1).setValues(synth);
  dataSet.getRange(2, 14, dsLastRow - 1).setValues(talc);
  dataSet.getRange(2, 15, dsLastRow - 1).setValues(wolla);

  dataSet.getRange(2, 21, dsLastRow - 1).setValues(asb1Quant);
  dataSet.getRange(2, 23, dsLastRow - 1).setValues(asb2Quant);
  dataSet.getRange(2, 25, dsLastRow - 1).setValues(asb3Quant);
}

// function to create the barebones, non-formatted app a
function createLabExport() {
  // get sample ids from the lab data
  let sampID = dataSet.getRange(2, 5, dsLastRow - 1).getValues();
  let homArea = [];
  // remove last letter from SampleID and store in homogeneous area array
  for (let i = 0; i < sampID.length; i++) {
    homArea[i] = [sampID[i].toString().slice(0, -1)];
  }

  // get values from the dbc
  let dbcSampID = dbc.getRange(2, 2, dbcLastRow - 1).getValues();
  let dbcMatType = dbc.getRange(2, 3, dbcLastRow - 1).getValues();
  let dbcMatDesc = dbc.getRange(2, 4, dbcLastRow - 1).getValues();
  let dbcFriable = dbc.getRange(2, 5, dbcLastRow - 1).getValues();
  let dbcCond = dbc.getRange(2, 6, dbcLastRow - 1).getValues();
  let dbcRA = dbc.getRange(2, 7, dbcLastRow - 1).getValues();

  // insert new columns and add headers
  let headers = [['Homogeneous Area', 'Material Type', 'Material Description', 'Friable', 'Condition', 'Sample ID', 'Sample Location', 'Layer (% of Combined Sample)', 'Asbestos %']];
  appa.getRange("A1:I1").setValues(headers);

  // initialize empty arrays to store the parsed data
  let matType = [];
  let matDesc = [];
  let friable = [];
  let cond = [];
  let rA = [];
  
  // match sample ids from lab data to sample ids from dbc
  // loop through data set sample ids
  for (let i = 0; i < sampID.length; i++) {
    for (let j = 0; j < dbcSampID.length; j++) {
      if (sampID[i][0] === dbcSampID[j][0]) {
        sampID[i] = [dbcSampID[j]];
        matType[i] = [dbcMatType[j]];
        matDesc[i] = [dbcMatDesc[j]];
        friable[i] = [dbcFriable[j]];
        cond[i] = [dbcCond[j]];
        rA[i] = [dbcRA[j]];
      }
    }
  }

  // get data needed for the "Layer (% of Combined Sample)" column
  // merge it all together into one array
  let layerNum = dataSet.getRange(2, 17, dsLastRow - 1).getValues();
  let layer = dataSet.getRange(2, 18, dsLastRow - 1).getValues();
  let pOfTotal = dataSet.getRange(2, 19, dsLastRow - 1).getValues();
  let layerP = [];
  for (let i = 0; i < layer.length; i++) {
    layerP[i] = [`${layerNum[i]}     ${layer[i]} (${pOfTotal[i]}%)`];
  }
  
  // setValues for added columns using the 2D arrays
  appa.getRange(2, 1, homArea.length).setValues(homArea);
  appa.getRange(2, 2, matType.length).setValues(matType);
  appa.getRange(2, 3, matDesc.length).setValues(matDesc);
  appa.getRange(2, 4, friable.length).setValues(friable);
  appa.getRange(2, 5, cond.length).setValues(cond);
  appa.getRange(2, 6, sampID.length).setValues(sampID);
  appa.getRange(2, 7, rA.length).setValues(rA);
  appa.getRange(2, 8, layerP.length).setValues(layerP);
  
  // update the last row and column for app a
  appaLastRow = appa.getLastRow();
  appaLastCol = appa.getLastColumn();
}

// function that reformats asbestos lab data
function handleAsbestos() {
  // get the values and names of all asbestos samples
  let asb1Name = dataSet.getRange(2, 20, dsLastRow - 1).getValues();
  let asb1Quant = dataSet.getRange(2, 21, dsLastRow - 1).getValues();
  let asb2Name = dataSet.getRange(2, 22, dsLastRow - 1).getValues();
  let asb2Quant = dataSet.getRange(2, 23, dsLastRow - 1).getValues();
  let asb3Name = dataSet.getRange(2, 24, dsLastRow - 1).getValues();
  let asb3Quant = dataSet.getRange(2, 25, dsLastRow - 1).getValues();
  
  // initialize empty arrays for each sample's total asbestos content
  // and asbestos quantities
  let totAsb = [];
  let a1Q = [];
  let a2Q = [];
  let a3Q = [];
  
  // loop through the asbestos data
  // change all trace samples to "<1" strings
  for (let i = 0; i < asb1Quant.length; i++) {
    if (asb1Quant[i][0] < 1 && asb1Quant[i][0] > 0) {
      a1Q[i] = '<1';
    } else a1Q[i] = asb1Quant[i];
    if (asb2Quant[i][0] < 1 && asb2Quant[i][0] > 0) {
      a2Q[i] = '<1';
    } else a2Q[i] = asb2Quant[i];
    if (asb3Quant[i][0] < 1 && asb3Quant[i][0] > 0) {
      a3Q[i] = '<1';
    } else a3Q[i] = asb3Quant[i];
  }

  // merge all asbestos data together into one array
  // every loop check to see if there is more than one type of asbestos
  // if no asbestos return "ND"
  for (let i = 0; i < asb1Quant.length; i++) {
    if (asb1Quant[i][0] > 0 && asb2Name[i][0] == '') {
      totAsb[i] = [`${a1Q[i]}% ${asb1Name[i][0]}`];
    } else if (asb2Quant[i][0] > 0 && asb3Name[i][0] == '') {
      totAsb[i] = [`${a1Q[i]}% ${asb1Name[i][0]}, ${a2Q[i]}% ${asb2Name[i][0]}`];
    } else if (asb3Quant[i][0] > 0) {
      totAsb[i] = [`${a1Q[i]}% ${asb1Name[i][0]}, ${a2Q[i]}% ${asb2Name[i][0]}, ${a3Q[i]}% ${asb3Name[i][0]}`];
    } else totAsb[i] = ['ND'];
  }
  
  // set the values from the array to the "Asbestos %" column
  appa.getRange(2, 9, totAsb.length).setValues(totAsb);
}

function createCopyAppA() {
  let copy = ss.insertSheet("copy");
  let full = appa.getRange(1, 1, appaLastRow, appaLastCol);
  full.copyValuesToRange(copy, 1, 1, 1, 1);
}

function formatExport() {
  appa.activate();
  // get the range that contains Homogeneous Area,	Material Type,	Material Description,	Friable,	Condition
  // and range that contains Sample ID,	Sample Location
  let hAMtMdFC = appa.getRange(2, 1, 1, 5);
  let sISL = appa.getRange(2, 6, 1, 2);

  // sort based off material order
  sortOrder(appa, appaLastRow, appaLastCol);

  // merge cells for same homogeneous area #s and merge cells for same sample ID #s
  mergeSameCells(hAMtMdFC, appa, 0, 5);
  mergeSameCells(sISL, appa, 0, 2);
}

// function that handles coloring postive asbestos samples
function colorAsbestos(location) {
  let asbPerc = appa.getRange(2, 9, appaLastRow - 1).getValues();
  // if the user inputs japan as the project location color red and bold all positive samples
  // else bold and color non trace samples red and trace samples color orange
  if (location === 'JAPAN') {
    for (let i = 0; i < asbPerc.length; i++) {
      if (asbPerc[i][0] != 'ND') {
        let mergedAsb = [];
        mergedAsb = appa.getRange(i+2, 1, 1, 7).getMergedRanges();
        appa.getRange(i+2, 1, 1, appaLastCol).setFontColor("red");
        appa.getRange(i+2, 1, 1, appaLastCol).setFontWeight("bold");
        if (mergedAsb.length !== 0) {
          for (let j = 0; j < mergedAsb.length; j++) {
            mergedAsb[j].setFontColor("red");
            mergedAsb[j].setFontWeight("bold");
          }
        }  
      } 
    } 
  } else {
    for (let i = 0; i < asbPerc.length; i++) {
      if (asbPerc[i][0].includes("<")) {
        let mergedTrace = [];
        mergedTrace = appa.getRange(i+2, 1, 1, 9).getMergedRanges();
        appa.getRange(i+2, 1, 1, appaLastCol).setFontColor("orange")
        if (mergedTrace.length !== 0) {
          for (let j = 0; j < mergedTrace.length; j++) {
            if (mergedTrace[j].getFontWeight() !== "bold") mergedTrace[j].setFontColor("orange")
          }
        }
      } else if (asbPerc[i][0].includes("%") && !asbPerc[i][0].includes("<")) {
        let mergedAsb = []
        mergedAsb = appa.getRange(i+2, 1, 1, 7).getMergedRanges();
        appa.getRange(i+2, 1, 1, appaLastCol).setFontColor("red");
        appa.getRange(i+2, 1, 1, appaLastCol).setFontWeight("bold");
        if (mergedAsb.length !== 0) {
          for (let j = 0; j < mergedAsb.length; j++) {
            mergedAsb[j].setFontColor("red");
            mergedAsb[j].setFontWeight("bold");
          }
        }
      }
    }
  }
}

function handleAssumed(location) {
  let copy = ss.getSheetByName("copy");
  let copyLR = copy.getLastRow();

  let sdbHeaders = sampleDB.getRange(1, 1, 1, sdbLastCol).getValues();
  let dbcHoMatNum = dbc.getRange(2, 1, dbcLastRow - 1).getValues();
  let dbcSampID = dbc.getRange(2, 2, dbcLastRow - 1).getValues();
  let dbcMatType = dbc.getRange(2, 3, dbcLastRow - 1).getValues();
  let dbcMatDesc = dbc.getRange(2, 4, dbcLastRow - 1).getValues();
  let dbcFriable = dbc.getRange(2, 5, dbcLastRow - 1).getValues();
  let dbcCond = dbc.getRange(2, 6, dbcLastRow - 1).getValues();
  let dbcRA = dbc.getRange(2, 7, dbcLastRow - 1).getValues();


  for (let i = 0; i < sdbHeaders[0].length; i++) {
    if (sdbHeaders[0][i] == 'Homogeneous Material Number' || sdbHeaders[0][i] == 'Homogenous Material Number') {
      var hoMatNum = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
    }
    if (sdbHeaders[0][i] == 'Asbestos') {
      var hA = i+1;
    }
  }
  let assumed = []
  for (let i = 0; i < hoMatNum.length; i++) {
    let arr = hoMatNum[i][0].split('-');
    if (arr[arr.length - 2] === 'AC' || arr[arr.length - 2] === 'AF' || arr[arr.length - 2] === 'AW' || arr[arr.length - 2] === 'AT' || arr[arr.length - 2] === 'AM') {
      if (location === 'JAPAN') {
        assumed.push([`${dbcHoMatNum[i]}`, `${dbcMatType[i]}`, `${dbcMatDesc[i]}`, `${dbcFriable[i]}`, `${dbcCond[i]}`, `${dbcSampID[i]}`, `${dbcRA[i]}`, 'N/A', '>0.1% Assumed']);
      } else assumed.push([`${dbcHoMatNum[i]}`, `${dbcMatType[i]}`, `${dbcMatDesc[i]}`, `${dbcFriable[i]}`, `${dbcCond[i]}`, `${dbcSampID[i]}`, `${dbcRA[i]}`, 'N/A', '>1% Assumed']);
      sampleDB.getRange(i+2, hA).setValue('ASSUMED')
    }
  }
  if (assumed.length > 0) {
    appa.getRange(appaLastRow + 1, 1, assumed.length, 9).setValues(assumed);
    copy.getRange(copyLR + 1, 1, assumed.length, 9).setValues(assumed);
  }

  // update the last row and column for app a
  appaLastRow = appa.getLastRow();
  appaLastCol = appa.getLastColumn();

  let asbPerc = appa.getRange(2, 9, appaLastRow - 1).getValues()
  for (let i = 0; i < asbPerc.length; i++) {
    if (asbPerc[i][0] === '>0.1% Assumed' || asbPerc[i][0] === '>1% Assumed') {
      appa.getRange(i+2, 1, 1, appaLastCol).setFontColor("red");
      appa.getRange(i+2, 1, 1, appaLastCol).setHorizontalAlignment("center");
    }
  }
}

function appafinalFormat() {
  // get the coversheet (assuming it's the first sheet)
  // and get the coversheet headers
  let cover = ss.getSheets()[0];
  let coverCol1 = cover.getRange(1, 1, cover.getLastRow()).getValues();

  // loop through the coversheet headers
  // search and get the building number, facility description, location, and survey date from the coversheet
  for (let i = 0; i < coverCol1.length; i++) {
    if (coverCol1[i][0] == 'Building No.') var bNum = cover.getRange(i+1, 2).getValue();
    if (coverCol1[i][0].trim() == 'Facility Description\n (e.g. \'storage shed\')') var facDesc = cover.getRange(i+1, 2).getValue();
    if (coverCol1[i][0] == 'Location') var loc = cover.getRange(i+1, 2).getValue();
    if (coverCol1[i][0] == 'Survey Date') var survDate = cover.getRange(i+1, 2).getValue();
  }
  
  // entire sheet format
  let sheet = appa.getRange(1, 1, appaLastRow, appaLastCol);
  sheet.setBorder(true, true, true, true, true, true);
  sheet.setFontFamily("Arial");
  sheet.setFontSize(11);

  // format resize columns and set aligment
  appa.autoResizeColumns(1, appaLastCol);
  for (let i = 1; i <= appaLastCol; i++) {
    if (i === 8) {
      appa.getRange(2, i, appaLastRow).setHorizontalAlignment("left");
      appa.getRange(2, i, appaLastRow).setVerticalAlignment("middle");
    } else {
      appa.getRange(2, i, appaLastRow).setHorizontalAlignment("center");
      appa.getRange(2, i, appaLastRow).setVerticalAlignment("middle");
    }
  }
  appa.getRange(1, 3, appaLastRow).setWrap(true).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  appa.setColumnWidth(3, 250);

  // inserts rows needed for the header
  appa.insertRowsBefore(1, 5);

  // header format as new cells in the sheet
  let headers = [['Laboratory Asbestos Results', '', '', '', '', '', '', '', ''],
  [`Building ${bNum}, ${facDesc}`, '', '', '', '', '', '', '', 'Asbestos Survey Report'],
  [`${loc}`, '', '', '', '', '', '', '', `Survey Date: ${survDate}`],
  ['', '', '', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '']];
  appa.getRange("A1:I5").setValues(headers);

  // header format
  appa.getRange(1, 1, 1, appaLastCol).mergeAcross().setHorizontalAlignment("center").setFontSize(14);
  appa.getRange(2, 1, 3).setHorizontalAlignment("left").setFontSize(11).setWrap(true).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  appa.getRange(2, appaLastCol, 3).setHorizontalAlignment("right").setFontSize(11).setWrap(true).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

  // data header format
  let dataHeader = appa.getRange(6, 1, 1, appaLastCol);
  dataHeader.setHorizontalAlignment("center")
  dataHeader.setBackground("Gainsboro");
  dataHeader.setFontWeight("bold");

  appa.setFrozenRows(6);
}

function updateCoversheet() {
  let cover = ss.getSheets()[0];
  let copy = ss.getSheetByName("copy");
  let copyLC = copy.getLastColumn();
  let copyLR = copy.getLastRow();

  sortOrder(copy, copyLR, copyLC);
  
  let vals = copy.getRange(2, 9, copyLR - 1).getValues();
  let f = copy.getRange(2, 4, copyLR - 1).getValues();
  let c = copy.getRange(2, 5, copyLR - 1).getValues();
  let l = copy.getRange(2, 8, copyLR - 1).getValues();
  let mt = copy.getRange(2, 3, copyLR - 1).getValues();
  let lo = copy.getRange(2, 7, copyLR - 1).getValues();
  let sids = copy.getRange(2, 6, copyLR - 1).getValues();
  let has = copy.getRange(2, 1, copyLR - 1).getValues();
  let asbPerc = [];
  let pos = [];
  let posHA = [];
  let assuA = '';

  for (let i = 0; i < copyLR - 1; i++) {
    if (vals[i][0] != 'ND') {
      asbPerc[i] = vals[i][0].split(", ");
    } else asbPerc[i] = ['']
  }
  for (let i = 0; i < asbPerc.length; i++) {
    for (let j = 0; j < asbPerc[i].length; j++) {
      if (asbPerc[i][0] != '' && asbPerc[i][0] !== '>1% Assumed' && asbPerc[i][0] !== '>0.1% Assumed') {
        let lay = l[i][0].substring(1).trim().split('(');
        pos.push([vals[i][0], lay[0].trim(), has[i][0], f[i][0], c[i][0], lo[i][0]])
        posHA.push(`${has[i][0]} ${vals[i][0]} ${lo[i][0]} ${lay[0].trim()}`)
      } else if (asbPerc[i][0] === '>1% Assumed' || asbPerc[i][0] === '>0.1% Assumed') assuA += `${mt[i][0]} (HA ${has[i][0]}, ${f[i][0]}, ${c[i][0]} condition) observed throughout.\n\n`;
    }
  }
  let uniqPos = [];
  let pAsb = [];
  posHA.forEach((p) => {
    if (!uniqPos.includes(p)) {
      uniqPos.push(p);
      let num = posHA.indexOf(p)
      pAsb.push(`${pos[num][0]} asbestos was identified in the ${pos[num][1]} (HA ${pos[num][2]}, ${pos[num][3]}, ${pos[num][4]} condition) collected from "${pos[num][5]}".`)
    }
  });
  cover.getRange(37, 2).setValue(pAsb.join('\n\n'))
  cover.getRange(38, 2).setValue(assuA.trim())
  
  let numAss = 0;
  for (let i = 0; i < has.length; i++) {
    has[i] = has[i][0];
  }
  let totHA = has.filter(onlyUnique);
  for (let i = 0; i < sids.length; i++) {
    sids[i] = sids[i][0];
  }
  let totSID = sids.filter(onlyUnique);
  
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0] === 'Assumed') numAss++;
  }
  cover.getRange(21, 3).setValue("total HAs surveyed:");
  cover.getRange(21, 4).setValue(totHA.length - numAss);
  cover.getRange(26, 3).setValue("total samples taken:");
  cover.getRange(26, 4).setValue(totSID.length - numAss);
}

function exportAppA() {
  let ssURL = ss.getUrl().slice(0,-5); // https://docs.google.com/spreadsheets/d/<KEY>
  let gid = appa.getSheetId(); // gid
  let expURL = `${ssURL}/export?format=xlsx&gid=${gid}`; // https://docs.google.com/spreadsheets/d/<KEY>/export?format=xlsx&gid=<GID>
  var htmlOutput = HtmlService
    .createHtmlOutput(`<a href="${expURL}" >Click to download</a>`)
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download below');
}

function addDBVals() {
  // get layer numbers, dates, and report numbers from the lab data
  let layerNum = dataSet.getRange(2, 17, dsLastRow - 1).getValues();
  let dA = dataSet.getRange(2, 6, dsLastRow - 1).getValues();
  let rN = dataSet.getRange(2, 1, dsLastRow - 1).getValues();
  
  // get sample ids and asbestos data from the lab data
  let whereA = dataSet.getRange(2, 3, dsLastRow - 1).getValues();
  let sampID = dataSet.getRange(2, 5, dsLastRow - 1).getValues();
  let asb1Name = dataSet.getRange(2, 20, dsLastRow - 1).getValues();
  let asb1Quant = dataSet.getRange(2, 21, dsLastRow - 1).getValues();
  let asb2Name = dataSet.getRange(2, 22, dsLastRow - 1).getValues();
  let asb2Quant = dataSet.getRange(2, 23, dsLastRow - 1).getValues();
  let asb3Name = dataSet.getRange(2, 24, dsLastRow - 1).getValues();
  let asb3Quant = dataSet.getRange(2, 25, dsLastRow - 1).getValues();

  // get other fibrous data from the lab data
  let cell = dataSet.getRange(2, 11, dsLastRow - 1).getValues();
  let fibGla = dataSet.getRange(2, 12, dsLastRow - 1).getValues();
  let synth = dataSet.getRange(2, 13, dsLastRow - 1).getValues();
  let talc = dataSet.getRange(2, 14, dsLastRow - 1).getValues();
  let wolla = dataSet.getRange(2, 15, dsLastRow - 1).getValues();

  // get sample ids from the dbc to compare to lab data
  let dbcHoMatNum = dbc.getRange(2, 1, dbcLastRow - 1).getValues();
  let dbcSampID = dbc.getRange(2, 2, dbcLastRow - 1).getValues();

  // get the headers, getvalues based off the header, and set col numbers 
  let headers = sampleDB.getRange(1, 1, 1, sdbLastCol).getValues()
  for (let i = 0; i < headers[0].length; i++) {
    if (headers[0][i] === 'Report Date') var repDateCol = i+1;
    if (headers[0][i] === 'Matrix / Layers' || headers[0][i] === 'Matrix/Layers') {
      var lay = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var layCol = i+1;
    }
    if (headers[0][i] === 'Laboratory Analysis Method 1') {
      var labAM1 = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var labAM1Col = i+1;
    }
    if (headers[0][i] === 'Laboratory Analysis Method 2') {
      var labAM2 = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var labAM2Col = i+1;
    }
    if (headers[0][i] === 'Laboratory Analysis Method 3') {
      var labAM3 = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var labAM3Col = i+1;
    }
    if (headers[0][i] === 'Laboratory Name') {
      var labName = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var labNameCol = i+1;
    }
    if (headers[0][i] === 'Laboratory Location City') {
      var labLoc = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var labLocCol = i+1;
    }
    if (headers[0][i] === 'Laboratory Location ST') {
      var labLocST = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var labLocSTCol = i+1;
    }
    if (headers[0][i] === 'Laboratory NVLAP Accreditation Number') {
      var labAN = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var labANCol = i+1;
    }
    if (headers[0][i] === 'Laboratory Report Number') {
      var labRepNum = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var lrnCol = i+1;
    }
    if (headers[0][i] === 'Laboratory Report Date') {
      var labRepDate = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var lrdCol = i+1;
    }
    if (headers[0][i] === 'Asbestos') {
      var hA = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var hACol = i+1;
    }
    if (headers[0][i] === 'Asbestos Content %') {
      var totQ = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var totQCol = i+1;
    }
    if (headers[0][i] === 'Asbestos Type 1') {
      var a1N = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var a1NCol = i+1;
    }
    if (headers[0][i] === 'Asbestos Type 2') {
      var a2N = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var a2NCol = i+1;
    }
    if (headers[0][i] === 'Asbestos Type 3') {
      var a3N = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var a3NCol = i+1;
    }
    if (headers[0][i] === 'Composite Non Asbestos Content %'){
      var nonACont = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var nonAContCol = i+1;
    }
    if (headers[0][i] === 'Composite Non Asbestos Content Type') {
      var nAType = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var nATypeCol = i+1;
    }
    if (headers[0][i] === 'Recommended Action') {
      var recAct = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var recActCol = i+1;
    }
    if (headers[0][i] === 'Additional Fields') {
      var addFields = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
      var addFieldsCol = i+1;
    }
  }

  let ce = [], fG = [], sy = [], ta = [], wo = [];
  let a1Q = [], a2Q = [], a3Q = [];

  // initialize the arrays
  for (let i = 0; i < dbcSampID.length; i++) {
    a1Q[i] = 0; a2Q[i] = 0; a3Q[i] = 0;
  }
  
  // loop through lab data sample ids and dbc sample ids
  // if there's a match add respective lab data
  for (let i = 0; i < sampID.length; i++) {
    for (let j = 0; j < dbcSampID.length; j++) {
      if (sampID[i][0] === dbcSampID[j][0]) {
        lay[j] = layerNum[i];
        labAM1[j] = ['PLM'];
        labAM2[j] = ['N/A'];
        labAM3[j] = ['N/A'];
        labName[j] = ['SGS Forensic Laboratories / 790202894'];
        labRepDate[j] = dA[i];
        labRepNum[j] = rN[i];
        ce[j] = cell[i];
        fG[j] = fibGla[i];
        sy[j] = synth[i];
        ta[j] = talc[i];
        wo[j] = wolla[i];
        addFields[j] = ['N/A'];

        // add lab info based on where the data was analyzed 
        if (whereA[i][0] === 'Los Angeles') {
          labLoc[j] = ['Carson'];
          labLocST[j] = ['CA'];
          labAN[j] = ['101459-1'];
        } else if (whereA[i][0] === 'Hayward') {
          labLoc[j] = ['Hayward'];
          labLocST[j] = ['CA'];
          labAN[j] = ['101459-1'];
        } else if (whereA[i][0] === 'Las Vegas') {
          labLoc[j] = ['Las Vegas'];
          labLocST[j] = ['NV'];
          labAN[j] = ['200908-0'];
        } else if (whereA[i][0] === 'Chicago') {
          labLoc[j] = ['Chicago'];
          labLocST[j] = ['IL'];
          labAN[j] = ['101732-0'];
        }

        // breakdown asbestos data into separate arrays
        if (asb1Quant[i][0] >= a1Q[j]) {
          a1Q[j] = asb1Quant[i][0];
          a1N[j] = asb1Name[i];
          // if the sample id is positive store the result in the 
          // has asbestos array
          if (asb1Quant[i][0] > 0) {
            hA[j] = ['YES'];
            recAct[j] = ['O&M'];
          } else {
            hA[j] = ['NO'];
            recAct[j] = ['N/A'];
          }  
        }
        if (asb2Quant[i][0] >= a2Q[j]) {
          a2Q[j] = asb2Quant[i][0];
          a2N[j] = asb2Name[i];
        }
        if (asb3Quant[i][0] >= a3Q[j]) {
          a3Q[j] = asb3Quant[i][0];
          a3N[j] = asb3Name[i];
        }
      }
    }
  }

  // change all trace samples to "<1" string
  // and add positive samples to the total quantity array
  for (let i = 0; i < a1Q.length; i++) {
    if (a1Q[i] < 1 && a1Q[i] > 0) a1Q[i] = '<1';
    if (a2Q[i] < 1 && a2Q[i] > 0) a2Q[i] = '<1';
    if (a3Q[i] < 1 && a3Q[i] > 0) a3Q[i] = '<1';
    if (a1Q[i] != 0 && a2Q[i] == 0) {
      totQ[i] = [`${a1Q[i]}`];
    } else if (a2Q[i] != 0 && a3Q[i] == 0) {
      totQ[i] = [`${a1Q[i]}, ${a2Q[i]}`];
    } else if (a3Q[i] > 0 || a3Q[i] == '<1') {
      totQ[i] = [`${a1Q[i]}, ${a2Q[i]}, ${a3Q[i]}`];
    } else if (totQ[i][0] === '') totQ[i] = ['N/A'];
  }

  // for any sample that doesn't have asbestos
  // change the asbestos type to "N/A"
  for (let i = 0; i < a1N.length; i++) {
    if (a1N[i][0] === '') a1N[i] = ['N/A'];
    if (a2N[i][0] === '') a2N[i] = ['N/A'];
    if (a3N[i][0] === '') a3N[i] = ['N/A'];
  }

  // check for to see if each sample has more than one layer
  // if more than one layer add "YES" to the "Matrix / Layers" column
  // "NO" if only one layer.
  for (let i = 0; i < lay.length; i++) {
    if (lay[i] > 1) lay[i] = ['YES'];
    else if (lay[i] == 1) lay[i] = ['NO'];
  }

  // merge all non asbestos content data together per sample id
  // add data to one array
  for (let i = 0; i < ce.length; i++) {
    if (ce[i] > 0) {
      if (ce[i] < 1) {
        nonACont[i][0] += 'Trace '
      } else {
        nonACont[i][0] += `${ce[i]} `
      }
      nAType[i][0] += 'CELLULOSE ';
    } else {
      ce[i] = '';
    }
    if (fG[i] > 0) {
      if (fG[i] < 1) {
        nonACont[i][0] += 'Trace '
      } else {
        nonACont[i][0] += `${fG[i]} `
      }
      nAType[i][0] += 'FIBROUS-GLASS ';
    } else {
      fG[i] = '';
    }
    if (sy[i] > 0) {
      if (sy[i] < 1) {
        nonACont[i][0] += 'Trace '
      } else {
        nonACont[i][0] += `${sy[i]} `
      }
      nAType[i][0] += 'SYNTHETIC ';
    } else {
      sy[i] = '';
    }
    if (ta[i] > 0) {
      if (ta[i] < 1) {
        nonACont[i][0] += 'Trace '
      } else {
        nonACont[i][0] += `${ta[i]} `
      }
      nAType[i][0] += 'TALC '
    } else {
      ta[i] = '';
    }
    if (wo[i] > 0) {
      if (wo[i] < 1) {
        nonACont[i][0] += 'Trace'
      } else {
        nonACont[i][0] += `${wo[i]}`
      }
      nAType[i][0] += 'WOLLASTONITE'
    } else {
      wo[i] = '';
    }
  }

  // format the non-asbestos content values
  for (let i = 0; i < ce.length; i++) {
    nonACont[i] = [nonACont[i][0].trim().replaceAll(" ", ", ")];
    nAType[i] = [nAType[i][0].trim().replaceAll(" ", ", ").replaceAll("-", " ")];
    if (nonACont[i][0] === '') {
      nonACont[i] = ['N/A'];
    }
    if (nAType[i][0] === '') {
      nAType[i] = ['N/A'];
    }
    if (hA[i][0] === '') {
      hA[i] = ['N/A']
    }
  }
 
  // loop through the lab report dates
  // and reformat it to YYYYMMDD from MM/DD/YY
  for (let i = 0; i < labRepDate.length; i++) {
    if (labRepDate[i][0] != 'N/A') {
      let arr = labRepDate[i][0].split("/");
      labRepDate[i][0] = `20${arr[2]}${arr[0]}${arr[1]}`;
    }
  }
  
  // get the current month and find the end of the month from the array
  // format end of month as YYYYMMDD
  const d = new Date();
  const endOfMonth = ['0131','0228','0331','0430','0531','0630','0731','0831','0930','1031','1130','1231'];
  let month = d.getMonth();
  let year = d.getFullYear();
  let repDate = [];
  for (let i = 0; i < sdbLastRow - 1; i++) repDate[i] = [`${year}${endOfMonth[month]}`];

  sampleDB.getRange(2, repDateCol, repDate.length).setValues(repDate);
  sampleDB.getRange(2, layCol, lay.length).setValues(lay);
  sampleDB.getRange(2, labAM1Col, labAM1.length).setValues(labAM1);
  sampleDB.getRange(2, labAM2Col, labAM2.length).setValues(labAM2);
  sampleDB.getRange(2, labAM3Col, labAM3.length).setValues(labAM3);
  sampleDB.getRange(2, labNameCol, labName.length).setValues(labName);
  sampleDB.getRange(2, labLocCol, labLoc.length).setValues(labLoc);
  sampleDB.getRange(2, labLocSTCol, labLocST.length).setValues(labLocST);
  sampleDB.getRange(2, labANCol, labAN.length).setValues(labAN);
  sampleDB.getRange(2, lrnCol, labRepNum.length).setValues(labRepNum);
  sampleDB.getRange(2, lrdCol, labRepDate.length).setValues(labRepDate);
  sampleDB.getRange(2, hACol, hA.length).setValues(hA);
  sampleDB.getRange(2, totQCol, totQ.length).setValues(totQ);
  sampleDB.getRange(2, a1NCol, a1N.length).setValues(a1N);
  sampleDB.getRange(2, a2NCol, a2N.length).setValues(a2N);
  sampleDB.getRange(2, a3NCol, a3N.length).setValues(a3N);
  sampleDB.getRange(2, nonAContCol, nonACont.length).setValues(nonACont);
  sampleDB.getRange(2, nATypeCol, nAType.length).setValues(nAType);
  sampleDB.getRange(2, recActCol, recAct.length).setValues(recAct);
  sampleDB.getRange(2, addFieldsCol, addFields.length).setValues(addFields);
  for (let i = 0; i < dbcHoMatNum.length; i++) {
    let arr = dbcHoMatNum[i][0].split('-');
    if (arr[arr.length - 2] === 'AC' || arr[arr.length - 2] === 'AF' || arr[arr.length - 2] === 'AW' || arr[arr.length - 2] === 'AT' || arr[arr.length - 2] === 'AM') {
      sampleDB.getRange(i+2, layCol, 1, sdbLastCol - layCol + 1).setValue('N/A');
      sampleDB.getRange(i+2, recActCol).setValue('O&M');
      sampleDB.getRange(i+2, hACol).setValue('ASSUMED');
    }
    if (arr[arr.length - 2] === 'NC' || arr[arr.length - 2] === 'NF' || arr[arr.length - 2] === 'NW' || arr[arr.length - 2] === 'NT' || arr[arr.length - 2] === 'NM') {
      sampleDB.getRange(i+2, layCol, 1, sdbLastCol - layCol + 1).setValue('N/A');
    }
  }
}

// function that formats the sample db based off user's selected location
function formatDB(location) {
  let headers = sampleDB.getRange(1, 1, 1, sdbLastCol).getValues()
  for (let i = 0; i < headers[0].length; i++) {
    if (headers[0][i] === 'Asbestos') var ha = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
    if (headers[0][i] === 'Asbestos Content %') var asc = sampleDB.getRange(2, i+1, sdbLastRow - 1).getValues();
  }

  // if the user inputs japan as the project location color red and bold all positive samples
  // else bold and color non trace samples red and trace samples color orange
  for (let i = 0; i < sdbLastRow - 1; i++) {
    if (location === 'JAPAN') {
      if (ha[i] == 'YES') {
        sampleDB.getRange(i+2, 1, 1, sdbLastCol).setFontColor("red");
        sampleDB.getRange(i+2, 1, 1, sdbLastCol).setFontWeight("bold");
      } 
      if (ha[i] == 'ASSUMED') sampleDB.getRange(i+2, 1, 1, sdbLastCol).setFontColor("red");
    } else {
      if (ha[i] != 'NO' && ha[i] != 'N/A') {
        if (asc[i] != '<1' && asc[i] != 'NO' && asc[i] != 'N/A') {
          sampleDB.getRange(i+2, 1, 1, sdbLastCol).setFontColor("red");
          sampleDB.getRange(i+2, 1, 1, sdbLastCol).setFontWeight("bold");
        } else if (asc[i] == '<1' && asc[i] != 'NO' && asc[i] != 'N/A') {
          sampleDB.getRange(i+2, 1, 1, sdbLastCol).setFontColor("orange");
        }
        if (ha[i] == 'ASSUMED') sampleDB.getRange(i+2, 1, 1, sdbLastCol).setFontColor("red");
      }
    }
  }
  
  // sort the sample db according the material order
  sortOrder(sampleDB, sdbLastRow, sdbLastCol);

  sampleDB.autoResizeColumns(1, sdbLastCol);
  sampleDB.getRange(1, 1, sdbLastRow, sdbLastCol).setHorizontalAlignment("center");
}

function deleteExtraSheets() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  //let dataSet = ss.getSheetByName("DATA");
  let cols = ss.getSheetByName("COLS")
  let dbc = ss.getSheetByName("dbc");
  let copy = ss.getSheetByName("copy");
  

  //ss.deleteSheet(dataSet);
  if (cols) ss.deleteSheet(cols);
  if (copy) ss.deleteSheet(copy);
  ss.deleteSheet(dbc);
}

// recursive function that merges rows if there's a match in the target range
// needs the target range, target sheet, initialized counter (always 0), 
// and the amount of columns you want to perform the merge on
function mergeSameCells(range, sheet, count, columns) {
  // if the value is equal to "", last row has been reached, merge and exit function
  if (range.offset(1,0).getValues()[0][0] == "") {
    sheet.getRange(`${range.offset(-count,0).getCell(1,1).getA1Notation()}:${range.getCell(1, columns).getA1Notation()}`).mergeVertically();
    return;
  }
  // if there's a matched value, increment the counter and go to the next row
  // if no match merge the matched rows, reset the counter to 0, and move to the next row
  if (range.getValues()[0][0] == range.offset(1,0).getValues()[0][0]) {
    count++;
    mergeSameCells(range.offset(1,0), sheet, count, columns);
  } else {
    sheet.getRange(`${range.offset(-count,0).getCell(1,1).getA1Notation()}:${range.getCell(1, columns).getA1Notation()}`).mergeVertically();
    count = 0;
    mergeSameCells(range.offset(1,0), sheet, count, columns);
  }
};

function construct2DArray(original, m, n) {
  if (original.length !== (m*n)) return []
  let result = []
  let arr = []
  for (let i = 0; i < original.length; i++){
    arr.push(original[i])
    if (arr.length === n){
      result.push(arr)
      arr = []
    }
  }
  return result
};

function sortOrder(sheet, lastRow, lastCol) {
  // Create sorting helper column
  var sortingCol = lastCol+1;
  sheet.getRange(1, sortingCol, 1).setValue("Letter");

  let headers = sheet.getRange(1, 1, 1, lastCol).getValues();
  for (let i = 0; i < headers[0].length; i++) {
    if (headers[0][i] == 'Homogeneous Area' || headers[0][i] == 'Homogeneous Material Number' || headers[0][i] == 'Homogenous Material Number') {
      var ha = sheet.getRange(2, i+1, lastRow - 1).getValues();
    }
  }

  let haParts = []
  let letter = [];
  for (let i = 0; i < ha.length; i++){ // From first non-header row to end
    haParts[i] = ha[i][0].split('-');
    letter[i] = haParts[i][haParts[i].length - 2]; // Penultimate SID part
  }
  
  let sortArr = [];
  for (let i = 0; i < letter.length; i++) {
    sortArr[i] = [letterOrder[letter[i]]]
  }

  sheet.getRange(2, sortingCol, lastRow - 1).setValues(sortArr)
 
  // Sort Sample IDs by letterOrder --> ID
  var fullRng = sheet.getRange(2, 1, lastRow-1, lastCol+1);
  fullRng.sort(sortingCol);
 
  // Clear extra sorting column
  sheet.getRange(1, sortingCol, lastRow).clear();
}

function onlyUnique(value, index, array) {
  return array.indexOf(value) === index;
}
