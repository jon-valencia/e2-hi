/* --- LAST UPDATED: 11/06/2023 --- */

function createSurveyRep() {
  let ui = SpreadsheetApp.getUi();
  ss = SpreadsheetApp.getActiveSpreadsheet()
  let response = ui.prompt('PASTE PROJECT TRACKING SHEET URL HERE', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) var url = response.getResponseText();
  else return;

  let response2 = ui.prompt('Tracking Sheet Name (enter exactly like the original)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) var trackName = response2.getResponseText();
  else return;

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
  if (!ss.getSheetByName("copy")) {
    appa = ss.insertSheet("copy");
    appaLastRow = appa.getLastRow();
    appaLastCol = appa.getLastColumn();
  } else {
    appa = ss.getSheetByName("copy");
    appaLastRow = appa.getLastRow();
    appaLastCol = appa.getLastColumn();
  }

  try {
    createDBC();
    formatLabData();
    createLabExport();
    handleAsbestos();
    handleAssumed();
    updateCoversheet();
    deleteExtraSheets();
    createAutocratSheet(url, trackName);
  } catch(err) {
    var htmlOutput = HtmlService
      .createHtmlOutput(`<a>${err}</a>`)
      .setWidth(300) //optional
      .setHeight(100); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Script Error Output');
  }
}

function createAutocratSheet(url, trackName) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let cover = ss.getSheets()[0];
  let dash = SpreadsheetApp.openByUrl(url);
  let track = dash.getSheetByName(trackName);
  let trackLRow = track.getLastRow();
  let trackLCol = track.getLastColumn();
  let tHeaders = track.getRange(2, 1, 1, trackLCol).getValues();
  for (let i = 0; i < tHeaders[0].length; i++) {
    if (tHeaders[0][i] === 'Building Number' || tHeaders[0][i] === 'Building No.' || tHeaders[0][i] === 'BUILDING NO.') var bNos = track.getRange(2, i+1, trackLRow).getValues();
    if (tHeaders[0][i] === 'Area (SF)') var sqftCol = i+1;
    if (i+1 === trackLCol) var filenameCol = i+1;
  }

  // Error checking if can't find building number column and area column
  if (bNos === undefined) throw new Error('Could not find "Building Number" column');
  if (sqftCol === undefined) sqftCol = 'NEED AREA';
  console.log(sqftCol)

  let insp = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1On-MPNHR9MSO0CnJXlEoek17StcftIDHbI9YoGbIvnQ/edit?usp=sharing").getSheets()[0];
  let inspLRow = insp.getLastRow();
  let inspLCol = insp.getLastColumn();
  let initials = insp.getRange(1, 1, inspLRow).getValues();
  for (let i = 0; i < initials.length; i++) {
    if (initials[i][0] === 'MH') initials[i][0] = 'MDH'; // fixed issue with Marvin's initials being MDH instead of MH on master sheet
  }

  let bNo = cover.getRange(3, 2).getValue();
  let facID = cover.getRange(4, 2).getValue();
  let stories = cover.getRange(5, 2).getValue();
  let fTeam = cover.getRange(6, 2).getValue().split(", ");
  let fTeamFull = cover.getRange(6, 2).getValue().split(", ");
  let bName = cover.getRange(7, 2).getValue();
  let loc = cover.getRange(8, 2).getValue();
  let survDate = formatDate(cover.getRange(9, 2).getValue())
  let deviations = cover.getRange(10, 2, 5).getValues();
  let totHA = cover.getRange(21, 4).getValue();
  let totSID = cover.getRange(26, 4).getValue();
  let posA = cover.getRange(37, 2).getValue();
  let assuA = cover.getRange(38, 2).getValue();
  for (let i = 0; i < bNos.length; i++) {
    if (bNo == bNos[i][0]) {
      if (sqftCol === 'NEED AREA') var sqft = sqftCol;
      else var sqft = track.getRange(i+2, sqftCol).getValue();  
      var filename = track.getRange(i+2, filenameCol).getValue();
    }
  }

  for (let i = 0; i < fTeam.length; i++) {
    for (let j = 0; j < initials.length; j++) {
      if (fTeam[i] === initials[j][0]) {
        fTeam[i] = insp.getRange(j+1, 1, 1, inspLCol - 1).getValues()[0];
        fTeamFull[i] = insp.getRange(j+1, 2).getValue()
      }
    }
  }
  let devs = "";
  for (let i = 0; i < deviations.length; i++) {
    if (deviations[i][0] !== '' && deviations[i][0] !== 'None' && deviations[i][0] !== 'N/A') {
      devs += `${deviations[i][0]}\n\n`;
    }
  }

  if (!ss.getSheetByName("auto")) var auto = ss.insertSheet("auto");
  else auto = ss.getSheetByName("auto");

  let ah = [['Filename', 'Building_No', 'Building_Name', 'Facility_ID', 'Location', 'SF', 'Stories', 
  'Survey_Date', 'Total_Samples', 'Total_HAs', 'Positive_Asbestos', 'Assumed_Asbestos', 'Deviations','Inspectors_All', 
  'initials_1','Inspector_1', 'first_1', 'middle_1', 'last_1', 'accLoc_1', 'Inspector_1_Accreditation', 'Company_1', 'compNo_1','date_1',
  'initials_2','Inspector_2', 'first_2', 'middle_2', 'last_2', 'accLoc_2', 'Inspector_2_Accreditation', 'Company_2', 'compNo_2','date_2',
  'initials_3','Inspector_3', 'first_3', 'middle_3', 'last_3', 'accLoc_3', 'Inspector_3_Accreditation', 'Company_3', 'compNo_3','date_3',
  'initials_4','Inspector_4', 'first_4', 'middle_4', 'last_4', 'accLoc_4', 'Inspector_4_Accreditation', 'Company_4', 'compNo_4','date_4',
  'initials_5','Inspector_5', 'first_5', 'middle_5', 'last_5', 'accLoc_5', 'Inspector_5_Accreditation', 'Company_5', 'compNo_5','date_5'
  ]];

  auto.getRange(1, 1, 1, 64).setValues(ah);
  auto.getRange(2, 1).setValue(filename);
  auto.getRange(2, 2).setValue(bNo);
  auto.getRange(2, 3).setValue(bName);
  auto.getRange(2, 4).setValue(facID);
  auto.getRange(2, 5).setValue(loc);
  auto.getRange(2, 6).setValue(sqft.toLocaleString());
  auto.getRange(2, 7).setValue(stories);
  auto.getRange(2, 8).setValue(survDate);
  auto.getRange(2, 9).setValue(totSID);
  auto.getRange(2, 10).setValue(totHA);
  auto.getRange(2, 11).setValue(posA);
  auto.getRange(2, 12).setValue(assuA);
  auto.getRange(2, 13).setValue(devs.trim());
  auto.getRange(2, 14).setValue(fTeamFull.join(", "));

  for (let i = 0; i < fTeam.length; i++) {
    auto.getRange(2, 15+i*fTeam[i].length, 1, 10).setValues([fTeam[i]]);
  }
  auto.setRowHeightsForced(2, 1, 21);
  let aLC = auto.getLastColumn()
  for (let i = 0; i < aLC; i++) {
    if (auto.getRange(2, i+1).getValue() == 'Element Environmental, LLC') auto.getRange(2, i+1).setValue('Inspector, E2')
  }
}

function formatDate(date) {
  let start = date.split(" ")[0];
  let end = date.split(" ")[2];
  console.log(end)
  if(end === undefined) {
    return new Date(`"${start.split("/")[0]},${start.split("/")[1]},${start.split("/")[2]}"`).toLocaleDateString('en-US', {year: 'numeric', month: 'long', day: 'numeric'});
  } else {
    const date_start = new Date(`"${start.split("/")[0]},${start.split("/")[1]},${start.split("/")[2]}"`).toLocaleDateString('en-US', {year: 'numeric', month: 'long', day: 'numeric'});
    const date_end = new Date(`"${end.split("/")[0]},${end.split("/")[1]},${end.split("/")[2]}"`).toLocaleDateString('en-US', {year: 'numeric', month: 'long', day: 'numeric'});
    return date_start + " - " + date_end;
  }
}
