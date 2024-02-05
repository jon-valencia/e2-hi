/* --- LAST UPDATED: 01/31/2024 --- */

function createSurveyRep() {
  let ui = SpreadsheetApp.getUi();
  ss = SpreadsheetApp.getActiveSpreadsheet()
  let response = ui.prompt('PASTE PROJECT TRACKING SHEET URL HERE', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) var url = response.getResponseText();
  else return;

  let response2 = ui.prompt('Tracking Sheet Name (enter exactly like the original)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) var trackName = response2.getResponseText();
  else return;

  try {
    acmDesc();
    createAutocratSheet(url, trackName);
  } catch(err) {
    var htmlOutput = HtmlService
      .createHtmlOutput(`<a>${err}</a>`)
      .setWidth(300) //optional
      .setHeight(100); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Script Error Output');
  }
}

function acmDesc() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let cover = ss.getSheets()[0];
  let fieldWork = sheet.getRange(2, 29, sheet.getLastRow() - 1, 14).getValues();
  let hasAsbestos = sheet.getRange(2, 79, sheet.getLastRow() - 1).getValues();
  let asbestosCont = sheet.getRange(2, 80, sheet.getLastRow() - 1).getValues();
  let asbestosType = sheet.getRange(2, 81, sheet.getLastRow() - 1).getValues();
  
  let posHA = [];
  let posHAFull = [];
  let posDesc = [];
  let posFri = [];
  let posCond = [];
  let posRooms = [];
  let posQuant = [];
  let posAsbQ = [];
  let posAsbT = [];

  let acm = '';
  let assACM = '';

  //for (let i of range) console.log(i);
  for (let i of fieldWork) {
    let row = fieldWork.indexOf(i)
    let split = i[0].split('-')
    let ha = split[split.length - 2] + '-' + split[split.length - 1];
    
    if (posHA.indexOf(ha) === -1 && hasAsbestos[row][0] === 'YES') {
      posHA.push(ha)
      posHAFull.push(i[0])
      posDesc.push(i[4])
      posFri.push(i[5] === 'NF' ? 'non-friable' : i[5])
      posCond.push(i[6])
      posRooms.push(i[7])
      posQuant.push(i[12])
      posAsbQ.push(`${asbestosCont[row][0]}%`)
      posAsbT.push(asbestosType[row][0])
    } else if(posHA.indexOf(ha) > -1 && hasAsbestos[row][0] === 'YES') {
      if (!posCond[posHA.indexOf(ha)].includes(i[6])) posCond[posHA.indexOf(ha)] += ', ' + i[6];
      if (!posRooms[posHA.indexOf(ha)].includes(i[7])) posRooms[posHA.indexOf(ha)] += ', ' + i[7];
      if (!posAsbQ[posHA.indexOf(ha)].includes(`${asbestosCont[row][0]}%`)) posAsbQ[posHA.indexOf(ha)] += `, ${asbestosCont[row][0]}%`;
      if (!posAsbT[posHA.indexOf(ha)].includes(asbestosType[row][0])) posAsbT[posHA.indexOf(ha)] += ', ' + asbestosType[row][0];
      posQuant[posHA.indexOf(ha)] += i[12]
    }
    if (hasAsbestos[row][0] === 'ASSUMED') {
      assACM += `\u2022 ${i[4]} (HA ${i[0]}) in ${i[5]}, ${i[6]} condition; observed in the ${i[7]} (approximately ${i[12]} square meters).\n`
    }
  }
  
  for (let i = 0; i < posHAFull.length; i++) {
    //console.log(posAsbQ[i].join(', '))
    acm += `\u2022 ${posAsbQ[i]} ${posAsbT[i]} asbestos was identified in the ${posDesc[i].toLowerCase()} (HA ${posHAFull[i]}) in ${posFri[i].toLowerCase()}, ${posCond[i].toLowerCase()} condition; collected from ${posRooms[i]} (approximately ${posQuant[i]} square meters).\n`
  }
  cover.getRange(37, 2).setValue(acm)
  cover.getRange(38, 2).setValue(assACM)
  console.log(posFri)
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

  let insp = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1On-MPNHR9MSO0CnJXlEoek17StcftIDHbI9YoGbIvnQ/edit?usp=sharing").getSheets()[0];
  let inspLRow = insp.getLastRow();
  let inspLCol = insp.getLastColumn();
  let initials = insp.getRange(1, 1, inspLRow).getValues();
  for (let i = 0; i < initials.length; i++) {
    if (initials[i][0] === 'MDH') initials[i][0] = 'MH'; // fixed issue with Marvin's initials being MDH instead of MH on master sheet
  }

  let bNo = cover.getRange(3, 2).getValue();
  let facID = cover.getRange(4, 2).getValue();
  let stories = cover.getRange(5, 2).getValue();
  let fTeam = cover.getRange(6, 2).getValue().split(", ");
  let fTeamFull = cover.getRange(6, 2).getValue().split(", ");
  let bName = cover.getRange(7, 2).getValue();
  let loc = cover.getRange(8, 2).getValue();
  let survDate = cover.getRange(9, 2).getValue();
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
    if (deviations[i][0] !== '' && deviations[i][0] !== 'None') {
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
