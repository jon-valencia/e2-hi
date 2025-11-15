
/* UPDATED 10/24/2025
A Google Sheets sheet (i.e. one tab) that is either a lab report copy, a dbc/sample database, 
or a new sheet for script outputs. Can be modified and may either contain information or be empty. 
*/
class scriptSheet {
  constructor (sheet, workCopy=true){ // takes the target sheet as a scripts Sheet object
    this.workCopy = workCopy;
    this.parentSheet = sheet;
 
    if (this.workCopy){
      this.sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Copy");
        var fullRng = this.parentSheet.getRange(1, 1, this.parentSheet.getLastRow(), this.parentSheet.getLastColumn());
        fullRng.copyValuesToRange(this.sheet, 1, 1, 1, 1);
    }else{
      this.sheet = sheet;
    }
 
    this.lastRow = this.sheet.getLastRow(); 
    this.lastCol = this.sheet.getLastColumn();
 
    this.sidCol; // init
    this.haCol; // init
 
    for (let i=1; i<=this.lastCol; i++){ 
      if (this.sheet.getRange(1, i, 1).getValue().trim() == "Sample ID"){
        this.sidCol = i;
        continue;
      }else if(this.sheet.getRange(1, i, 1).getValue().trim() == "Homogeneous Material Number"){
        this.haCol = i;
        continue;
      }else if(this.sheet.getRange(1, i, 1).getValue().trim() == "Homogenous Material Number"){
        this.parentSheet.getRange(1, i, 1).setValue("Homogeneous Material Number");
        this.haCol = i;
        continue;
      }
    }
  }
 
  deleteCopy(){
    if (!this.workCopy) {return;} // not a working copy - do not delete the sheet!
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(this.sheet);
    return;
  }
 
  updateLastRow(){
    this.lastRow = this.sheet.getLastRow();
    return;
  }
 
  deleteExtraRows(){
    for (let i=2; i<this.lastRow+1; i++){ // From first non-header row to end
      let rng = this.sheet.getRange(i, 1, 1, this.lastCol); // full row
      if (rng.isBlank()){
        rng.deleteCells(SpreadsheetApp.Dimension.ROWS);
      }
    }
 
    this.lastRow = this.sheet.getLastRow(); // updateLastRow(); //?
    return;
  }
 
  sortOrder(){
    // Create sorting helper column
    var sortingCol = this.lastCol+1;
    this.sheet.getRange(1, sortingCol, 1).setValue("Letter"); 
 
    for (let i=2; i<this.lastRow+1; i++){ // From first non-header row to end
      let ha = this.sheet.getRange(i, this.haCol, 1).getValue();
      if (this.sheet.getRange(i, 1, 1).isBlank()){
        this.sheet.getRange(i, sortingCol, 1).setValue(16); // there should not be an empty HA
        continue;
      }
      let haParts = ha.split('-');
      let letter = haParts[haParts.length - 2]; // Penultimate SID part
 
      this.sheet.getRange(i, sortingCol, 1).setValue(letterOrder[letter]);
    }
 
    // Sort Sample IDs by letterOrder --> ID
    var fullRng = this.sheet.getRange(2, 1, this.lastRow-1, this.lastCol+1);
    fullRng.sort([{column: sortingCol, ascending: true}, {column: this.sidCol, ascending: true}]);
 
    // Clear extra sorting column
    this.sheet.getRange(1, sortingCol, this.lastRow).clear();
 
    return;
  }
 
  unhideCols(){
    this.sheet.getRange('A:A').activate();
    SpreadsheetApp.getActiveSheet().showColumns(this.sheet.getActiveRange().getColumn(), this.sheet.getActiveRange().getNumColumns());
    return;
  }
}
 
 
function createMLeaders(){
  const sheetObj = new scriptSheet(SpreadsheetApp.getActiveSheet(), true);
  const sheet = sheetObj.sheet;
  sheetObj.unhideCols();
  sheetObj.deleteExtraRows();
  sheetObj.sortOrder();
 
  var fileText = ''; // init
 
  fileText += '(command "HPLINETYPE" "ON")\n'; // constants
  fileText += '(command "OSNAP" "")\n';
  fileText += '(command "COLOR" "BYLAYER")\n';
  fileText += '(command "CMLEADERSTYLE" "E2 ASBESTOS - NEG")\n'; // CHECK - write only once since recoloring later?
 
  var sidCol = sheetObj.sidCol;
 
  // Designates starting coordinates for mleaders. The first mleader will have a point at -15,15 in paperspace
  // All mleaders have 0.5 vertical separation and 2.5 horizontal separation from point to point
  // x1 & x2 and y1 & y2 are differentiated in order to preserve the x1 and y1 variables
  var x1 = -15
  var x2 = x1 + 0.5
  var y1 = 15.5
  var y2 = y1 + 1
 
  var curParts = sheet.getRange(2, sidCol, 1).getValue().split('-');
  var curMaterial = curParts[curParts.length-2];
  var curSID = sheet.getRange(2, sidCol, 1).getValue();
  var curHA = curSID.substring(0, curSID.length-1); // remove the last letter
 
  fileText += '(command "CLAYER" "' + curHA + '")\n';
 
  for (i=2; i<=sheet.getLastRow(); i++){ 
    var sidParts = sheet.getRange(i, sidCol, 1).getValue().split('-');
    if (sidParts.length < 3){continue};
    var material = sidParts[sidParts.length-2];
    var sid = sheet.getRange(i, sidCol, 1).getValue();
    var ha = sid.substring(0, sid.length-1); // remove the last letter
 
    if (material[0] == "A" || material[0] == "N"){ // Assumed and non-suspect samples do not need mleaders
      continue;
      // break; // Sheet has been reordered above, so an assumed or non-suspect sample should be the end
    }else if (material == curMaterial){ // same material, increment down the column
      y1 -= 0.5;
      y2 = y1 + 1;
    }else{ // go to the next column
      curMaterial = material;
      x1 += 2.5;
      x2 = x1 + 0.5;
      y1 = 15;
      y2 = y1 + 1;
    }
 
    if (curHA != ha){
      curHA = ha;
      fileText += '(command "CLAYER" "' + curHA + '")\n';
    }
 
    fileText += '(command "MLEADER" "' + x1 + ',' + y1 + '" "' + x2 + ',' + y2 + '" "' + sid + '")\n';
  }
 
  legend = createHALegend(sheetObj); // for debugging
  fileText += legend;
 
  var facNum; // init 
  for (i=1; i<=SpreadsheetApp.getActiveSpreadsheet().getLastColumn(); i++){ 
    if (sheet.getRange(1, i, 1).getValue() == "Facility Number"){
      facNum = sheet.getRange(2, i, 1).getValue();
      if (sheet.getRange(2, i, 1).isBlank()){
        facNum = "FacilityNumberNotFound";
      }
      facNum = String(facNum);
      facNum.replaceAll(' ', '_'); // no spaces in filename
      facNum.replaceAll('/', '_'); // no slashes in filename
      break;
    }
  }
  fileName = '' + facNum + '_MLeaderGenerator.scr';
 
  sheetObj.deleteCopy(); // get rid of extra sheet before download pops up
 
  href = DriveApp.createFile(fileName, fileText).getDownloadUrl();
  var htmlOutput = HtmlService
    .createHtmlOutput('<a target="_blank" href="' + href + '" >Click to download</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download below');
 
  return;
}
 
 
function createHALegend(sheetObj){ // takes in mleader sheet, assumes it has already been checked for errors
  var sheet = sheetObj.sheet;
  var haCol = sheetObj.haCol;
 
  var x = -15;
  var y = 23;
  var distance = 4;
 
  var legend = '(command "CLAYER" "LEGEND")\n(command "COLOR" "BYLAYER")\n(command "LTSCALE" "1")\n';
  legend += '(command "MTEXT" "' + String(x) + ',' + String(y) + '" "' + String(Number(x)+Number(distance)) + ',' + String(y) + '" "\\{HOMOGENEOUS AREAS\\\\P\\\\P'; // legend does not have any tabs
  x += Number(distance);
  var descCol; // init
 
  for (let i=1; i<=SpreadsheetApp.getActiveSpreadsheet().getLastColumn(); i++){ 
    if (sheet.getRange(1, i, 1).getValue().trim() == "Material Description"){
      descCol = i;
      break;
    }
  }
 
  var fullHA = sheet.getRange(2, haCol, 1).getValue().split('-'); 
  var curHA = fullHA[fullHA.length-2] + '-' + fullHA[fullHA.length-1];
  var curMat = curHA.substring(0, 1); // only take the first character
  var curDesc = sheet.getRange(2, descCol, 1).getValue();
  curDesc = curDesc.replaceAll('"', '\\"');
  legend += '' + curHA + ' ' + curDesc + '\\\\P\\\\P';
 
  // needs to change per type as well
  for (i=2; i<=sheet.getLastRow(); i++){
    fullHA = sheet.getRange(i, haCol, 1).getValue().split('-'); 
    ha = fullHA[fullHA.length-2] + '-' + fullHA[fullHA.length-1];
    mat = ha.substring(0, 1);
    if (mat != curMat){
      curMat = mat;
      legend += '\\}")\n\n(command "MTEXT" "' + String(x) + ',' + String(y) + '" "' + String(Number(x)+Number(distance)) + ',' + String(y) + '" "\\{HOMOGENEOUS AREAS\\\\P\\\\P';
      x += Number(distance);
    }
    if (ha != curHA){ // only act if we're in a new HA
      curHA = ha;
      curMat = curHA.substring(0, 1);
      curDesc = sheet.getRange(i, descCol, 1).getValue().trim();
      curDesc = curDesc.replaceAll('"', '\\"');
      legend += '' + curHA + ' ' + curDesc + '\\\\P\\\\P';
    }
  }
  
  legend += `\\}")\n\n(defun C:MT2R ; = MText [to] Romans
  (/ tss n tdata)
  (if (setq tss (ssget "_X" '((8 . "LEGEND"))))
    (repeat (setq n (sslength tss)); then
      (setq
        tdata (entget (ssname tss (setq n (1- n))))
        tdata (subst '(7 . "ROMANS") (assoc 7 tdata) tdata); Style
        tdata (subst '(40 . 0.09) (assoc 40 tdata) tdata); height
      ); setq
      (entmod tdata)
    ); repeat
  ); if
  (princ)
); defun
MT2R\n`
 
  return legend;
}
 
 
// Layer creation for 
// @requires 'Homogeneous Material Number' and 'Facility Number' to be consistent columns in all SDs/DBCs
function createLayersHALinetype(){
  const sheetObj = new scriptSheet(SpreadsheetApp.getActiveSheet(), true);
  const sheet = sheetObj.sheet;
  sheetObj.unhideCols();
  sheetObj.deleteExtraRows();
  sheetObj.sortOrder();
 
  var fileText = ''; // init
 
  fileText += '(setvar "expert" 3)\n' // this line is only essential for HA linetypes
  fileText += '(ltscale 1)\n'
  fileText += '-layer m LEGEND c green LEGEND \n'
 
  var haCol = sheetObj.haCol;
 
  var curHA = sheet.getRange(2, haCol, 1).getValue(); 
  var split = curHA.split('-');
  var haEnd = split[split.length-2] + '-' + split[split.length-1];
  var color = 96; //negative
  if (haEnd[0] == "A"){color = 10}; // assumed positive
  fileText += '(command "-linetype" "load" "_' + haEnd + '" "acadiso.lin" "")\n'
  fileText += '-layer m ' + curHA + ' c ' + color + ' ' + curHA + ' l _' + haEnd + ' ' + curHA + ' \n'
 
  for (i=2; i<=sheet.getLastRow(); i++){
    if (sheet.getRange(i, haCol, 1).getValue() != curHA){ // only act if we're in a new HA
      curHA = sheet.getRange(i, haCol, 1).getValue();
      split = curHA.split('-');
      haEnd = split[split.length-2] + '-' + split[split.length-1];
      fileText += '(command "-linetype" "load" "_' + haEnd + '" "acadiso.lin" "")\n'
      color = 96; //negative
      if (haEnd[0] == "A"){color = 10};
      fileText += '-layer m ' + curHA + ' c ' + color + ' ' + curHA + ' l _' + haEnd + ' ' + curHA + ' \n'
    }
  }
 
  var facNum; // init
  for (i=1; i<=SpreadsheetApp.getActiveSpreadsheet().getLastColumn(); i++){ 
    if (sheet.getRange(1, i, 1).getValue() == "Facility Number"){
      facNum = sheet.getRange(2, i, 1).getValue();
      if (sheet.getRange(2, i, 1).isBlank()){
        facNum = "FacilityNumberNotFound";
      }
      facNum = String(facNum);
      facNum.replaceAll(' ', '_'); // no spaces in filename
      facNum.replaceAll('/', '_'); // no slashes in filename
      break;
    }
  }
  fileName = '' + facNum + '_LayerGenerator_HALinetype.scr';
 
  sheetObj.deleteCopy(); // get rid of extra sheet before download pops up
 
  href = DriveApp.createFile(fileName, fileText).getDownloadUrl();
  var htmlOutput = HtmlService
    .createHtmlOutput('<a target="_blank" href="' + href + '" >Click to download</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download below');
 
  return;
}
 
// Create wrapper for layer creation - continuous linetype or HA linetype, start with HA linetype here, then reconfig
// @requires 'Homogeneous Material Number' and 'Facility Number' to be consistent columns in all SDs/DBCs
function createLayersContinuousLinetype(){
  const sheetObj = new scriptSheet(SpreadsheetApp.getActiveSheet(), true);
  const sheet = sheetObj.sheet;
  sheetObj.unhideCols();
  sheetObj.deleteExtraRows();
  sheetObj.sortOrder();
 
  var fileText = ''; // init
 
  fileText += '(setvar "expert" 3)\n' // this line is only essential for HA linetypes
  fileText += '(ltscale 1)\n'
 
  var haCol = sheetObj.haCol;
 
  var curHA = sheet.getRange(2, haCol, 1).getValue(); 
  var split = curHA.split('-');
  var haEnd = split[split.length-2];
  var color = 96; //negative
  if (haEnd[0] == "A"){color = 10}; // assumed positive
  fileText += '-layer m ' + curHA + ' c ' + color + ' ' + curHA + ' l Continuous ' + curHA + ' \n'
 
  for (i=2; i<=sheet.getLastRow(); i++){
    if (sheet.getRange(i, haCol, 1).getValue() != curHA){ // only act if we're in a new HA
      curHA = sheet.getRange(i, haCol, 1).getValue();
      split = curHA.split('-');
      haEnd = split[split.length-2];
      color = 96; //negative
      if (haEnd[0] == "A"){color = 10}; // assumed positive
      fileText += '-layer m ' + curHA + ' c ' + color + ' ' + curHA + ' l Continuous ' + curHA + ' \n'
    }
  }
 
  var facNum; // init
  for (i=1; i<=SpreadsheetApp.getActiveSpreadsheet().getLastColumn(); i++){ 
    if (sheet.getRange(1, i, 1).getValue() == "Facility Number"){
      facNum = sheet.getRange(2, i, 1).getValue();
      if (sheet.getRange(2, i, 1).isBlank()){
        facNum = "FacilityNumberNotFound";
      }
      facNum = String(facNum);
      facNum.replaceAll(' ', '_'); // no spaces in filename
      facNum.replaceAll('/', '_'); // no slashes in filename
      break;
    }
  }
  fileName = '' + facNum + '_LayerGenerator_ContinuousLinetype.scr';
 
  sheetObj.deleteCopy(); // get rid of extra sheet before download pops up
 
  href = DriveApp.createFile(fileName, fileText).getDownloadUrl();
  var htmlOutput = HtmlService
    .createHtmlOutput('<a target="_blank" href="' + href + '" >Click to download</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download below');
 
  return;
}

function createLBPMLeaders(){
  let ss = SpreadsheetApp.getActiveSheet();
  let selection = ss.getSelection();
  
  let fileText = ''; // init
 
  fileText += '(command "HPLINETYPE" "ON")\n'; // constants
  fileText += '(command "OSNAP" "")\n';
  fileText += '(command "COLOR" "BYLAYER")\n';
  fileText += '(command "CMLEADERSTYLE" "E2 LEAD - NEG")\n'; // CHECK - write only once since recoloring later?
  
  // Designates starting coordinates for mleaders. The first mleader will have a point at -15,15 in paperspace
  // All mleaders have 0.5 vertical separation and 2.5 horizontal separation from point to point
  // x1 & x2 and y1 & y2 are differentiated in order to preserve the x1 and y1 letiables
  let x1 = -15
  let x2 = x1 + 0.5
  let y1 = 15.5
  let y2 = y1 + 1
 
  let sids = selection.getActiveRange().getValues();
  let curSID = selection.getActiveRange().getValue();
  let curArea = curSID.substring(0, curSID.length - 3).replace("/","_");
  
  fileText += '(command "CLAYER" "' + curArea + '")\n';
  
  for (let i of sids) {
    let sid = i[0].replace("/","_");
    console.log(sid)
    let area = sid.substring(0, sid.length - 3);
    curArea = curArea.replace("/","_")
    console.log(curArea)
    if (area === curArea) {
      y1 -= 0.5;
      y2 = y1 + 1;
    } else {
      curArea = area;
      fileText += `(command "CLAYER" "${area}")\n`
      x1 += 2.5;
      x2 = x1 + 0.5;
      y1 = 15;
      y2 = y1 + 1;
    }

    fileText += `(command "MLEADER" "${x1}, ${y1}" "${x2}, ${y2}" "${sid}")\n`;
  }

  console.log(fileText)
 
  let split = curSID.split("-");
  let facNum = split[split.length - 3];
  fileName = '' + facNum + '_LBPMLeaderGenerator.scr';
  
  console.log(fileName)
  href = DriveApp.createFile(fileName, fileText).getDownloadUrl();
  let htmlOutput = HtmlService
    .createHtmlOutput('<a target="_blank" href="' + href + '" >Click to download</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download below');
}

function createLBPLayers(){
  let ss = SpreadsheetApp.getActiveSheet();
  let selection = ss.getSelection();
 
  let fileText = ''; // init
 
  fileText += '(setlet "expert" 3)\n'; // this line is only essential for HA linetypes
  fileText += '(ltscale 1)\n';

  let sids = selection.getActiveRange().getValues();
  let curSID = selection.getActiveRange().getValue().replace("/","_");
  let curArea = curSID.substring(0, curSID.length - 3).replace("/","_");
  let color = 96;
  
  fileText += `-layer m ${curArea} c ${color} ${curArea} l Continuous ${curArea} \n`
  console.log(fileText)

  for (let i of sids) {
    let sid = i[0];
    let area = sid.substring(0, sid.length - 3).replace("/","_");
    if (area !== curArea) {
      curArea = area;
      fileText += `-layer m ${curArea} c ${color} ${curArea} l Continuous ${curArea} \n`
    }
  }
  console.log(fileText)
 
  let split = curSID.split('-');
  let facNum = split[split.length - 3];
  fileName = '' + facNum + '_LBPLayerGenerator.scr';
  
  href = DriveApp.createFile(fileName, fileText).getDownloadUrl();
  let htmlOutput = HtmlService
    .createHtmlOutput('<a target="_blank" href="' + href + '" >Click to download</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download below');
 
}

function createLBPLegend(){ // takes in mleader sheet, assumes it has already been checked for errors
 
  var legend = '(command "CLAYER" "LEGEND")\n(command "COLOR" "BYLAYER")\n(command "LTSCALE" "1")\n';
 
  legend += `\\}")\n\n(defun C:MT2R ; = MText [to] Romans
  (/ tss n tdata)
  (if (setq tss (ssget "_X" '((8 . "LEGEND"))))
    (repeat (setq n (sslength tss)); then
      (setq
        tdata (entget (ssname tss (setq n (1- n))))
        tdata (subst '(7 . "ROMANS") (assoc 7 tdata) tdata); Style
        tdata (subst '(40 . 0.09) (assoc 40 tdata) tdata); height
      ); setq
      (entmod tdata)
    ); repeat
  ); if
  (princ)
); defun
MT2R`
 
  return legend;
}

/* Asbestos Mleader/Hatch Recoloring Script */ 
function createRecolor(){ // Last priority
  //throw new Error("NIY");
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi()
  let database = ss.getActiveSheet();
  let posHAs = [];
  let traceHAs = [];
  let posSIDs = [];
  let traceSIDs = [];
  let assHAs = [];
  let doTrace;

  let response = ui.alert('Does this project use trace results?', 'Generally, Japan-based project do not use trace, but US-based projects do.', ui.ButtonSet.YES_NO_CANCEL);
  switch (response) {
    case ui.Button.YES:
      doTrace = true;
      break;
    case ui.Button.NO:
      console.log("yello")
      doTrace = false;
      break;
    default:
      return;
  }
  console.log(doTrace)
  // Change 2nd number to match 'Asbestos' column in sample database
  let hasAsbestos = database.getRange(1, 76, database.getLastRow(), 2).getValues();

  // row[0] - 'Asbestos' column in sample database
  // row[1] - 'Asbestos Content %' column in sample database
  // 29 - 'Homogeneous Area' column in sample database
  // 65 - 'Sample ID' column in sample database

  // iterate through 'Asbestos' column, if 'YES', check for '<1' (trace)
  // separate SIDs based on positive and trace
  // separate HAs based on positive and trace. If HA has positive and trace samples, HA is positive.
  for (let row of hasAsbestos) {
    let rowNum = hasAsbestos.indexOf(row) + 1;
    if (row[0] == 'YES') {
      if (row[1] == '<1' && doTrace || row[1] < 1 && doTrace) {
        traceSIDs.push(database.getRange(rowNum, 65).getValue()) 
        if (traceHAs.indexOf(database.getRange(rowNum, 29).getValue()) == -1 && posHAs.indexOf(database.getRange(rowNum, 29).getValue()) == -1) 
          traceHAs.push(database.getRange(rowNum, 29).getValue());
      }
      else {
        posSIDs.push(database.getRange(rowNum, 65).getValue());
        if (posHAs.indexOf(database.getRange(rowNum, 29).getValue()) == -1) 
          posHAs.push(database.getRange(rowNum, 29).getValue());
        // remove from trace HA array if a positive sample is identified for the same HA
        if (traceHAs.indexOf(database.getRange(rowNum, 29).getValue() > -1) && posHAs.indexOf(database.getRange(rowNum, 29).getValue()) == -1) {
          let haToRemove = traceHAs.indexOf(database.getRange(rowNum, 29).getValue());
          traceHAs.splice(haToRemove, 1);
          posHAs.push(database.getRange(rowNum, 29).getValue());
        }
      }
    }
  }

  //console.log(posHAs);
  //console.log(posSIDs);

  let posSIDText = '';
  let traceSIDText = '';
  let haText = '';

  // formatting for different types of sample results
  // '\\L' - underline for posSIDs
  // '\\Q10' - italics for traceSIDs
  for (let sid of posSIDs) {
    posSIDText += `((wcmatch str "*${sid}*") (list "${sid}" "{\\\\L${sid}}"));\n`;
  }
  for (let sid of traceSIDs) {
    traceSIDText += `((wcmatch str "*${sid}*") (list "${sid}" "{\\\\Q10${sid}}"));\n`
  }
  for (let ha of posHAs) {
    haText += `"_color" 10 "${ha}";\n`;
  }

  for (let ha of traceHAs) {
    haText += `"_color" 30 "${ha}";\n`
  }

  //console.log(sidText);
  //console.log(haText);

  let text = 
`(defun c:LCA (); = Layer Color Assignment
  (command "_.layer"
    ${haText}
    "";
  )
)
(c:LCA)
(defun c:REPLACEMLEADERS  (/ sel e)
    (if (setq sel (ssget '((0 . "MULTILEADER")(8 . "*"))));
      (repeat (sslength sel)
            (setq e (vlax-ename->vla-object (ssname sel 0))
                  str (vla-get-textstring e))
            (if (setq f	 
              (cond
                ${posSIDText}
            ))
            (progn 
                      (vla-put-textstring e
                      (vl-string-subst (cadr f)(car f) str) )
                      (vlax-put-property e 'StyleName "E2 ASBESTOS - POS")
                  )
                                )
            (if (setq f	 
              (cond
                ${traceSIDText}
            ))
            (progn 
                      (vla-put-textstring e
                      (vl-string-subst (cadr f)(car f) str) )
                      (vlax-put-property e 'StyleName "E2 ASBESTOS - TRACE")
                  )
                                )
            (ssdel (ssname sel 0) sel)
            )
        )
        (princ)
    )
(c:REPLACEMLEADERS)
all

`;

  //console.log(text)
  
  let href = DriveApp.createFile('recolor_mleaders.scr', text).getDownloadUrl();
  let htmlOutput = HtmlService
    .createHtmlOutput('<a target="_blank" href="' + href + '" >Click to download</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download below');
}
