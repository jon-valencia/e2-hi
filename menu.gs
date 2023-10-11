const letterOrder = {
'C': 1, 
'F': 2, 
'W': 3, 
'T': 4, 
'M': 5, 
'AC': 6, 
'AF': 7, 
'AW': 8, 
'AT': 9, 
'AM': 10, 
'NC': 11, 
'NF': 12, 
'NW': 13, 
'NT': 14, 
'NM': 15
};

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('E2 Scripts')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('CAD Scripts')
        .addItem('MLeaders v1.0', 'createMLeaders')
        .addItem('Layers - HA Linetype v1.0', 'createLayersHALinetype')
        .addItem('Layers - Continuous Linetype v1.0', 'createLayersContinuousLinetype')
        .addItem('MLeaders LBP', 'createLBPMLeaders')
        .addItem('Layers LBP', 'createLBPLayers'))
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Sample DB Scripts')
        .addItem('Create App A', 'createAppA')
        .addItem('Update DB', 'updateDB'))
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Survey Report Script')
        .addItem('Create Report Data Sheet', 'createSurveyRep'))
      .addToUi();
}
