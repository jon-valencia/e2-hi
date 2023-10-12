function superSizeMe() {
  let ss = SpreadsheetApp.getActiveSheet();
  let selection = ss.getSelection();
  let arr = selection.getActiveRange().getValues()
  let arr2 = []

  for (i of arr) {
    arr2.push(i)
    arr2.push(i)
    arr2.push(i)
  }

  range = ss.getRange(selection.getActiveRange().getRow(), selection.getActiveRange().getColumn(), arr2.length, selection.getActiveRange().getLastColumn() - selection.getActiveRange().getColumn() + 1);
  range.setValues(arr2)

  console.log(selection.getActiveRangeList().toString())

}
