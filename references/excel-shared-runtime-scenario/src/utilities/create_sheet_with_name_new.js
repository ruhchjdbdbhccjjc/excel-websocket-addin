async function createsheetswithselectvalues() {
  await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    range.load("values,address");
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("position,name");
   
   
    
    await context.sync();
    console.log(`The active worksheet is "${sheet.name} {sheet.position}`);

    console.log(JSON.stringify(range.values, null, 4) + range.address);
    let sheetname = range.values[0][0];
    let dataSheet = context.workbook.worksheets.getItemOrNullObject(sheetname);

    await context.sync();

    if (dataSheet.isNullObject) {
      dataSheet = context.workbook.worksheets.add(sheetname);

    }
    
    dataSheet.activate();
    range.values = "";
    // Set `dataSheet` to be the second worksheet in the workbook.
    dataSheet.position = sheet.position + 1;
  });
}
