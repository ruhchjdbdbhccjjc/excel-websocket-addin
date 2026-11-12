async function openhyperlink() {
  Excel.run(async (context) => {
    console.debug("trying to  go to SelectedRange hyperlink");
    const range = context.workbook.getSelectedRange().load("hyperlink");
    await context.sync();
    if (range.hyperlink) {
      if (range.hyperlink.address != null) {
        
        window.open(range.hyperlink.address);
        console.debug(` go to  hyperlink ${range.hyperlink.address}`);
      } else if (range.hyperlink.documentReference != null) {
        
       
        await selecterange_by_address(range.hyperlink.documentReference);
        console.debug(`go to document ${range.hyperlink.documentReference}`);
      } else {
      }
    } else {
      console.error("trying to open current SelectedRange hyperlink failed ! SelectedRange dont have hyperlink ");
    }
    //return context.sync().then(() => console.log(range.hyperlink));
  });
}
/**
 * @param {Date} myDate The date
 * @param {string} myString The string
 */
async function selecterange_by_address(address) {
 
  
  Excel.run(async (context) => {
  let sheetname = await get_s_n_by_address(address);
  let sheetrange = await get_r_by_address(address);
  let recodeSheet = context.workbook.worksheets.getItemOrNullObject(sheetname);

  await context.sync();
  if (!recodeSheet.isNullObject) {
    recodeSheet.activate();
    recodeSheet.getRange(sheetrange).select();
    
  }
  });
}

