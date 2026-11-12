// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    ensureStateInitialized(true);
    console.log("ensure state initialized from the office.initialize");
    isOfficeInitialized = true;
    registerClickHandler();
    registerrecodeClickHandler();
    var intervalID = setInterval(getcommand, 200);
    var intervalID_indes = setInterval(resetpreviosindex, 180000);
    var intervalID_reload = setInterval(Refresh, 6000000);
   
 
 
    monitorSheetChanges();

    document.getElementById("connectService").onclick = connectService; // in office-apis-helpers.js
    document.getElementById("selectFilter").onclick = insertFilteredData;
    //new Service
    document.getElementById("connectService_new").onclick = reconnectService_new;
    
    updateRibbon();
    updateTaskPaneUI();
  }
});
function Refresh(){
  location.reload();
 }

async function insertFilteredData() {
  try {
    //Determine which data source the user selected from the radio buttons.
    const radioExcel = document.getElementById("communicationFilter");
    if (radioExcel.checked) {
      generateCustomFunction("Communications");
    } else {
      generateCustomFunction("Groceries");
    }
  } catch (error) {
    console.error(error);
  }
}
//
async function reconnectService_new(){
  try {
    await registerClickHandler();
    await registerrecodeClickHandler();
    var intervalID = setInterval(getcommand, 200);
    var intervalID = setInterval(resetpreviosindex, 180000);
   
  } catch (error) {
    console.error(error);
  }

  
}
