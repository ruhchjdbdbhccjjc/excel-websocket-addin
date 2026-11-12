const navigation_sheet_name = "导航";
const high_distance = 1;
const length_distance = 3;
const isolation_distance = 2;

async function setnagition() {
  await settargetsheet(navigation_sheet_name);
  //await sheetname();
  await write_navigation();
}
/* get sheet name  "1


Basically, I think what you want to do is get a reference to the Workbook.worksheets property. Load the name property and call context.sync---excel - How do I list all workbook sheets in task pane [add-in] with office-js? - Stack Overflowundefined---Last Sync: 21/3/2024 @ 13:44:4" https://stackoverflow.com/questions/56435621/how-do-i-list-all-workbook-sheets-in-task-pane-add-in-with-office-js#:~:text=created%20(oldest%20first)-,1,call%20context.sync,-.%20After%20the%20sync */
async function sheetname() {
  await Excel.run(function(context) {
    var worksheets = context.workbook.worksheets;
    worksheets.load("name");
    return context.sync().then(function() {
      for (var i = 0; i < worksheets.items.length; i++) {
        console.log(worksheets.items[i].name);
      }
    });
  });
}

async function write_navigation_with_address() {
  await Excel.run(async (context) => {
    let sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();

    if (sheets.items.length > 1) {
      console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
    } else {
      console.log(`There is one worksheet in the workbook:`);
    }

    for (var i = 0; i < sheets.items.length; i++) {
      const worksheet = context.workbook.worksheets.getItem(navigation_sheet_name);
      const cell_name = worksheet.getCell(high_distance * i, length_distance);
      cell_name.values = [[sheets.items[i].name]];
      const cell_hyperlink = worksheet.getCell(high_distance * i, length_distance + isolation_distance);
      let cellText = sheets.items[i].name + "!A1";
      let hyperlink = {
        textToDisplay: cellText,
        screenTip: "Go to '" + cellText + "'",
        documentReference: cellText
      };
      cell_hyperlink.hyperlink = hyperlink;
      console.log(sheets.items[i].name);
    }
    await context.sync();
  });
}
async function write_navigation() {
  await Excel.run(async (context) => {
    let sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();

    if (sheets.items.length > 1) {
      console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
    } else {
      console.log(`There is one worksheet in the workbook:`);
    }

    for (var i = 0; i < sheets.items.length; i++) {
      const worksheet = context.workbook.worksheets.getItem(navigation_sheet_name);
      const cell_name = worksheet.getCell(high_distance * i, length_distance);
      cell_name.values = [[sheets.items[i].name]];

      let cellText = sheets.items[i].name + "!A1";
      let hyperlink = {
        textToDisplay: sheets.items[i].name,
        screenTip: "Go to '" + cellText + "'",
        documentReference: cellText
      };
      cell_name.hyperlink = hyperlink;
      console.log(sheets.items[i].name);
    }
    await context.sync();
  });
}
//https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties:~:text=by%20using%20the%20getItemOrNullObject()%20method.%20If%20a%20worksheet%20with%20that%20name%20does%20not%20exist%2C%20a%20new%20sheet%20is%20created.%20Note%20that%20the%20code%20does%20not%20load%20the%20isNullObject%20property
async function settargetsheet(sheetname) {
  await Excel.run(async (context) => {
    let dataSheet = context.workbook.worksheets.getItemOrNullObject(sheetname);

    await context.sync();

    if (dataSheet.isNullObject) {
      dataSheet = context.workbook.worksheets.add(sheetname);
    }

    // Set `dataSheet` to be the second worksheet in the workbook.
    dataSheet.position = 0;
  });
}
async function settargetsheet_old() {
  await Excel.run(function(context) {
    var worksheets = context.workbook.worksheets;
    worksheets.load("name");
    return context.sync().then(async function() {
      var sheet_exist = false;
      for (var i = 0; i < worksheets.items.length; i++) {
        var sn = worksheets.items[i].name;
        if ((sn = navigation_sheet_name)) {
          sheet_exist = true;
          console.log(navigation_sheet_name + " exist");
          break;
        }
        //console.log(sn);
      }
      if (!sheet_exist) {
        let sheet = worksheets.add(navigation_sheet_name);
        await context.sync();
        console.log(navigation_sheet_name + " created!");
      }
    });
  });
}
