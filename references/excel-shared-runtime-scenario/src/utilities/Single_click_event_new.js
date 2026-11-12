
const officeguid = uuidv4();
const crossorigin = "https://ruhchjdbdbhccjjc.github.io";
//const crossorigin = "https://script-lab-runner.azureedge.net";
//const crossorigin = "*";
//const crossorigin = "https://ruhchjdbdbhccjjc.github.io/Office-Add-in-samples/Samples/excel-shared-runtime-scenario";

const homename = "导航";
var filename = "";
loadFileName();
var isoncommand = false;
var officecommandwaitforruncollection = [];
var officecommandfinisedruncollection = [];
function uuidv4() {
  return ([1e7] + -1e3 + -4e3 + -8e3 + -1e11).replace(/[018]/g, (c) =>
    (c ^ (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (c / 4)))).toString(16)
  );
}

async function loadFileName() {
  return new Promise((resolve) => {
    Office.context.document.getFilePropertiesAsync(null, (res) => {
      if (res && res.value && res.value.url) {
        let name = res.value.url.substr(res.value.url.lastIndexOf("/") + 1);
        filename = name;
        //return name;
        //filenameood = name;

        resolve(name);
      }
      resolve("");
    });
  });
}

var postreturncommandjson = {
  postreturncommandjson: "",
  officeguid: officeguid,
  officecommand: officecommand,
  crossorigin: crossorigin
};
var cmdjson = {
  commandguid: "",
  setboard: false,
  addarrow: false,
  insertrow: false,
  insertcoloumn: false,
  gonagivetion: false,

  setnagition: false,

  recoderange: "",
  readrecoderange: "",
  recodesheetrange: "",
  readsheetrange: "",

  rangeprevios: false,
  rangesheetprevio: false,
  rangepreviosindex: false,
  rangesheetpreviosindex: false,
  create_sheet_with_name: false,
  resetpreviosindex: false,
  resetrecodeinfomation: false,
  openhyperlink: false,
  result: ""
};

var cmdjson_setboard = false;
Object.defineProperty(cmdjson, "setboard", {
  set: async function(newAge) {
    cmdjson_setboard = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await setborder();
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_setboard;
    //return this.age;
  }
});
var cmdjson_addarrow = false;
Object.defineProperty(cmdjson, "addarrow", {
  set: async function(newAge) {
    cmdjson_addarrow = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await addarrowLine();
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_addarrow;
    //return this.age;
  }
});
var cmdjson_insertrow = false;
Object.defineProperty(cmdjson, "insertrow", {
  set: async function(newAge) {
    cmdjson_insertrow = newAge;
    console.log(this.commandguid + " : " + newAge);
    if (newAge != true) return;
    isoncommand = true;
    // Shows all indexes, including deleted
    //await arrowLine();
    await insertrow();
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_insertrow;
    //return this.age;
  }
});
var cmdjson_insertcoloumn = false;
Object.defineProperty(cmdjson, "insertcoloumn", {
  set: async function(newAge) {
    cmdjson_insertcoloumn = newAge;
    console.log(this.commandguid + " : " + newAge);
    if (newAge != true) return;
    isoncommand = true;
    // Shows all indexes, including deleted
    //await arrowLine();
    await insertcoloumn();
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_insertcoloumn;
    //return this.age;
  }
});

//2024/03/22 17:39:06 add go to nagivetion
var cmdjson_gonagivetion = false;
Object.defineProperty(cmdjson, "gonagivetion", {
  set: async function(newAge) {
    cmdjson_gonagivetion = newAge;
    console.log(this.commandguid + " : " + newAge);
    if (newAge != true) return;
    isoncommand = true;
    // Shows all indexes, including deleted
    //await arrowLine();
    await gonagivetion();
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_gonagivetion;
    //return this.age;
  }
});

var cmdjson_setnagition = false;
Object.defineProperty(cmdjson, "setnagition", {
  set: async function(newAge) {
    cmdjson_setnagition = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await setnagition();
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_setnagition;
    //return this.age;
  }
});

var cmdjson_recoderange = "";
Object.defineProperty(cmdjson, "recoderange", {
  set: async function(newAge) {
    cmdjson_recoderange = newAge;
    if (newAge == null) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    console.log("recoderange start recode ! ");
    await recoderange(recodejsonname, "A1", newAge);
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_recoderange;
    //return this.age;
  }
});

var cmdjson_readrecoderange = "";
Object.defineProperty(cmdjson, "readrecoderange", {
  set: async function(newAge) {
    cmdjson_readrecoderange = newAge;
    if (newAge == null) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    await readrecoderange(recodejsonname, "A1", newAge);
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_readrecoderange;
    //return this.age;
  }
});

var cmdjson_recodesheetrange = "";
Object.defineProperty(cmdjson, "recodesheetrange", {
  set: async function(newAge) {
    cmdjson_recodesheetrange = newAge;
    if (newAge == null) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    await recodesheetrange(recodejsonname, "A1", newAge);
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_recodesheetrange;
    //return this.age;
  }
});

var cmdjson_readsheetrange = "";
Object.defineProperty(cmdjson, "readsheetrange", {
  set: async function(newAge) {
    cmdjson_readsheetrange = newAge;
    if (newAge == null) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await readsheetrange(recodejsonname, "A1", newAge);
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_readsheetrange;
    //return this.age;
  }
});
var cmdjson_rangeprevios = false;
Object.defineProperty(cmdjson, "rangeprevios", {
  set: async function(newAge) {
    cmdjson_rangeprevios = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await readrecodeworkbookselection(recodeselectionjsonname, "C1", 0);
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_rangeprevios;
    //return this.age;
  }
});

var cmdjson_rangesheetprevio = false;
Object.defineProperty(cmdjson, "rangesheetprevio", {
  set: async function(newAge) {
    cmdjson_rangesheetprevio = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await readsheetselection(recodeselectionjsonname, "C1", 0);
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_rangesheetprevio;
    //return this.age;
  }
});

var cmdjson_rangepreviosindex = false;
Object.defineProperty(cmdjson, "rangepreviosindex", {
  set: async function(newAge) {
    cmdjson_rangepreviosindex = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await readrecodeworkbookselection(recodeselectionjsonname, "C1", previousindex);
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_rangepreviosindex;
    //return this.age;
  }
});

var cmdjson_rangesheetpreviosindex = false;
Object.defineProperty(cmdjson, "rangesheetpreviosindex", {
  set: async function(newAge) {
    cmdjson_rangesheetpreviosindex = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await readsheetselection(recodeselectionjsonname, "C1", previousindex);
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_rangesheetpreviosindex;
    //return this.age;
  }
});

var cmdjson_create_sheet_with_name = false;
Object.defineProperty(cmdjson, "create_sheet_with_name", {
  set: async function(newAge) {
    cmdjson_create_sheet_with_name = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await createsheetswithselectvalues();
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_create_sheet_with_name;
    //return this.age;
  }
});

var cmdjson_resetpreviosindex = false;
Object.defineProperty(cmdjson, "resetpreviosindex", {
  set: async function(newAge) {
    cmdjson_resetpreviosindex = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await resetpreviosindex();
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_resetpreviosindex;
    //return this.age;
  }
});
var cmdjson_resetrecodeinfomation = false;
Object.defineProperty(cmdjson, "resetrecodeinfomation", {
  set: async function(newAge) {
    cmdjson_resetrecodeinfomation = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await resetrecodeinfomation();
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_resetrecodeinfomation;
    //return this.age;
  }
});

var cmdjson_openhyperlink = false;
Object.defineProperty(cmdjson, "openhyperlink", {
  set: async function(newAge) {
    cmdjson_openhyperlink = newAge;
    if (newAge != true) return;
    isoncommand = true;
    console.log(this.commandguid + " : " + newAge);
    // Shows all indexes, including deleted
    //await arrowLine();
    await openhyperlink();
    var returncommand = officecommandfinisedruncollection.find((value, index) => {
      var result = null;
      // Delete element 5 on first iteration
      if (value.commandjson.commandguid == this.commandguid) {
        console.log("finded return command :", JSON.stringify(value));
        var newcommand = value;
        newcommand.commandjson.isfinised = true;
        var finallycommand = postreturncommandjson;
        finallycommand.officecommand = newcommand;
        console.log("postreturncommand finallycommand " + JSON.stringify(finallycommand));
        // need to postreturn command here .else will failed !
        postreturncommand(JSON.stringify(finallycommand));
        isoncommand = false;
        return JSON.stringify(finallycommand);
        //result = finallycommand;
        //return newcommand;
        //console.log("postreturncommand" + (returncommand));
      }
      // Element 5 is still visited even though deleted
      //console.log("Visited index" + index + "with value", JSON.stringify(value));
      return result;
    });
  },
  get: function() {
    return cmdjson_openhyperlink;
    //return this.age;
  }
});

var commandjson = {
  commandguid: "",
  isfinised: false,
  cmdjson: cmdjson
};
var officecommand = {
  officeinstance: officeguid,
  commandjson: commandjson
};
var getcommandjson = {
  getcommandjson: "",
  officeguid: officeguid,
  officecommand: officecommand,
  crossorigin: crossorigin
};
async function postinstance() {
  var officeinstancejson = {
    officeinstanceguid: officeguid,
    officetype: filename,
    crossorigin: crossorigin
  };
  console.log("postinstance ： " + JSON.stringify(officeinstancejson));
  // Make a request for a user with a given ID
  
  axios({
    method: "post",
    url: "http://localhost:8080/api",
    headers: {
      "Content-Type": "multipart/form-data"
    },
    data: JSON.stringify(officeinstancejson)
  }).then(function(response) {
      // handle success
      console.log("postinstace recived : " + JSON.stringify(response.data));
    })
    .catch(function(error) {
      // handle error
      console.log("postinstace never recived : " + error);
      console.log(error);
    })
    .finally(function() {
      // always executed
    });
}
async function postreturncommand(jsoncommadnew) {
  // Make a request for a user with a given ID
  
  axios({
    method: "post",
    url: "http://localhost:8080/api",
    headers: {
      "Content-Type": "multipart/form-data"
    },
    data: jsoncommadnew
  }).then(function(response) {
      // handle success
      console.log("post return command recived : " + JSON.stringify(response.data));
    })
    .catch(function(error) {
      // handle error
      console.log("post return  command never recived : " + error);
      console.log(error);
    })
    .finally(function() {
      // always executed
    });
}

async function postcommand(jsoncommad) {
  // Make a request for a user with a given ID
  
  axios({
    method: "post",
    url: "http://localhost:8080/api",
    headers: {
      "Content-Type": "multipart/form-data"
    },
    data: jsoncommad
  }).then(function(response) {
      // handle success
      console.log("postcommand recived : " + JSON.stringify(response.data));
    })
    .catch(function(error) {
      // handle error
      console.log("post command never recived : " + error + crossorigin);
      console.log(error);
    })
    .finally(function() {
      // always executed
    });
}

function finisedcheck(checkstring) {
  var finised = false;
  let check = officecommandfinisedruncollection.find((value, index) => {
    var checkis = false;
    // Delete element 5 on first iteration
    var cmdguid = value.commandjson.commandguid;
    if (cmdguid == null) return;
    if (cmdguid == checkstring) {
      console.log(" finisedcheck finded command :" + JSON.stringify(value));

      finised = true;

      checkis = true;

      //delete officecommandwaitforruncollection[index];
    }
    // Element 5 is still visited even though deleted
    console.log(" isneedtorun  Visited index" + index + "with value" + JSON.stringify(value));
    return checkis;
  });

  return finised;
}
function isneedtorun(guidstring) {
  var is = false;

  // Shows all indexes, including deleted
  console.log("  isneedtorun with value" + JSON.stringify(officecommandwaitforruncollection));

  let check = officecommandwaitforruncollection.find((value, index) => {
    var checkis = false;
    // Delete element 5 on first iteration
    var cmdguid = value.commandjson.commandguid;
    if (cmdguid == null) return;
    var dontfinised = finisedcheck(cmdguid);
    if (cmdguid == guidstring && dontfinised == false) {
      console.log(" isneedtorun finded command :" + JSON.stringify(value));
      var newcommand = value;
      newcommand.commandjson.isfinised = true;
      officecommandfinisedruncollection.push(newcommand);
      is = true;
      checkis = true;
      officecommandwaitforruncollection.splice(index, 1);
      //delete officecommandwaitforruncollection[index];
    }
    // Element 5 is still visited even though deleted
    console.log(" isneedtorun  Visited index" + index + "with value" + JSON.stringify(value));
    return checkis;
  });

  console.log("is " + is);
  return is;
}

async function getcommand() {
  if (isoncommand == true) return;
  console.log("getcommand send : " + JSON.stringify(getcommandjson));
  // Make a request for a user with a given ID
  
  axios({
    method: "post",
    url: "http://localhost:8080/api",
    headers: {
      "Content-Type": "multipart/form-data"
    },
    data: JSON.stringify(getcommandjson)
  }).then(function(response) {
      // handle success
      var resopnsecommand = response.data;
      var resopnsecommandjson = response.data.officecommand;
      console.log("getcommand recived : " + JSON.stringify(response.data));
      if (
        !JSON.stringify(resopnsecommand).includes("getcommandjson") ||
        resopnsecommandjson.commandjson.commandguid == null
      )
        return;
      officecommandwaitforruncollection.push(resopnsecommandjson);
      console.log("check if need to run !");

      console.log(resopnsecommandjson.officeinstance + resopnsecommandjson.commandjson.commandguid);
      var ismine = resopnsecommandjson.officeinstance == officeguid;

      var mm = isneedtorun(resopnsecommandjson.commandjson.commandguid);
      var isneed = mm;
      console.log("is mine " + ismine + "isneeed " + isneed);
      if (ismine && isneed) {
        console.log("need to run !");
        const returnedTarget = Object.assign(cmdjson, resopnsecommandjson.commandjson.cmdjson);
      }
    })
    .catch(function(error) {
      // handle error
      console.log("get command never recived : " + error);
      console.log(error);
    })
    .finally(function() {
      // always executed
    });
}
async function registerClickHandler() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets;
    //context.workbook.onSelectionChanged
    sheet.onSingleClicked.add((event) => {
      return Excel.run(async (context) => {
        console.log("file  name " + filename);
        await postinstance();
        await setsharpposition();
        /*
        console.log(
          `Click detected at ${event.address} (pixel offset from upper-left cell corner: ${event.offsetX}, ${event.offsetY})`
        );
        */
        //await getcommand();
        return context.sync();
      });
    });

    //context.workbook.onSelectionChanged
    //sheet.onSingleClicked.add((event) => {

    context.workbook.onSelectionChanged.add((event) => {
      return Excel.run(async (context) => {
        console.log("onSelectionChanged " + event.workbook);

        await postinstance();
        await setsharpposition();

        /*
        //console.log(
          `Click detected at ${event.address} (pixel offset from upper-left cell corner: ${event.offsetX}, ${event.offsetY})`
        );
        */
        //await getcommand();
        return context.sync();
      });
    });

    console.log("The worksheet click handler is registered.");

    await context.sync();
  });
}
registerClickHandler();

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

var shapepostion = [];
async function registerselectionchangedHandler() {
  await Excel.run(async (context) => {
    // workbook shelection changed don't work only workshhetschaged !
    const sheet = context.workbook.worksheets;

    sheet.onSelectionChanged.add((event) => {
      return Excel.run(async (context) => {
        console.log(`selection  changed to ${event})`);
        //await postinstance();
        //await setsharpposition();
        return context.sync();
      });
    });

    console.log("The worksheet selection changed  handler is registered.");

    await context.sync();
  });
}

async function setsharpposition() {
  //https://stackoverflow.com/questions/62766879/columnwidth-of-rangeformat-is-giving-some-random-number-instead-giving-the-actua
  await Excel.run(async (context) => {
    console.log(" setsharpposition() ");
    var activecell = context.workbook.getActiveCell();

    activecell.load("address,addressLocal ");
    await context.sync();
    console.log("activecell address " + activecell.address + "activecell.addressLocal " + activecell.addressLocal);
    console.log(activecell.address.substring(activecell.address.indexOf("!") + 1));
    var position = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange("A1:" + activecell.address.substring(activecell.address.indexOf("!") + 1));

    position.load("height,width");
    await context.sync();
    //console.log("position.format.columnWidth" + position.format.columnWidth + 'position.format.rowHeight' + position.format.rowHeight + 'height  ' + position.height +'width ' + position.width);
    shapepostion.push([position.height, position.width]);
    await context.sync();
  });
}
//registerselectionchangedHandler();
async function addarrowLine() {
  console.log("start add arrowline ");
  //https://stackoverflow.com/questions/62294200/is-there-a-way-to-change-the-color-of-a-line-shape-in-an-office-js-excel-add-in

  await Excel.run(async (context) => {
    console.log("get activeworksheet shapes ");
    const shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
    var startx = shapepostion[shapepostion.length - 2][1];
    var stary = shapepostion[shapepostion.length - 2][0];

    var endx = shapepostion[shapepostion.length - 1][1];
    var endy = shapepostion[shapepostion.length - 1][0];
    const lineshape = shapes.addLine(startx, stary, endx, endy, Excel.ConnectorType.straight);
    lineshape.lineFormat.color = "red";
    const line = lineshape.line;
    line.beginArrowheadLength = Excel.ArrowheadLength.long;
    line.beginArrowheadWidth = Excel.ArrowheadWidth.wide;
    line.beginArrowheadStyle = Excel.ArrowheadStyle.oval;

    line.endArrowheadLength = Excel.ArrowheadLength.long;
    line.endArrowheadWidth = Excel.ArrowheadWidth.wide;
    line.endArrowheadStyle = Excel.ArrowheadStyle.triangle;

    await context.sync();
  });
}
async function setborder() {
  console.log("set border ");
  //const context = new Excel.RequestContext();
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const rangeFormat = range.format;
    rangeFormat.fill.load();
    rangeFormat.borders.load();
    var rangeborders = rangeFormat.borders;

    const colors = ["#E26D5C", "#FFFFFF", "#C7CC7A", "#7560BA", "#9DD9D2", "#FFE1A8"];

    return context.sync().then(function() {
      const rangeTarget = context.workbook.getSelectedRange();
      let currentColor = -1;
      for (let i = 0; i < colors.length; i++) {
        if (colors[i] == rangeborders.items[0].color) {
          currentColor = i;
          break;
        }
      }
      if (currentColor == -1) {
        currentColor = 0;
      } else if (currentColor == colors.length - 1) {
        currentColor = 0;
      } else {
        currentColor++;
      }
      //alert(colors[currentColor]);
      //rangeTarget.format.fill.color = colors[currentColor];
      rangeTarget.format.borders.getItem("EdgeBottom").style = "Continuous";
      rangeTarget.format.borders.getItem("EdgeLeft").style = "Continuous";
      rangeTarget.format.borders.getItem("EdgeRight").style = "Continuous";
      rangeTarget.format.borders.getItem("EdgeTop").style = "Continuous";

      //rangeTarget.format.borders.getItem('InsideHorizontal').weight = "Thick";
      //rangeTarget.format.borders.getItem('InsideVertical').weight = "Thick";
      rangeTarget.format.borders.getItem("EdgeBottom").weight = "Thick";
      rangeTarget.format.borders.getItem("EdgeLeft").weight = "Thick";
      rangeTarget.format.borders.getItem("EdgeRight").weight = "Thick";
      rangeTarget.format.borders.getItem("EdgeTop").weight = "Thick";

      //rangeTarget.format.borders.getItem('InsideVertical').color = colors[currentColor];
      rangeTarget.format.borders.getItem("EdgeBottom").color = colors[currentColor];
      rangeTarget.format.borders.getItem("EdgeLeft").color = colors[currentColor];
      rangeTarget.format.borders.getItem("EdgeRight").color = colors[currentColor];
      rangeTarget.format.borders.getItem("EdgeTop").color = colors[currentColor];
      //rangeTarget.getEntireRow().insert(Excel.InsertShiftDirection.down);
      //rangeTarget.getEntireColumn().insert(Excel.InsertShiftDirection.right);
    });
  });
}
async function insertrow() {
  await Excel.run(async (context) => {
    //const range = context.workbook.getSelectedRange();
    return context.sync().then(function() {
      const rangeTarget = context.workbook.getSelectedRange();
      rangeTarget.getEntireRow().insert(Excel.InsertShiftDirection.down);
    });
  });
}
async function insertcoloumn() {
  await Excel.run(async (context) => {
    //const range = context.workbook.getSelectedRange();
    return context.sync().then(function() {
      const rangeTarget = context.workbook.getSelectedRange();
      rangeTarget.getEntireColumn().insert(Excel.InsertShiftDirection.right);
    });
  });
}

async function gonagivetion() {
  await active_sheets(homename);
}

async function active_sheets(sheetname) {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem(sheetname);
    sheet.activate();
    sheet.load("name");
    await context.sync();
    console.log(`The active worksheet is "${sheet.name}"`);

    /*
    let recodeSheet = context.workbook.worksheets.getItemOrNullObject(sheetname);
    //recodeSheet.activate();

    await context.sync();
    if (!recodeSheet.isNullObject) {
      recodeSheet.activate();
     // recodeSheet.getRange("A1").select();
      //console.log(`selected recode sheetrange ${codeinfomation.address}`);
    }
    */
  });
}

//var intervalID = setInterval(getcommand, 250);
//var intervalID = setInterval(getcommand, 250);
//2024/04/21 18:06:58 ,trying to keep http keep alive 
const domain = "http://localhost:8080/api";
let axios_instance;

function create_instance(){
    if (!axios_instance)
    {
        //create axios instance
        axios_instance = axios.create({
            baseURL: domain,
            timeout: 600000000000000000000000000000000000000000000, //optional
            httpsAgent: { keepAlive: true },
            headers: {'Content-Type':'multipart/form-data'}
        })
    }

    return axios_instance;
}
create_instance();
