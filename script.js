Office.onReady((info) => {
  console.log("Office.js is now ready in ${info.host} host.");
  $("#initialisation").on("click", () => tryCatch(initialisation));
  $("#Start").on("click", () => tryCatch(start_Timer));
  $("#Stop").on("click", () => tryCatch(stop_Timer));
  $("#Pause").on("click", () => tryCatch(pause));
  $("#Reprendre").on("click", () => tryCatch(reprendre));
});
 
let tab_timers=[]
let timer = 0;
let time_spend_pause = 0;
let is_paused = false;
// This function is for refresh the select when the number of incident is modify.
function pause() {
  if (is_paused == false && timer != 0) 
  {
    var actualTime = new Date();
    time_spend_pause = actualTime.getTime() - timer;
    is_paused = true;
  } else console.log("oui")
}

function reprendre() {
  if (is_paused == true) {
    var actualTime = new Date();
    timer = actualTime.getTime() - time_spend_pause;
    is_paused = false;
  } else console.log("oui")
}

function displayTimer(tab_timers) {
  // Fund the the table by the id
  const table = document.getElementById('timersTable');

  // Effacer les lignes existantes
  table.innerHTML = '';

  for (const [domain, timer] of Object.entries(tab_timers)) {
    let row = table.insertRow();
    let cellDomain = row.insertCell(0);
    let cellStartTime = row.insertCell(1);

    cellDomain.innerHTML = domain;
    cellStartTime.innerHTML = timer;
  }
}

function stop_Timer() {
  const nniInput = document.getElementById("NNI");
  const nniValue = nniInput.value;

  const domaineInput = document.getElementById("ListDomaine");
  const domaineValue = domaineInput.value;

  var flag = 0;

  if (!nniValue) {
    console.log("oui")
    return;
  }

  if (!timer) {
    console.log("oui")
    return;
  }

  if (is_paused) {
    console.log("oui")
    return;
  }

  return Excel.run(function(context) {
    // DMT sheet **************************/
    var sheetDMT = context.workbook.worksheets.getItem("DMT");

    var usedRangesheetDMT = sheetDMT.getUsedRange();
    usedRangesheetDMT.load("rowCount");

    // DMT total sheet ********************/
    var sheetDMTTtl = context.workbook.worksheets.getItem("DMT Total");

    var usedRangeSheetDMTTtl = sheetDMTTtl.getUsedRange();
    usedRangeSheetDMTTtl.load("rowCount");

    // nb demande sheet ********************/
    var sheetNbD = context.workbook.worksheets.getItem("Nb demande");

    var usedRangeSheetNbD = sheetNbD.getUsedRange();
    usedRangeSheetNbD.load("rowCount");

    var headerRangeDMT = sheetDMT.getRange("A1:ZZ1");
    headerRangeDMT.load("values");

    var headerRangeDMTTtl = sheetDMTTtl.getRange("A1:ZZ1");
    headerRangeDMTTtl.load("values");

    var headerRangeNbD = sheetNbD.getRange("A1:ZZ1");
    headerRangeNbD.load("values");

    return context.sync().then(function() {
      var lastRowDMT = usedRangesheetDMT.rowCount;
      var actualDate = new Date().toLocaleDateString();

      var columnIndex = headerRangeDMT.values[0].indexOf(domaineValue);
      if (columnIndex === -1) {
        throw new Error("L'entête spécifié n'a pas été trouvé.");
      }

      var rangeNNI = sheetDMT.getRange("B2:B" + lastRowDMT);
      var rangeDate = sheetDMT.getRange("A2:A" + lastRowDMT);

      rangeNNI.load("values");
      rangeDate.load("values");

      return context.sync().then(function() {
        for (var i = 0; i < rangeDate.values.length; i++) {
          if (rangeDate.values[i][0] === actualDate && rangeNNI.values[i][0] === nniValue) {
            var goodRow = i + 1;
            flag = 1;
            break;
          }
        }

        if (flag == 1) {
          var idDomaineCellDMT = sheetDMT.getCell(goodRow, columnIndex);
          idDomaineCellDMT.load("values");

          var idDomaineCellDMTTtl = sheetDMTTtl.getCell(goodRow, columnIndex);
          idDomaineCellDMTTtl.load("values");

          var idDomaineCellNbD = sheetNbD.getCell(goodRow, columnIndex);
          idDomaineCellNbD.load("values");
          return context.sync().then(function() {
            var actualTime = new Date();
            idDomaineCellDMT.values = [[idDomaineCellDMT.values[0][0] + (actualTime.getTime() - timer) / 1000]];
            idDomaineCellNbD.values = [[idDomaineCellNbD.values[0][0] + 1]];
            idDomaineCellDMTTtl.values = [[idDomaineCellDMT.values[0][0] / idDomaineCellNbD.values[0][0]]];
            timer = 0;
            tab_timers[domaineValue] = idDomaineCellDMTTtl.values[0][0]
            displayTimer(tab_timers)
            return context.sync();
          });
        } else {
          console.log("oui");
          return context.sync();
        }
      });
    });
  });
}

function getDomaine() {
  var select = document.getElementById("ListDomaine");
  
  var domaines = {};

  return Excel.run(function(context) {
    // We get the sheet Save for doing operation on it.
    var params = context.workbook.worksheets.getItem("Params");
    var usedRange = params.getUsedRange(true);
    usedRange.load("rowCount");

    return context.sync().then(function() {
      var lastRow = usedRange.rowCount;
      var range = params.getRange("A2:A" + lastRow);
      range.load("values"); // Charger les valeurs

      return context.sync().then(function() {
        var values = range.values;
        for (var i = 0; i < values.length; i++) {
          let key = values[i][0];
          domaines[key] = 0;
        }
        select.innerHTML = "";
        Object.keys(domaines).forEach(function(option) {
          var el = document.createElement("option");
          el.textContent = option;
          el.value = option;
          select.appendChild(el);
        });
      });
    });
  });
}

// For adding an incident
function initialisation() {
  const nniInput = document.getElementById("NNI");
  const nniValue = nniInput.value;
  getDomaine();
  $("#ListDomaine").select2({
    placeholder: "Select an option",
    width: "100%"
  });
  var flag = 0;

  if (nniValue) {
    return Excel.run(function(context) {
      // sheetDMT initialisation ***************************************************/
      var sheetDMT = context.workbook.worksheets.getItem("DMT");

      // Load all row for check where we are going to put the value
      var usedRangesheetDMT = sheetDMT.getUsedRange();
      usedRangesheetDMT.load("rowCount");

      // nb demande sheet initialisation ***************************************************/

      var sheetNbD = context.workbook.worksheets.getItem("Nb demande");

      // Load all row for check where we are going to put the value
      var usedRangeSheetNbD = sheetNbD.getUsedRange();
      usedRangeSheetNbD.load("rowCount");

      // DMT Total initialisation ***************************************************/

      var sheetDMTTtl = context.workbook.worksheets.getItem("DMT Total");

      // Load all row for check where we are going to put the value
      var usedRangeSheetDMTTtl = sheetDMTTtl.getUsedRange();
      usedRangeSheetDMTTtl.load("rowCount");

      return context.sync().then(function() {
        //DMT Total sheet range initialisation ***************************************************/

        var lastRowDMT = usedRangeSheetDMTTtl.rowCount; // The last row used in 'Suivi'
        var rangeDateDMT = sheetDMTTtl.getRange("A2:A" + lastRowDMT);
        rangeDateDMT.load("values");
        var rangeNNIDMT = sheetDMTTtl.getRange("B2:B" + lastRowDMT);
        rangeNNIDMT.load("values");

        // nb Demande range initialisation  ***************************************************/

        var lastRowNbD = usedRangeSheetNbD.rowCount; // The last row used in 'Suivi'
        var rangeDateNbD = sheetNbD.getRange("A2:A" + lastRowNbD);
        rangeDateNbD.load("values");
        var rangeNNINbD = sheetNbD.getRange("B2:B" + lastRowNbD);
        rangeNNINbD.load("values");

        // DMT range initialisation ***************************************************/

        var lastRowNniSuivi = usedRangesheetDMT.rowCount; // The last row used in 'Suivi'
        var rangeDate = sheetDMT.getRange("A2:A" + lastRowNniSuivi);
        rangeDate.load("values");
        var rangeNNI = sheetDMT.getRange("B2:B" + lastRowNniSuivi);
        rangeNNI.load("values");

        return context.sync().then(function() {
          // DMT cells
          var nniCellNni = sheetDMT.getCell(lastRowNniSuivi, 1);
          var idDateCell = sheetDMT.getCell(lastRowNniSuivi, 0);

          // NbD cells
          var nniCellNniNbD = sheetNbD.getCell(lastRowNbD, 1);
          var idDateCellNbD = sheetNbD.getCell(lastRowNbD, 0);

          // DMT Total cells
          var nniCellNniDMT = sheetDMTTtl.getCell(lastRowDMT, 1);
          var idDateCellDMT = sheetDMTTtl.getCell(lastRowDMT, 0);

          var actualDate = new Date();
          var valuesDate = rangeDate.values;
          var valuesNNI = rangeNNI.values;
          for (var i = 0; i < valuesDate.length; i++) {
            if (valuesDate[i][0] === actualDate.toLocaleDateString() && valuesNNI[i][0] === nniValue) {
              flag = 1;
            }
          }

          if (flag == 0) {
            nniCellNni.values = [[nniValue]];
            idDateCell.values = [[actualDate.toLocaleDateString()]];

            nniCellNniNbD.values = [[nniValue]];
            idDateCellNbD.values = [[actualDate.toLocaleDateString()]];

            nniCellNniDMT.values = [[nniValue]];
            idDateCellDMT.values = [[actualDate.toLocaleDateString()]];

            console.log("oui")
          } else {
            console.log("oui")
          }

          return context.sync();
        });
      });
    });
  } else {
    console.log("oui")
  }
}

//Start the timer
function start_Timer() {
  if (timer == 0) {
    var actualDate = new Date();
    timer = actualDate.getTime();
  } else console.log("oui")
}
/** Default helper for invoking an action and handling errors. */
function tryCatch(callback) {
  Promise.resolve()
    .then(callback)
    .catch(function(error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
    });
}
