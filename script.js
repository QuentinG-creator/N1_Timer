
// Office.initialize = function (reason)
// {
//   getDomaine();
//   $("#ListDomaine").select2({
//     placeholder: "Select an option",
//     width: "100%"
//   });
//   $("#initialisation").on("click", () => tryCatch(initialisation));
//   $("#Start").on("click", () => tryCatch(start_Timer));
//   $("#Stop").on("click", () => tryCatch(stop_Timer));
//   $("#Pause").on("click", () => tryCatch(pause));
//   $("#Reprendre").on("click", () => tryCatch(reprendre));
// }
 
// let tab_timers=[]
// let timer = 0;
// let time_spend_pause = 0;
// let is_paused = false;
// // This function is for refresh the select when the number of incident is modify.
// function pause() {
//   if (is_paused == false && timer != 0) {
//     var actualTime = new Date();
//     time_spend_pause = actualTime.getTime() - timer;
//     is_paused = true;
//   } else alert("Votre timer est déjà en pause ou alors vous n'avez pas démarrer de timer");
// }

// function reprendre() {
//   if (is_paused == true) {
//     var actualTime = new Date();
//     timer = actualTime.getTime() - time_spend_pause;
//     is_paused = false;
//   } else alert("Votre timer n'est pas en pause");
// }

// function displayTimer(tab_timers) {
//   // Fund the the table by the id
//   const table = document.getElementById('timersTable');

//   // Effacer les lignes existantes
//   table.innerHTML = '';

//   for (const [domain, timer] of Object.entries(tab_timers)) {
//     let row = table.insertRow();
//     let cellDomain = row.insertCell(0);
//     let cellStartTime = row.insertCell(1);

//     cellDomain.innerHTML = domain;
//     cellStartTime.innerHTML = timer;
//   }
// }

// function stop_Timer() {
//   const nniInput = document.getElementById("NNI");
//   const nniValue = nniInput.value;

//   const domaineInput = document.getElementById("ListDomaine");
//   const domaineValue = domaineInput.value;

//   var flag = 0;

//   if (!nniValue) {
//     alert("! Veuillez entrer votre NNI !");
//     return;
//   }

//   if (!timer) {
//     alert("! Le timer n'est pas lancé !");
//     return;
//   }

//   if (is_paused) {
//     alert("! Le timer est en pause !");
//     return;
//   }

//   return Excel.run(function(context) {
//     // DMT sheet **************************/
//     var sheetDMT = context.workbook.worksheets.getItem("DMT");

//     var usedRangesheetDMT = sheetDMT.getUsedRange();
//     usedRangesheetDMT.load("rowCount");

//     // DMT total sheet ********************/
//     var sheetDMTTtl = context.workbook.worksheets.getItem("DMT Total");

//     var usedRangeSheetDMTTtl = sheetDMTTtl.getUsedRange();
//     usedRangeSheetDMTTtl.load("rowCount");

//     // nb demande sheet ********************/
//     var sheetNbD = context.workbook.worksheets.getItem("Nb demande");

//     var usedRangeSheetNbD = sheetNbD.getUsedRange();
//     usedRangeSheetNbD.load("rowCount");

//     var headerRangeDMT = sheetDMT.getRange("A1:ZZ1");
//     headerRangeDMT.load("values");

//     var headerRangeDMTTtl = sheetDMTTtl.getRange("A1:ZZ1");
//     headerRangeDMTTtl.load("values");

//     var headerRangeNbD = sheetNbD.getRange("A1:ZZ1");
//     headerRangeNbD.load("values");

//     return context.sync().then(function() {
//       var lastRowDMT = usedRangesheetDMT.rowCount;
//       var actualDate = new Date().toLocaleDateString();

//       var columnIndex = headerRangeDMT.values[0].indexOf(domaineValue);
//       if (columnIndex === -1) {
//         throw new Error("L'entête spécifié n'a pas été trouvé.");
//       }

//       var rangeNNI = sheetDMT.getRange("B2:B" + lastRowDMT);
//       var rangeDate = sheetDMT.getRange("A2:A" + lastRowDMT);

//       rangeNNI.load("values");
//       rangeDate.load("values");

//       return context.sync().then(function() {
//         for (var i = 0; i < rangeDate.values.length; i++) {
//           if (rangeDate.values[i][0] === actualDate && rangeNNI.values[i][0] === nniValue) {
//             var goodRow = i + 1;
//             flag = 1;
//             break;
//           }
//         }

//         if (flag == 1) {
//           var idDomaineCellDMT = sheetDMT.getCell(goodRow, columnIndex);
//           idDomaineCellDMT.load("values");

//           var idDomaineCellDMTTtl = sheetDMTTtl.getCell(goodRow, columnIndex);
//           idDomaineCellDMTTtl.load("values");

//           var idDomaineCellNbD = sheetNbD.getCell(goodRow, columnIndex);
//           idDomaineCellNbD.load("values");
//           return context.sync().then(function() {
//             var actualTime = new Date();
//             idDomaineCellDMT.values = [[idDomaineCellDMT.values[0][0] + (actualTime.getTime() - timer) / 1000]];
//             idDomaineCellNbD.values = [[idDomaineCellNbD.values[0][0] + 1]];
//             idDomaineCellDMTTtl.values = [[idDomaineCellDMT.values[0][0] / idDomaineCellNbD.values[0][0]]];
//             timer = 0;
//             tab_timers[domaineValue] = idDomaineCellDMTTtl.values[0][0]
//             displayTimer(tab_timers)
//             return context.sync();
//           });
//         } else {
//           alert("! Veuillez appuyer sur initialiser avec votre NNI avant tout autre action !");
//           return context.sync();
//         }
//       });
//     });
//   });
// }

// function getDomaine() {
//   var select = document.getElementById("ListDomaine");

//   var domaines = {};

//   return Excel.run(function(context) {
//     // We get the sheet Save for doing operation on it.
//     var params = context.workbook.worksheets.getItem("Params");
//     var usedRange = params.getUsedRange(true);
//     usedRange.load("rowCount");

//     return context.sync().then(function() {
//       var lastRow = usedRange.rowCount;
//       var range = params.getRange("A2:A" + lastRow);
//       range.load("values"); // Charger les valeurs

//       return context.sync().then(function() {
//         var values = range.values;
//         for (var i = 0; i < values.length; i++) {
//           let key = values[i][0];
//           domaines[key] = 0;
//         }
//         select.innerHTML = "";
//         Object.keys(domaines).forEach(function(option) {
//           var el = document.createElement("option");
//           el.textContent = option;
//           el.value = option;
//           select.appendChild(el);
//         });
//       });
//     });
//   });
// }

// // For adding an incident
// function initialisation() {
//   const nniInput = document.getElementById("NNI");
//   const nniValue = nniInput.value;

//   var flag = 0;

//   if (nniValue) {
//     return Excel.run(function(context) {
//       // sheetDMT initialisation ***************************************************/
//       var sheetDMT = context.workbook.worksheets.getItem("DMT");

//       // Load all row for check where we are going to put the value
//       var usedRangesheetDMT = sheetDMT.getUsedRange();
//       usedRangesheetDMT.load("rowCount");

//       // nb demande sheet initialisation ***************************************************/

//       var sheetNbD = context.workbook.worksheets.getItem("Nb demande");

//       // Load all row for check where we are going to put the value
//       var usedRangeSheetNbD = sheetNbD.getUsedRange();
//       usedRangeSheetNbD.load("rowCount");

//       // DMT Total initialisation ***************************************************/

//       var sheetDMTTtl = context.workbook.worksheets.getItem("DMT Total");

//       // Load all row for check where we are going to put the value
//       var usedRangeSheetDMTTtl = sheetDMTTtl.getUsedRange();
//       usedRangeSheetDMTTtl.load("rowCount");

//       return context.sync().then(function() {
//         //DMT Total sheet range initialisation ***************************************************/

//         var lastRowDMT = usedRangeSheetDMTTtl.rowCount; // The last row used in 'Suivi'
//         var rangeDateDMT = sheetDMTTtl.getRange("A2:A" + lastRowDMT);
//         rangeDateDMT.load("values");
//         var rangeNNIDMT = sheetDMTTtl.getRange("B2:B" + lastRowDMT);
//         rangeNNIDMT.load("values");

//         // nb Demande range initialisation  ***************************************************/

//         var lastRowNbD = usedRangeSheetNbD.rowCount; // The last row used in 'Suivi'
//         var rangeDateNbD = sheetNbD.getRange("A2:A" + lastRowNbD);
//         rangeDateNbD.load("values");
//         var rangeNNINbD = sheetNbD.getRange("B2:B" + lastRowNbD);
//         rangeNNINbD.load("values");

//         // DMT range initialisation ***************************************************/

//         var lastRowNniSuivi = usedRangesheetDMT.rowCount; // The last row used in 'Suivi'
//         var rangeDate = sheetDMT.getRange("A2:A" + lastRowNniSuivi);
//         rangeDate.load("values");
//         var rangeNNI = sheetDMT.getRange("B2:B" + lastRowNniSuivi);
//         rangeNNI.load("values");

//         return context.sync().then(function() {
//           // DMT cells
//           var nniCellNni = sheetDMT.getCell(lastRowNniSuivi, 1);
//           var idDateCell = sheetDMT.getCell(lastRowNniSuivi, 0);

//           // NbD cells
//           var nniCellNniNbD = sheetNbD.getCell(lastRowNbD, 1);
//           var idDateCellNbD = sheetNbD.getCell(lastRowNbD, 0);

//           // DMT Total cells
//           var nniCellNniDMT = sheetDMTTtl.getCell(lastRowDMT, 1);
//           var idDateCellDMT = sheetDMTTtl.getCell(lastRowDMT, 0);

//           var actualDate = new Date();
//           var valuesDate = rangeDate.values;
//           var valuesNNI = rangeNNI.values;
//           for (var i = 0; i < valuesDate.length; i++) {
//             if (valuesDate[i][0] === actualDate.toLocaleDateString() && valuesNNI[i][0] === nniValue) {
//               flag = 1;
//             }
//           }

//           if (flag == 0) {
//             nniCellNni.values = [[nniValue]];
//             idDateCell.values = [[actualDate.toLocaleDateString()]];

//             nniCellNniNbD.values = [[nniValue]];
//             idDateCellNbD.values = [[actualDate.toLocaleDateString()]];

//             nniCellNniDMT.values = [[nniValue]];
//             idDateCellDMT.values = [[actualDate.toLocaleDateString()]];

//             alert("Vous avez bien initialiser");
//           } else {
//             alert("! Vous avez déjà initialiser votre NNI pour aujourd'hui !");
//           }

//           return context.sync();
//         });
//       });
//     });
//   } else {
//     alert("! Veuillez entrer votre NNI !");
//   }
// }

// //Start the timer
// function start_Timer() {
//   if (timer == 0) {
//     var actualDate = new Date();
//     timer = actualDate.getTime();
//   } else alert("vous avez déjà un incident en cours");
// }
// /** Default helper for invoking an action and handling errors. */
// function tryCatch(callback) {
//   Promise.resolve()
//     .then(callback)
//     .catch(function(error) {
//       // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
//       console.error(error);
//     });
// }

Office.onReady((info) => {
  console.log("Office.js is now ready in ${info.host} host.");
  $("#initialisation").on("click", () => tryCatch(initialisation));
  $("#AddIncident").on("click", () => tryCatch(addIncident));
  $("#AllRetake").on("click", () => tryCatch(allRetake));
  $("#Retake").on("click", () => tryCatch(retake));
  $("#Sollicitation").on("click", () => tryCatch(sollicitation()));
  $("#RetourGA").on("click", () => tryCatch(retourGA()));
  $("#DemandeValCom").on("click", () => tryCatch(demandeValCom()));
  $("#RetourValCom").on("click", () => tryCatch(retourValCom()));
  $("#FinPre").on("click", () => tryCatch(finPre()));
});

let incidentTimer = {};
let timerForSave = {};

// This function is for refresh the select when the number of incident is modify.
function refreshList(ids) {
  var select = document.getElementById("IdIncident");

  select.innerHTML = "";
  ids.forEach(function(option) {
    var el = document.createElement("option");
    el.textContent = option;
    el.value = option;
    select.appendChild(el);
  });
}

function addCellSave(cellpos, values, key) {
  return Excel.run(function(context) {
    var save = context.workbook.worksheets.getItem("Save");
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");
    save.load("values");
    return context.sync().then(function() {
      var lastRow = usedRangeSave.rowCount;
      var cell = save.getCell(cellpos[0] + 1, cellpos[1]);
      var cellNNI = save.getCell(cellpos[0] + 1, 2);
      cell.load("values");
      return context.sync().then(function() {
        cell.values = [[values + (cell.values[0][0] - timerForSave[key])]];
        cellNNI.values = [[""]];
        timerForSave[key] = values;
        return context.sync();
      });
    });
  });
}

/** the function for the initialisation of all timers */
function initialisation() {
  // We get the NNI of the users
  const nniInput = document.getElementById("NNI");
  const nniValue = nniInput.value;

  if (nniValue) {
    return Excel.run(function(context) {
      // We get the sheet Save for doing operation on it.
      var save = context.workbook.worksheets.getItem("Save");
      var usedRange = save.getUsedRange(true);
      usedRange.load("rowCount");

      return context.sync().then(function() {
        var lastRow = usedRange.rowCount;
        var range = save.getRange("A2:A" + lastRow);
        range.load("values"); // Charger les valeurs
        var rangeNNI = save.getRange("C2:C" + lastRow);
        rangeNNI.load("values");

        return context.sync().then(function() {
          var values = range.values;
          var valuesNNI = rangeNNI.values;
          for (var i = 0; i < values.length; i++) {
            if (!valuesNNI[i][0] && values[i][0]) {
              let key = values[i][0];
              timerForSave[key] = 0;
              incidentTimer[key] = new Date();
              save.getCell(i + 1, 2).values = [[nniValue]];
            }
          }
          refreshList(Object.keys(incidentTimer));
          console.log("Tout est bien initialisé.");
        });
      });
    });
  } else {
    console.log("Entrer un NNI");
  }
}

/**
 * SaveTimer need to by modify, is here for save the timer and for all other agents when he's take in charge the incident
 */
function allRetake() {
  return Excel.run(function(context) {
    var save = context.workbook.worksheets.getItem("Save");
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");
    save.load("values");
    return context.sync().then(function() {
      var lastRow = usedRangeSave.rowCount;
      var rangeSave = save.getRange("A2:A" + lastRow);
      // Ici vous devez charger les valeurs pour pouvoir les utiliser après context.sync()
      rangeSave.load("values");

      return context.sync().then(function() {
        var values = rangeSave.values;
        let promises = [];
        for (let key in incidentTimer) {
          for (let i = 0; i < values.length; i++) {
            if (values[i][0] === key) {
              // Ici vous pouvez accéder à la cellule et mettre à jour les valeurs
              var actualTime = new Date();

              promises.push(
                promises,
                addCellSave([i, 1], (actualTime.getTime() - incidentTimer[key].getTime()) / 1000 / 60, key)
              );
            }
          }
        }
        return Promise.all(promises).then(() => {
          for (let key in incidentTimer) delete incidentTimer[key];
          refreshList(Object.keys(incidentTimer));
          console.log("Liste d'incident vide");
          return context.sync();
        });
      });
    });
  });
}

/**
 * SaveTimer need to by modify, is here for save the timer and for all other agents when he's take in charge the incident
 */
function retake() {
  return Excel.run(function(context) {
    var save = context.workbook.worksheets.getItem("Save");
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");
    save.load("values");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    return context.sync().then(function() {
      var lastRow = usedRangeSave.rowCount;
      var rangeSave = save.getRange("A2:A" + lastRow);
      // Ici vous devez charger les valeurs pour pouvoir les utiliser après context.sync()
      rangeSave.load("values");

      return context.sync().then(function() {
        var values = rangeSave.values;
        let promises = [];
        for (let i = 0; i < values.length; i++) {
          if (values[i][0] === id) {
            // Ici vous pouvez accéder à la cellule et mettre à jour les valeurs
            var cell = save.getCell(i + 1, 1); // i + 1 car les index dans Excel commencent à 1, et non 0
            var actualTime = new Date();
            promises.push(
              promises,
              addCellSave([i, 1], (actualTime.getTime() - incidentTimer[id].getTime()) / 1000 / 60, id)
            );
          }
        }
        refreshList(Object.keys(incidentTimer));
        return Promise.all(promises).then(() => {
          delete incidentTimer[id];
          refreshList(Object.keys(incidentTimer));
          console.log("Incident retiré");
          return context.sync();
        });
      });
    });
  });
}

// For adding an incident
function addIncident() {
  const nniInput = document.getElementById("NNI");
  const nniValue = nniInput.value;

  const app = document.getElementById("Application");
  const appValue = app.value;

  const type_inc = document.getElementById("type_inc");
  const type_incValue = type_inc.value;

  if (nniValue && appValue) {
    return Excel.run(function(context) {
      var suivi = context.workbook.worksheets.getItem("Suivi");
      var save = context.workbook.worksheets.getItem("Save");

      // Load all row for check where we are going to put the value
      var usedRangeSuivi = suivi.getUsedRange();
      usedRangeSuivi.load("rowCount");

      // Load all row for check where we are going to put the value
      var usedRangeSave = save.getUsedRange();
      usedRangeSave.load("rowCount");

      return context.sync().then(function() {
        var lastRowSuivi = usedRangeSuivi.rowCount; // The last row used in 'Suivi'
        var lastRowSave = usedRangeSave.rowCount; // The last row used in 'Save'

        return context.sync().then(function() {
          var domaineCell = suivi.getCell(lastRowSuivi, 1);
          var idCellSuivi = suivi.getCell(lastRowSuivi, 0);
          var idCellSave = save.getCell(lastRowSave, 0);
          var cellSaveTimer = save.getCell(lastRowSave, 1);
          var idCellNNI = save.getCell(lastRowSave, 2);
          if (type_incValue == "MA") {
            var id = "MA-" + appValue[0] + appValue[1] + appValue[2] + lastRowSuivi;
          } else if (type_incValue == "MCP") {
            var id = "MCP-" + appValue[0] + appValue[1] + appValue[2] + lastRowSuivi;
          } else {
            var id = "TRANSV-" + appValue[0] + appValue[1] + appValue[2] + lastRowSuivi;
          }

          domaineCell.values = [[appValue]];
          idCellSuivi.values = [[id]];
          idCellSave.values = [[id]];
          idCellNNI.values = [[nniValue]];
          cellSaveTimer.values = [[0]];

          // Adding the new timer in the variable reserved to it
          incidentTimer[id] = new Date();
          timerForSave[id] = 0;
          refreshList(Object.keys(incidentTimer));
          return context.sync();
        });
      });
    });
  } else {
    console.log("Entrer le NNI et l'application concerné avant de vouloir ajouter un incident !");
  }
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

// Set the past time between "prise en charge" and "sollicitation" in the correct cell
function sollicitation() {
  return Excel.run(function(context) {
    var suivi = context.workbook.worksheets.getItem("Suivi");
    var save = context.workbook.worksheets.getItem("Save");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    // Load all row from the worksheet
    var usedRangeSuivi = suivi.getUsedRange();
    usedRangeSuivi.load("rowCount");

    // Load all row from the workseets save
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");

    return context.sync().then(function() {
      var searchCell = usedRangeSuivi.find(id, { matchCase: true });
      searchCell.load("rowIndex");
      var saveCell = usedRangeSave.find(id, { matchCase: true });
      saveCell.load("rowIndex");
      return context.sync().then(function() {
        var sollicitationCell = suivi.getCell(searchCell.rowIndex, 2);
        sollicitationCell.load("values");

        var saveTimer = save.getCell(saveCell.rowIndex, 1);
        saveTimer.load("values");

        return context.sync().then(function() {
          if (sollicitationCell.values[0][0] === null || sollicitationCell.values[0][0] === "") {
            var actualTime = new Date();
            sollicitationCell.values = [
              [(actualTime.getTime() - incidentTimer[id]) / 1000 / 60 + saveTimer.values[0][0]]
            ];
          } else {
            console.log("Vous avez déjà renseigner cette catégorie pour l'incident que vous avez selectionné");
          }

          return context.sync();
        });
      });
    });
  });
}

// Set the past time between "sollicitation" and "retour du GA" in the correct cell
function retourGA() {
  return Excel.run(function(context) {
    var suivi = context.workbook.worksheets.getItem("Suivi");
    var save = context.workbook.worksheets.getItem("Save");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    // Load all row from the worksheet
    var usedRangeSuivi = suivi.getUsedRange();
    usedRangeSuivi.load("rowCount");

    // Load all row from the workseets save
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");

    return context.sync().then(function() {
      var searchCell = usedRangeSuivi.find(id, { matchCase: true });
      searchCell.load("rowIndex");

      var saveCell = usedRangeSave.find(id, { matchCase: true });
      saveCell.load("rowIndex");

      return context.sync().then(function() {
        var retourGACell = suivi.getCell(searchCell.rowIndex, 3);
        var sollicitationCell = suivi.getCell(searchCell.rowIndex, 2);
        sollicitationCell.load("values");
        retourGACell.load("values");

        var saveTimer = save.getCell(saveCell.rowIndex, 1);
        saveTimer.load("values");

        return context.sync().then(function() {
          if (retourGACell.values[0][0] === null || retourGACell.values[0][0] === "") {
            var actualTime = new Date();
            retourGACell.values = [
              [
                (actualTime.getTime() - incidentTimer[id]) / 1000 / 60 +
                  saveTimer.values[0][0] -
                  sollicitationCell.values[0][0]
              ]
            ];
          } else {
            console.log("Vous avez déjà renseigner cette catégorie pour l'incident que vous avez selectionné");
          }

          return context.sync();
        });
      });
    });
  });
}

// Set the past time between "retour du GA" and "demande de validation de comm." in the correct cell
function demandeValCom() {
  return Excel.run(function(context) {
    var suivi = context.workbook.worksheets.getItem("Suivi");
    var save = context.workbook.worksheets.getItem("Save");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    // Load all row from the worksheet
    var usedRangeSuivi = suivi.getUsedRange();
    usedRangeSuivi.load("rowCount");

    // Load all row from the workseets save
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");

    return context.sync().then(function() {
      var searchCell = usedRangeSuivi.find(id, { matchCase: true });
      searchCell.load("rowIndex");

      var saveCell = usedRangeSave.find(id, { matchCase: true });
      saveCell.load("rowIndex");

      return context.sync().then(function() {
        var demandeValComCell = suivi.getCell(searchCell.rowIndex, 4);
        var retourGACell = suivi.getCell(searchCell.rowIndex, 3);
        var sollicitationCell = suivi.getCell(searchCell.rowIndex, 2);
        demandeValComCell.load("values");
        sollicitationCell.load("values");
        retourGACell.load("values");

        var saveTimer = save.getCell(saveCell.rowIndex, 1);
        saveTimer.load("values");

        return context.sync().then(function() {
          if (demandeValComCell.values[0][0] === null || demandeValComCell.values[0][0] === "") {
            var actualTime = new Date();
            demandeValComCell.values = [
              [
                (actualTime.getTime() - incidentTimer[id]) / 1000 / 60 +
                  saveTimer.values[0][0] -
                  (sollicitationCell.values[0][0] + retourGACell.values[0][0])
              ]
            ];
          } else {
            console.log("Vous avez déjà renseigner cette catégorie pour l'incident que vous avez selectionné");
          }

          return context.sync();
        });
      });
    });
  });
}

// Set the past time between the "demande de validation" and "retour sur la validation"
function retourValCom() {
  return Excel.run(function(context) {
    var suivi = context.workbook.worksheets.getItem("Suivi");
    var save = context.workbook.worksheets.getItem("Save");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    // Load all row from the worksheet
    var usedRangeSuivi = suivi.getUsedRange();
    usedRangeSuivi.load("rowCount");

    // Load all row from the workseets save
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");

    return context.sync().then(function() {
      var searchCell = usedRangeSuivi.find(id, { matchCase: true });
      searchCell.load("rowIndex");

      var saveCell = usedRangeSave.find(id, { matchCase: true });
      saveCell.load("rowIndex");

      return context.sync().then(function() {
        var dureeTtlCell = suivi.getCell(searchCell.rowIndex, 6);
        var validationCell = suivi.getCell(searchCell.rowIndex, 5);
        var demandeValComCell = suivi.getCell(searchCell.rowIndex, 4);
        var retourGACell = suivi.getCell(searchCell.rowIndex, 3);
        var sollicitationCell = suivi.getCell(searchCell.rowIndex, 2);
        demandeValComCell.load("values");
        retourGACell.load("values");
        sollicitationCell.load("values");
        validationCell.load("values");

        var saveTimer = save.getCell(saveCell.rowIndex, 1);
        saveTimer.load("values");

        return context.sync().then(function() {
          if (validationCell.values[0][0] === null || validationCell.values[0][0] === "") {
            var actualTime = new Date();
            console.log(saveTimer.values[0][0]);
            validationCell.values = [
              [
                (actualTime.getTime() - incidentTimer[id]) / 1000 / 60 +
                  saveTimer.values[0][0] -
                  (sollicitationCell.values[0][0] + retourGACell.values[0][0] + demandeValComCell.values[0][0])
              ]
            ];
            dureeTtlCell.values = [
              [
                (actualTime.getTime() - incidentTimer[id]) / 1000 / 60 +
                  saveTimer.values[0][0] +
                  (sollicitationCell.values[0][0] + retourGACell.values[0][0] + demandeValComCell.values[0][0])
              ]
            ];
            save.getRange("A" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            save.getRange("B" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            save.getRange("C" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            delete incidentTimer[id];
            refreshList(Object.keys(incidentTimer));
          } else {
            console.log("Vous avez déjà renseigner cette catégorie pour l'incident que vous avez selectionné");
          }

          return context.sync();
        });
      });
    });
  });
}

// This for the case where the incident doesn't need more investigation on it.
function finPre() {
  return Excel.run(function(context) {
    var suivi = context.workbook.worksheets.getItem("Suivi");
    var save = context.workbook.worksheets.getItem("Save");

    const select = document.getElementById("IdIncident");
    const id = select.value;

    // Load all row from the worksheet
    var usedRangeSuivi = suivi.getUsedRange();
    usedRangeSuivi.load("rowCount");

    // Load all row from the workseets save
    var usedRangeSave = save.getUsedRange();
    usedRangeSave.load("rowCount");

    return context.sync().then(function() {
      var searchCell = usedRangeSuivi.find(id, { matchCase: true });
      searchCell.load("rowIndex");

      var saveCell = usedRangeSave.find(id, { matchCase: true });
      saveCell.load("rowIndex");

      return context.sync().then(function() {
        var dureeTtlCell = suivi.getCell(searchCell.rowIndex, 6);
        var validationCell = suivi.getCell(searchCell.rowIndex, 5);
        var demandeValComCell = suivi.getCell(searchCell.rowIndex, 4);
        var retourGACell = suivi.getCell(searchCell.rowIndex, 3);
        var sollicitationCell = suivi.getCell(searchCell.rowIndex, 2);
        demandeValComCell.load("values");
        retourGACell.load("values");
        sollicitationCell.load("values");
        validationCell.load("values");

        var saveTimer = save.getCell(saveCell.rowIndex, 1);
        saveTimer.load("values");

        return context.sync().then(function() {
          if (demandeValComCell.values[0][0] === null || demandeValComCell.values[0][0] === "") {
            var actualTime = new Date();

            demandeValComCell.values[0][0] = 0;
            validationCell.values[0][0] = 0;

            dureeTtlCell.values = [
              [
                (sollicitationCell.values[0][0] +
                  retourGACell.values[0][0] +
                  demandeValComCell.values[0][0] +
                  actualTime.getTime() -
                  incidentTimer[id]) /
                  1000 /
                  60
              ]
            ];

            save.getRange("A" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            save.getRange("B" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            save.getRange("C" + (saveCell.rowIndex + 1).toString()).delete(Excel.DeleteShiftDirection.up);
            delete incidentTimer[id];
            refreshList(Object.keys(incidentTimer));
          } else {
            console.log(
              "Cette option n'est plus disponible pour cette incident, ou vous n'avez pas rempli les condition pour confirmer cette action."
            );
          }

          return context.sync();
        });
      });
    });
  });
}
