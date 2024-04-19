Office.onReady((info) => {
  console.log("Office.js is now ready in ${info.host} host.");
  $("#initialisation").on("click", () => tryCatch(initialisation));
});

let incidentTimer = {};

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

// For adding an incident
function initialisation() {
  const nniInput = document.getElementById("NNI");
  const nniValue = nniInput.value;

  if (nniValue) {
    return Excel.run(function(context) {
      var nniSheet = context.workbook.worksheets.getItem(nniValue);

      // Load all row for check where we are going to put the value
      var usedRangeNniSheet = nniSheet.getUsedRange();
      usedRangeNniSheet.load("rowCount");

      return context.sync().then(function() {
        var lastRowNniSuivi = usedRangeNniSheet.rowCount; // The last row used in 'Suivi'

        return context.sync().then(function() {
          var nniCellNni = nniSheet.getCell(lastRowNniSuivi, 1);
          var idDateCell = nniSheet.getCell(lastRowNniSuivi, 0);
          var actualDate = new Date()
                                          
          nniCellNni.values = [[nniValue]];
          idDateCell.values = [[actualDate.toLocaleDateString()]];

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
