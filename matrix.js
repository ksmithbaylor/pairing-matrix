///////////////////////////////////////////////////////////////////////////////
// Public functions (bound to buttons)
///////////////////////////////////////////////////////////////////////////////

// Increments the pairing total and turns the names orange. Called
// from the orange "PAIR" button
function pair() {
  pairOrUnpair(true, false);
};

// Decrements the pairing total and turns the names back to green.
// Called from the green "UNPAIR" button.
function unpair() {
  pairOrUnpair(false, false);
};

// Same as `pair`, but turns the names blue. Called from the blue
// "TOW TRUCK" button.
function towTruckPair() {
  pairOrUnpair(true, true);
}

// Clears all notes and changes the name cells back to green, but
// leaves the historical pairing totals alone. Called by the gray
// "Restart Pairing" button.
function restartPairing() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getMaxRows();
  var cols = sheet.getMaxColumns();
  var colors = getColors();

  for (var row = 1; row <= rows - 2; row++) {
    sheet.getRange(row, row + 1, 1, 1).setBackground(colors.green);
  }

  sheet.getRange(2, 1, rows - 3, 1).setBackground(colors.green);

  var allCells = sheet.getRange(1, 1, rows, cols);
  setBorder(allCells, false);
  allCells.clearNote();
}

// Resets all historical pairing totals to zero and then calls
// `restartPairing`. Called by the gray "Reset All" button.
function resetAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getMaxRows();
  var cols = sheet.getMaxColumns();
  var values = sheet.getSheetValues(1, 1, rows, cols);

  for (var row = 2; row <= rows; row++) {
    for (var col = 1; col <= cols - 1; col++) {
      if (isNonzeroNumber(values[row - 1][col - 1])) {
        sheet.getRange(row, col, 1, 1).setValue('0');
      }
    }
  }

  restartPairing();
};


///////////////////////////////////////////////////////////////////////////////
// Private functions (used internally by the public functions)
///////////////////////////////////////////////////////////////////////////////

// Does all the steps to pair or unpair two people, based on the values of
// `shouldBePaired` and `shouldBeTowTruck`
function pairOrUnpair(shouldBePaired, shouldBeTowTruck) {
  var context = getContext();

  var newValue = shouldBePaired ? context.value + 1 : context.value - 1;
  var newValueSafe = Math.max(0, newValue);
  context.cell.setValue(newValueSafe);

  setBorder(context.cell, shouldBePaired);
  setNote(context.cell, shouldBePaired);

  var colors = getColors();
  if (shouldBePaired) {
    var color = shouldBeTowTruck ? colors.blue : colors.orange;
  } else {
    var color = colors.green;
  }
  setNameColors(context, color)
};

// Returns useful information about the current cell, to prevent
// duplicating work everywhere this is needed.
function getContext() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell  = sheet.getActiveCell();
  var row   = cell.getRow();
  var col   = cell.getColumn();

  return {
    sheet: sheet,
    cell: cell,
    row: row,
    col: col,
    inBottomTwo: (row + 2 > sheet.getMaxRows()),
    value: cell.getValue(),
    topNameCells: [
      (col < 3) ? undefined : sheet.getRange(col - 1, 1, 1, 1),
      sheet.getRange(col - 1, col),
    ],
    bottomNameCells: [
      sheet.getRange(row, 1, 1, 1),
      sheet.getRange(row, Math.min(row + 1, sheet.getMaxColumns()), 1, 1)
    ]
  };
}

// Returns color constants.
function getColors() {
  return {
    orange: '#e69138',
    green:  '#38761d',
    blue:   '#3c78d8'
  };
}

// Sets the border of `range` if `x` is true, or removes it if `x` is false.
function setBorder(range, x) {
  range.setBorder(x, x, x, x, x, x);
}

// Sets the note of all cells in `range` to the current date if `x` is
// true, or clears all notes if `x` is false.
function setNote(range, x) {
  if (x) {
    range.setNote(new Date().toDateString());
  } else {
    range.clearNote();
  }
}

// Based on the values in `context` (an object returned by `getContext`),
// sets the background color of the relevant name cells to `color`.
function setNameColors(context, color) {
  context.topNameCells.forEach(function(cell) {
    if (cell) cell.setBackground(color);
  });

  if (!context.inBottomTwo) {
    context.bottomNameCells.forEach(function(cell) {
      cell.setBackground(color);
    });
  }
}

// Effeciently returns true if and only if the value passed in is a non-zero
// number
function isNonzeroNumber(value) {
  if (value === undefined) return false;
  if (value === 0) return false;
  if (typeof value === 'number') return true;
  return false;
}
