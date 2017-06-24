/*
  The function must be usen only on solution sheets. Count how
  how many sheets are  charts or dedicated to analysis and set
  the starting point to the first actual solution sheet.
*/

/*
  Main method of this script which employes other methods in this script
  to retrieve metrics from each solution sheet on the spreadsheet and
  writes the information of each sheet on a row of an analysis sheet.
*/
function spreadSheetAnalysis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var actual = SpreadsheetApp.getActiveSheet();
  var limite = sheets.length;

  for(planilha =19; planilha < 35; planilha++){
    var sheet = sheets[planilha];
    var formula = sheet.getDataRange().getFormulas().toString();
    var cellsWithFormula = countCellsWithFormulas(sheet);
    var occupiedCells = countOccupiedCells(sheet);
    actual.appendRow([sheet.getName(), getNumberFunctionCalls(sheet), occupiedCells, uniqueFunctions(formula), occupiedCells - cellsWithFormula]);
  }
}

/*
  Retrieves the number of cells occupied 
  by functions in a sheet
*/
function countCellsWithFormulas(sheet){
    var cellsWithFormula = 0;
    var countTotal = 0;
    for(i = 0; i < sheet.getDataRange().getFormulas().length; i++){
      var blankLine = true;
      for(j = 0; j < sheet.getDataRange().getFormulas()[i].length; j++){
        var formulaN = sheet.getDataRange().getFormulas()[i][j];
        var values = sheet.getDataRange().getValues()[i][j];
        if (uniqueFunctions(formulaN) > 0 && values.toString().trim() != "") {
          cellsWithFormula +=  1;
        }
        if (formulaN.toString().trim() != "") {
          countTotal += 1;
          blankLine = false;
        }
      }
      if (blankLine) {
        //break;
      }
    }
    
    return cellsWithFormula;

}

/*
  Retrieves the number of occupied cells  in a sheet
*/
function countOccupiedCells(sheet) {
  var countTotal = 0;
  var values = sheet.getDataRange().getValues();
  var formulas = sheet.getDataRange().getFormulas();
  for(i = 0; i < values.length; i++){
    var blankLine = true;
    for(j = 0; j < values[i].length; j++){
      if (values[i][j].toString().trim() != "") {
        blankLine = false;
        countTotal += 1;
      }
    }
    if (blankLine) {
      //break;
    }
  }
  return countTotal;
}

/*
  Retrieves the number of times functions 
  were invoked in a sheet
*/
function getNumberFunctionCalls(sheet) {
  var countTotal = 0;
  var values = sheet.getDataRange().getValues();
  var formulas = sheet.getDataRange().getFormulas();
  for(i = 0; i < values.length; i++){
    var blankLine = true;
    for(j = 0; j < values[i].length; j++){
      if (values[i][j].toString().trim() != "") {
        blankLine = false;
        countTotal += numberFunctions(formulas[i][j]);
      }
    }
  }
  return countTotal;
}


/*
  Retrieves the number of function calls
  in a formula.
*/
function numberFunctions(formula){
  var match = formula.match(/[\w\.]+\(/g);
  if (match) {
    return match.length;
  }
  return 0;
}

/*
  Retrieves the number of different types
  of functions employed in a sheet.
*/
function uniqueFunctions(formula){
  var match = formula.match(/[\w\.]+\(/g);
  if(match){
    var list = match.toString().split(",");
    list = list.filter( function( item, index, inputArray ) {
      return inputArray.indexOf(item) == index;
    });
    return list.length;
  }
  return 0;
}
