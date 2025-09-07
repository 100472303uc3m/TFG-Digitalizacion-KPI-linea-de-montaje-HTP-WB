// === Helper Functions ====================================================================================
// Function that normalizes strings
function normalize(str) {
  return String(str).toLowerCase().replace(/\s+/g, '').trim();
}

// Function that checks if current row is full of blank cells
function isWhiteRow(row) {
  return row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
}

// Function to process integers read from spreadsheet
function processData(values) {
  // Iterate through array of arrays
  for (const row in values) {
    for (const col in values[row]) {
      if (values[row][col] != '') { // Check if cell was empty
        values[row][col] = Math.round(values[row][col] * 1000) / 1000;
      }
      else {  // Fill empty cell values with 0´s
        values[row][col] = 0;
      }
    }
  }

  return values;
};

// Function to safely parse and avoid divide-by-zero
function safeDivide(numerator, denominator) {
  // Define variables
  var num = parseFloat(numerator) || 0;
  var den = parseFloat(denominator);

  // If denominator doesn´t exist or is 0 we return 0
  if (!den || den === 0) {
    return 0;
  }

  // Return division (formatted)
  return +(num / den).toFixed(1);
}

// === Main Processing Functions ==============================================================================
// Function that processes tabular data from spreadsheets given specific information about the data
function ProcessSpreadsheetData(title, rowTitles, data, sheet) {
  // Iterate through data
  for (let row = 0; row < data.length; row++) {
    for (let col = 0; col < data[row].length; col++) {
      // Define current cell
      var cell = data[row][col];

      // Check if cell value matches the title
      if (normalize(String(cell)) == normalize(title)) {
        // LOG
        console.log('Found title ' + title + ' at row ' + row + ', col ' + col);

        // Define start of data in table using the date bar as reference
        var dataStart = extractDateCol(data, row, col);

        // Return empty array if date bars weren´t found
        if (dataStart == -1) {
          // LOG
          console.log('Unable to find dates in sheet');

          // Return empty array
          return [];
        }

        // Define end of data in table using the date bar as reference and adding more columns if the title matches specific cases (actual & target comparison tables)
        if (normalize(title).indexOf(normalize('YtD (%)')) != -1 && sheet.indexOf('C') != -1) {
          var dataEnd = dataStart + 37;
        }
        else {
          var dataEnd = dataStart + 11;
        }
        
        // We try a vertical search
        var table = extractTableData(data, row + 1, col, dataStart, dataEnd, rowTitles);

        // We try a horizontal search if no table was found
        if (table.length == 0) {
          var table = extractTableData(data, row, col + 1, dataStart, dataEnd, rowTitles);
        }

        // Return an empty array if no table was found
        if (table.length == 0) {
          // LOG
          console.log('Unable to extract table at ' + row + ', ' + col);

          // Return empty array
          return [];
        }

        // LOG
        console.log('Extracted table:\n' + table);

        // Return table data       
        return table;
      }
    }
  }

  // LOG
  console.log('Title ' + title + ' not found!');

  // Return empty array
  return [];
};

// Function that extracts column where dates start
function extractDateCol(data, startRow, startCol) {
  // Search for dates to the right of the title position
  for (let nextCol = startCol + 1; nextCol < data[0].length; nextCol++) {
    if (data[startRow][nextCol].indexOf('ene') != -1) {
      return (nextCol);
    }
  }

  // Search for dates from the beginning of the sheet by iterating through data again
  for (let row = 0; row < data.length; row++) {
    for (let col = 0; col < data[row].length; col++) {
      // Return column when we find start of dates
      if (data[row][col].indexOf('ene') != -1) {
        return (col);
      }
    }
  }

  // Return error if we found no dates
  return -1;
};


// Function that extracts tabular data in proximity to a found title
function extractTableData(data, startRow, startCol, dataStart, dataEnd, rowTitles) {
  // Define auxiliary variables to stop extracting table data
  const MAX_EMPTY_ROWS = 2;
  let emptyRowCount = 0;
  
  // Define array for the table to return
  var table = [];

  // Define number of rows to extract
  var numRows = data.length - startRow;
  
  // Define number of columns to extract
  var numCols = dataEnd - startCol;

  // Iterate through number of rows
  for (let i = 0; i < numRows; i++) {
    // Define array for header parts and row to append, current row index, and data in row
    var row = [];
    var headerParts = [];
    var rowIndex = startRow + i;
    var rowData = data[rowIndex];
    
    // Check if current row is empty
    if (isWhiteRow(rowData)) {
      emptyRowCount++;
      if (emptyRowCount >= MAX_EMPTY_ROWS) {
        console.log(`Stopped parsing after ${MAX_EMPTY_ROWS} empty rows.`);
        break;
      }
      
      // Skip row
      continue;
    }
    // Reset counter when non-empty row is found
    else {
      emptyRowCount = 0; 
    }

    // If row titles are included, check if current row is to be added
    var rowContainsTitle = !rowTitles || rowTitles.length === 0 || rowTitles.some(title =>
      rowData.some(cell => normalize(cell) == normalize(title))
    );

    // Skip if it isn´t present
    if (!rowContainsTitle) {
      continue;
    }

    // Iterate through number of columns
    for (let j = 0; j <= numCols; j++) {
      // Define column index and data in current cell
      var colIndex = startCol + j;
      var cell = rowData[colIndex];

      // If the column index is before the start of the data, we push into header
      if (colIndex < dataStart) {
        if (cell !== '' && cell !== null && cell !== undefined) {
          headerParts.push(String(cell).trim());
        }
      } 
      // Append current value into row as normal otherwise
      else {
        row.push(cell);
      }
    }

    // Combine headers into one string and insert at the beginning
    row.unshift(headerParts.join(' - '));

    // Append row into table array
    table.push(row);
  }

  // Return found table
  return table;
};

// === Auxiliary Processing Functions ==============================================================================================
// Function that returns the month index based on the name of the spreadsheet
function ReturnMonthIndex(activeSpreadsheet) {
  // Get active spreadsheet
  const ss = SpreadsheetApp.openById(activeSpreadsheet);

  // Get spreadsheet name
  var ssName = ss.getName();

  // Auxiliary list
  var monthList = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

  // Retrieve month from name
  var monthIndex;
  monthList.forEach((month) => {
    if (ssName.includes(month)) {
      monthIndex = monthList.indexOf(month);
      console.log('Month Index: ' + monthIndex);
    }
  });

  // If not found, return actual date
  if (monthIndex == undefined) {
    var date = new Date;
    monthIndex = date.getMonth();
    console.log('Month Index: ' + monthIndex);
  }

  // Return month index
  return monthIndex;
};

// Function that retrieves data ranges from sheets
function ReturnDataRangeValues(spreadsheetId, sheetName) {
  // Check if the spreadsheet ID has been defined
  if (spreadsheetId != undefined) {
    try {
      // Try to extract data range values
      var ss = SpreadsheetApp.openById(spreadsheetId);
      var values = ss.getSheetByName(sheetName).getDataRange().getDisplayValues();
      
      // Return values
      return values;
    }
    catch(error) {
      // LOG
      console.log(error);

      // Return empty array
      return [];
    }
  }
  else {
    // Return empty array
    return [];
  }
};

function CalculateS19UnitHours(spreadsheet) {
  // Get sheet values
  const incurredValues = ReturnDataRangeValues(spreadsheet, 'W-103');
  const productionValues = ReturnDataRangeValues(spreadsheet, 'Presum');

  // Check values
  if (incurredValues.length == 0 || productionValues == 0) {
    // LOG
    console.log('Error accessing spreadsheet with ID ' + spreadsheet);

    // Return empty array
    return [];
  }

  // Get incurred hours
  const incurredHoursTableTitle = tableDict['W-103'].title;
  const incurredHoursRowTitles = tableDict['W-103'].rowTitles;
  var incurredHours = ProcessSpreadsheetData(incurredHoursTableTitle, incurredHoursRowTitles, incurredValues, 'W-103');

  // Get production plan
  const productionPlanTableTitle = tableDict['Presum'].title;
  const productionPlanRowTitles = tableDict['Presum'].rowTitles;
  var productionPlan = ProcessSpreadsheetData(productionPlanTableTitle, productionPlanRowTitles, productionValues, 'Presum');

  // Check values
  if (incurredHours.length == 0 || productionPlan.length == 0) {
    // We just return an empty array, no need to log
    return [];
  }

  // Define keywords to clean the incurred hours data table
  const keywords = ['HORAS A', 'HORAS B', 'HORAS C'];

  // Filter rows
  incurredHours = incurredHours.filter(row =>
    keywords.some(keyword =>
      row.some(cell =>
        String(cell).toLowerCase().includes(keyword.toLowerCase())
      )
    )
  );

  // Define an array to calculate and store unit hours
  var actualValues = [[], [], []];
  
  // Push Pintura Total
  for (let i = 1; i < incurredHours[0].length; i++) {
    var totalIncurred = (parseFloat(incurredHours[0][i]) || 0) + (parseFloat(incurredHours[1][i]) || 0) + (parseFloat(incurredHours[2][i]) || 0);
    const unitHour = safeDivide(totalIncurred, productionPlan[0][i]);
    actualValues[0].push(unitHour);
  }

  // Push Cuarto Sellantes Total
  for (let i = 1; i < incurredHours[3].length; i++) {
    var totalIncurred = (parseFloat(incurredHours[3][i]) || 0) +(parseFloat(incurredHours[4][i]) || 0) + (parseFloat(incurredHours[5][i]) || 0);
    const unitHour = safeDivide(totalIncurred, productionPlan[0][i]);
    actualValues[1].push(unitHour);
  }

  // Push UH Totals (sum of both programs)
  for (let i = 0; i < actualValues[0].length; i++) {
    var sum = (parseFloat(actualValues[0][i]) || 0) + (parseFloat(actualValues[1][i]) || 0);
    actualValues[2].push(+sum.toFixed(1));
  }
  // LOG
  console.log('Calculated values:\n' + actualValues);

  // Return calculated values
  return actualValues;
}