// === Table Definitions ==========================================================================================
const tableDict = {
                    'Actual': 
                              {
                                '100': 
                                      [
                                        {title: '100 Actual', rowTitles: []},
                                        {title: '100 Report', rowTitles: []}
                                      ],
                                '101': 
                                      [
                                        {title: '101 Actual', rowTitles: []},
                                        {title: '101 Report', rowTitles: []}
                                      ],
                                '102': 
                                      [
                                        {title: '102 Actual', rowTitles: []},
                                        {title: '102 Report', rowTitles: []}
                                      ],
                                '103': 
                                      [
                                        {title: '103 Actual', rowTitles: []},
                                        {title: '103 Report', rowTitles: []}
                                      ]
                              },
                    'Presum':
                              {
                                title: 'A350 S19', rowTitles: ['Plan de prod']
                              },
                    'W-100':  
                            {
                              title: 'Unit hours YtD', rowTitles: ['HORAS A', 'HORAS B', 'HORAS C', 'UNIT HOURS YtD']
                            },
                    'W-101': 
                            {
                              title: 'Unit hours YtD', rowTitles: ['HORAS A', 'HORAS B', 'HORAS C', 'UNIT HOURS YtD']
                            },
                    'W-102': 
                            {
                              title: 'Unit hours YtD', rowTitles: ['HORAS A', 'HORAS B', 'HORAS C', 'UNIT HOURS YtD']
                            },
                    'W-103': 
                            {
                              title: 'Incurrido mensual', rowTitles: ['A350 S19'],
                            },
                    'W-A320':
                            {
                              title: 'Unit Hours YtD', rowTitles: ['Total', 'UNIT HOURS YtD']
                            },
                    'W-A330':
                            {
                              title: 'Unit Hours YtD', rowTitles: ['Total', 'UNIT HOURS YtD']
                            },
                    'W-A350':
                            {
                              title: 'Unit Hours YtD', rowTitles: ['Total', 'UNIT HOURS YtD']
                            },
                    'C-100': 
                            [
                              {title: 'A320 HTP', rowTitles: []},
                              {title: '100 Competitividad YtD (FTE)', rowTitles: []},
                              {title: '100 YtD (%)', rowTitles: []}
                            ],
                    'C-101': 
                            [
                              {title: 'A330 HTP', rowTitles: []},
                              {title: '101 Competitividad YtD (FTE)', rowTitles: []},
                              {title: '101 YtD (%)', rowTitles: []}
                            ],
                    'C-102': 
                            [
                              {title: 'A350 HTP', rowTitles: []},
                              {title: '102 Competitividad YtD (FTE)', rowTitles: []},
                              {title: '102 YtD (%)', rowTitles: []}
                            ],
                    'C-103': 
                            [
                              {title: 'Secciones Auxiliares', rowTitles: []},
                              {title: 'Secciones Aux (YtD)', rowTitles: []}
                            ],
                  };

// === Spreadsheet ID Extraction ==================================================================================
// Global variables
var forecastQuarterOneId;
var forecastQuarterTwoId;
var forecastQuarterThreeId;
var forecastQuarterFourId;
var activeSpreadsheet;

// Retrieve links from spreadsheet
try {
  const ss = SpreadsheetApp.openById('1rHltJDUtCu6zGz-jziYVCfU5ZxQofNAHeOdsP9KYLGo');
  const links = ss.getRange('Enlaces!D6:D9').getValues();
  Logger.log(links);

  // Initialize variables for link extraction
  var link1 = links[0][0];
  var link2 = links[1][0];
  var link3 = links[2][0];
  var link4 = links[3][0];
}
catch(error) {
  // LOG
  console.log(error);
}

// Try extracting the ID from the link with a regular expression
// ID 1
try {
  var capturedId1 = link1.match(/\/d\/(.+)\//);
  forecastQuarterOneId = capturedId1[1];  
}
catch(error) {
  // Pass
}
// ID 2
try {
  var capturedId2 = link2.match(/\/d\/(.+)\//);
  forecastQuarterTwoId = capturedId2[1];
}
catch(error) {
  // Pass
}
// ID 3
try {
  var capturedId3 = link3.match(/\/d\/(.+)\//);
  forecastQuarterThreeId = capturedId3[1];
}
catch(error) {
  // Pass
}
// ID 4
try {
  var capturedId4 = link4.match(/\/d\/(.+)\//);  
  forecastQuarterFourId = capturedId4[1];
}
catch(error) {
  // Pass
}

// === HTML Related Functions =====================================================================================
// Function to process HTTP GET requests
function doGet(e) {
  activeSpreadsheet = e.parameter.sheetId;
  let page = e.parameter.mode || "Workload";
  let html = HtmlService.createTemplateFromFile(page).evaluate();
  let htmlOutput = HtmlService.createHtmlOutput(html);
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');

  // Replaces {{NAVBAR}} in pages with the navigation bar content
  htmlOutput.setContent(htmlOutput.getContent().replace("{{NAVBAR}}",getNavbar(page)));
  return htmlOutput;
};

// Function that creates the navigation bar and returns it to the doGet function to append it into the pages
function getNavbar(activePage) {
  var scriptURLPage1 = getScriptURL('program=first&sheetId=' + activeSpreadsheet);
  var scriptURLPage2 = getScriptURL('mode=Competitividad&program=first&sheetId=' + activeSpreadsheet);

  var navbar = 
    `<nav class="navbar navbar-expand-lg navbar-dark sticky-top" style="background-color: rgb(0, 32, 91);">
        <div class="container">
        <a class="navbar-brand" style="color: white;">Data Visualization</a>
        <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNavAltMarkup" aria-controls="navbarNavAltMarkup" aria-expanded="false" aria-label="Toggle navigation">
          <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navbarNavAltMarkup">
          <div class="navbar-nav">
            <a class="nav-item nav-link ${activePage === 'Workload' ? 'active' : ''} mx-2" href="${scriptURLPage1}">Workload</a>
            <a class="nav-item nav-link ${activePage === 'Competitividad' ? 'active' : ''} mx-2" href="${scriptURLPage2}">Competitividad</a>
          </div>
        </div>
        </div>
      </nav>`;
  return navbar;
};

// Returns the URL of the Google Apps Script web app
function getScriptURL(qs = null) {
  var url = ScriptApp.getService().getUrl();
  if(qs){
    if (qs.indexOf("?") === -1) {
      qs = "?" + qs;
    }
    url = url + qs;
  }
  return url;
};

// Function to include HTML files (JS and CSS as well)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
};

// === Spreadsheet Reading Functions ==============================================================================
// Function that returns program unit hours -> Array of Actual + Forecast and Target UH values
function getUnitHours(activeSpreadsheet, program) {
  // Define sheet name
  const sheet = 'W-' + program;

  // Define data to return
  var finalData = {};

  // Special case: Program S19
  if (program == 'S19') {
    // Call function that handles special case - Actual Values
    var actualValues = CalculateS19UnitHours(activeSpreadsheet);

    // Call function that handles special case - Target Values
    // Inactive spreadsheet values - Target
    var targetValues = [];
    if (forecastQuarterOneId) {
      var targetOneValues = CalculateS19UnitHours(forecastQuarterOneId);
      targetValues.push(targetOneValues);
    }
    if (forecastQuarterTwoId) {
      var targetTwoValues = CalculateS19UnitHours(forecastQuarterTwoId);
      targetValues.push(targetTwoValues);
    }
    if (forecastQuarterThreeId) {
      var targetThreeValues = CalculateS19UnitHours(forecastQuarterThreeId);
      targetValues.push(targetThreeValues);
    }
    if (forecastQuarterFourId) {
      var targetFourValues = CalculateS19UnitHours(forecastQuarterFourId);
      targetValues.push(targetFourValues);
    }

    // Check values
    targetValues.forEach((dataArray, index) => {
      console.log(dataArray)
      if (dataArray.length != 0) {
        finalData[('TargetValues' + parseInt(index+1))] = dataArray;
      }
      else {
        finalData[('TargetValues' + parseInt(index+1))] = [];
      }
    });
  }
  // Default case:
  else {
    // Active spreadsheet values - Actual + Forecast
    const activeSpreadsheetValues = ReturnDataRangeValues(activeSpreadsheet, sheet);

    // Check values
    if (activeSpreadsheetValues.length == 0) {
      // LOG
      console.log('Error accessing spreadsheet with ID ' + activeSpreadsheet);

      // Return empty object
      return finalData;
    }

    // Define data table parameters
    const title = tableDict[sheet].title;
    const rowTitles = tableDict[sheet].rowTitles;
    
    // Process actual values with search script
    var actualValues = ProcessSpreadsheetData(title, rowTitles, activeSpreadsheetValues, sheet);

    // Inactive spreadsheet values - Target
    var targetValues = [];
    var targetOneValues = ReturnDataRangeValues(forecastQuarterOneId, sheet);
    var targetTwoValues = ReturnDataRangeValues(forecastQuarterTwoId, sheet);
    var targetThreeValues = ReturnDataRangeValues(forecastQuarterThreeId, sheet);
    var targetFourValues = ReturnDataRangeValues(forecastQuarterFourId, sheet);
    targetValues.push(targetOneValues, targetTwoValues, targetThreeValues, targetFourValues);

    // Check values
    targetValues.forEach((dataArray, index) => {
      if (dataArray.length != 0) {
        var processedArray = ProcessSpreadsheetData(title, rowTitles, dataArray, sheet);
        finalData[('TargetValues' + parseInt(index+1))] = processedArray;
      }
      else {
        finalData[('TargetValues' + parseInt(index+1))] = [];
      }
    });
  }

  // Add actual values to final data to return
  finalData.ActualValues = actualValues;

  // Define month index
  var monthIndex = ReturnMonthIndex(activeSpreadsheet);
  finalData.MonthIndex = monthIndex;

  // LOG
  console.log('(' + sheet + ') Values:');
  console.log('=============================================================================================')
  Object.entries(finalData).forEach(([key, value]) => {
    console.log(key + ': ' + value);
  });

  // Return final data
  return finalData;
};

// Function that returns competitiveness values -> Array of Actual + Forecast and Target competitiveness values
function getCompetitiviness(activeSpreadsheet, program) {
  // Define sheet name
  const sheet = 'C-' + program;

  // Define data to return
  var finalData = {};

  // Active spreadsheet values - Actual + Forecast
  const activeSpreadsheetValues = ReturnDataRangeValues(activeSpreadsheet, sheet);

  // Check values
  if (activeSpreadsheetValues.length == 0) {
    // LOG
    console.log('Error accessing spreadsheet with ID ' + activeSpreadsheet);

    // Return empty object
    return finalData;
  }

  // Special case: Program 103 doesn´t have an FTE table
  if (program == '103') {
    // Define data table parameters depending on program
    const plantillaTableTitle = tableDict[sheet][0].title;
    const plantillaRowTitles = tableDict[sheet][0].rowTitles;
    const compTableTitle = tableDict[sheet][1].title;
    const compRowTitles = tableDict[sheet][1].rowTitles;

    // Process values with search script
    var plantillaValues = ProcessSpreadsheetData(plantillaTableTitle, plantillaRowTitles, activeSpreadsheetValues, sheet);
    var compValues = ProcessSpreadsheetData(compTableTitle, compRowTitles, activeSpreadsheetValues, sheet);

    // Add processed values to final data to return
    finalData.PlantillaValues = plantillaValues;
    finalData.CompValues = compValues;
  }
  // Default case:
  else {
    // Define data table parameters depending on program
    const plantillaTableTitle = tableDict[sheet][0].title;
    const plantillaRowTitles = tableDict[sheet][0].rowTitles;
    const compFteTableTitle = tableDict[sheet][1].title;
    const compFteRowTitles = tableDict[sheet][1].rowTitles;
    const compTableTitle = tableDict[sheet][2].title;
    const compRowTitles = tableDict[sheet][2].rowTitles;

    // Process values with search script
    var plantillaValues = ProcessSpreadsheetData(plantillaTableTitle, plantillaRowTitles, activeSpreadsheetValues, sheet);
    var compFteValues = ProcessSpreadsheetData(compFteTableTitle, compFteRowTitles, activeSpreadsheetValues, sheet);
    var compValues = ProcessSpreadsheetData(compTableTitle, compRowTitles, activeSpreadsheetValues, sheet);

    // Add processed values to final data to return
    finalData.PlantillaValues = plantillaValues;
    finalData.CompFteValues = compFteValues;
    finalData.CompValues = compValues;
  }

  // Define month index
  var monthIndex = ReturnMonthIndex(activeSpreadsheet);
  finalData.MonthIndex = monthIndex;

  // LOG
  console.log('(' + sheet + ') Values:');
  console.log('=============================================================================================')
  Object.entries(finalData).forEach(([key, value]) => {
    console.log(key + ': ' + value);
  });
  
  // Return final data
  return finalData;
};

// Function that gets routing reports for all programs
function getReports(activeSpreadsheet) {
  // Define sheet name
  const sheet = 'Actual';

  // Define data to return
  var finalData = {};

  // Active spreadsheet values - Actual + Forecast
  const activeSpreadsheetValues = ReturnDataRangeValues(activeSpreadsheet, sheet);

  // Check values
  if (activeSpreadsheetValues.length == 0) {
    // LOG
    console.log('Error accessing spreadsheet with ID ' + activeSpreadsheet);

    // Return empty object
    return finalData;
  }

  // Define data table parameters
  const timesTableTitle100 = tableDict[sheet]['100'][0].title;
  const timesTableRowTitles100 = tableDict[sheet]['100'][0].rowTitles;
  const reportTitle100 = tableDict[sheet]['100'][1].title;
  const reportRowTitles100 = tableDict[sheet]['100'][1].rowTitles;

  const timesTableTitle101 = tableDict[sheet]['101'][0].title;
  const timesTableRowTitles101 = tableDict[sheet]['101'][0].rowTitles;
  const reportTitle101 = tableDict[sheet]['101'][1].title;
  const reportRowTitles101 = tableDict[sheet]['101'][1].rowTitles;

  const timesTableTitle102 = tableDict[sheet]['102'][0].title;
  const timesTableRowTitles102 = tableDict[sheet]['102'][0].rowTitles;
  const reportTitle102 = tableDict[sheet]['102'][1].title;
  const reportRowTitles102 = tableDict[sheet]['102'][1].rowTitles;

  const timesTableTitle103 = tableDict[sheet]['103'][0].title;
  const timesTableRowTitles103 = tableDict[sheet]['103'][0].rowTitles;
  const reportTitle103 = tableDict[sheet]['103'][1].title;
  const reportRowTitles103 = tableDict[sheet]['103'][1].rowTitles;
  
  // Process actual values with search script
  var timesTableValues100 = ProcessSpreadsheetData(timesTableTitle100, timesTableRowTitles100, activeSpreadsheetValues, sheet);
  var reportValues100 = ProcessSpreadsheetData(reportTitle100, reportRowTitles100, activeSpreadsheetValues, sheet);
  
  var timesTableValues101 = ProcessSpreadsheetData(timesTableTitle101, timesTableRowTitles101, activeSpreadsheetValues, sheet);
  var reportValues101 = ProcessSpreadsheetData(reportTitle101, reportRowTitles101, activeSpreadsheetValues, sheet);

  var timesTableValues102 = ProcessSpreadsheetData(timesTableTitle102, timesTableRowTitles102, activeSpreadsheetValues, sheet);
  var reportValues102 = ProcessSpreadsheetData(reportTitle102, reportRowTitles102, activeSpreadsheetValues, sheet);

  var timesTableValues103 = ProcessSpreadsheetData(timesTableTitle103, timesTableRowTitles103, activeSpreadsheetValues, sheet);
  var reportValues103 = ProcessSpreadsheetData(reportTitle103, reportRowTitles103, activeSpreadsheetValues, sheet);

  // Add actual values to final data to return
  finalData.TimesTableValues100 = timesTableValues100;
  finalData.ReportValues100 = reportValues100;

  finalData.TimesTableValues101 = timesTableValues101;
  finalData.ReportValues101 = reportValues101;

  finalData.TimesTableValues102 = timesTableValues102;
  finalData.ReportValues102 = reportValues102;

  finalData.TimesTableValues103 = timesTableValues103;
  finalData.ReportValues103 = reportValues103;

  // Define month index
  var monthIndex = ReturnMonthIndex(activeSpreadsheet);
  finalData.MonthIndex = monthIndex;

  // LOG
  console.log('(' + sheet + ') Values:');
  console.log('=============================================================================================')
  Object.entries(finalData).forEach(([key, value]) => {
    console.log(key + ': ' + value);
  });

  // Return final data
  return finalData;
};