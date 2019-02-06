/*                              CREATE A PIVOT TABLE
*                                      FOR
*                      ROLES ALLOCATIONS IN MULTIPLE PROJECTS
*
*                                   Coert Vonk
*                                 February 2019
*              https://github.com/cvonk/gas-sheets-projectrolealloc/
*
*
* DESCRIPTION:
*
* Create project role allocation pivot table based on projects names, users and
* their roles.
*
* USE:
*
* The names of tables at the far end of this code.
* Even if you don't ready anything else, please read through the examples below.
*
* DEPENDENCIES:
*
* Requires the Sheets API (otherwise you get: Reference error: Sheets is not
* defined).  To enable the API, refer to 
*   https://stackoverflow.com/questions/45625971/referenceerror-sheets-is-not-defined
* Note that this API has a default quota of 100 reads/day, after that it 
* throws an "Internal Error".
*
* EXAMPLES:
*
* Refer to https://github.com/cvonk/gas-sheets-projectrolealloc/blob/master/README.md
* 
* LEGAL:
*
* (c) Copyright 2019 by Coert Vonk
*
* This program is free software: you can redistribute it and/or modify
* it under the terms of the GNU General Public License as published by
* the Free Software Foundation, either version 3 of the License, or
* (at your option) any later version.
*
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
* GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License
* along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

// update raw sheet based on src sheet

var OnProjectRoleAlloc = {};
(function () {

  // create an associative array that maps the "column label" to the
  // "column index" in the src
  //   key = the column label,
  //   value = the index of the column with that label in the src
  // note: when a "column label" ends in '*', it matches any column that
  // starts with that label

  this.getSrcColumns = function(srcColumnLabels, srcHeader) {

    var srcColumns = [];
    var rawColIdx = 0;

    for each (var label in srcColumnLabels) {

      var matches = [label];
      var wildchard = label.substr(-1) == "*";
      if (wildchard) {
        label = label.slice(0, -1).trim();
        var matches = srcHeader.filter(function(value, index) {
          return value.indexOf(label) === 0;
        });
      }
      var srcColIdx = [];
      for each (var sublabel in matches) {
        srcColIdx.push(srcHeader.indexOf(sublabel));
      };
      srcColumns.push({label: label, src: srcColIdx });
    }
    return srcColumns;
  }

  // recusively walk the src columns to determine the actions needed to create the Raw Sheet

  this.walkSrcColumns = function(srcColumns, action, actions) {

    if (srcColumns.length == 0) {
      actions.push(action);
      return;
    }
    for each (var src in srcColumns[0].src) {
      var actioncopy = action.slice();
      actioncopy.push(src);
      this.walkSrcColumns(srcColumns.slice(1), actioncopy, actions);
    }
    return actions;
  }

  // get Theme Name for a given Project Name

  this.getTheme = function(theValues, prjName) {

    for each (var row in theValues) {
      if (row[0] == prjName) {
        return row[1];
      }
    }
    return undefined;
  }

  this.writeRawHeader = function(lines, theValues, srcColumns) {

    var row = [];
    if (theValues != undefined) {
      row.push("Theme");
    }
    for each (var column in srcColumns) {
      row.push(column.label);
    }
    row.push("Ratio");  // end with Project Allocation column */
    lines.push(row);
  }

  this.writeRawValues = function(lines, srcValues, actions, theValues) {

    var rowNr = 1;
    for each (var srcRow in srcValues) {
      for each (var action in actions) {
        var row = [];
        if (theValues != undefined) {
          var prjName = srcRow[action[0]];
          var prjTheme = this.getTheme(theValues, prjName);
          row.push(prjTheme);
        }
        for each (var src in action) {
          row.push(srcRow[src]);
        }
        if (Common.arrayHasNoEmptyEl(row)) {
          row.push(1);
          lines.push(row);
          var usernameColIdx = lines[0].indexOf("Username");
          var usernameColLetter = String.fromCharCode(65 + usernameColIdx);
          rowNr++;
        }
      }
    }
  }

  this.writeRawAlloc = function(lines) {

    var unameColIdx = lines[0].indexOf("Username");
    var ratioColIdx = lines[0].indexOf("Ratio");
    
    for (var ii = 1; ii < lines.length; ii++) {
      var uname = lines[ii][unameColIdx];
      var linesWithSameUname = lines.filter(function(row) { return row[unameColIdx] === uname }).length;
      lines[ii][ratioColIdx] = 1 / linesWithSameUname;
    }
  }

  this.writeRaw = function(srcColumns, actions, srcValues, rawSheet, theValues) {

    var lines = [];
    this.writeRawHeader(lines, theValues, srcColumns);
    this.writeRawValues(lines, srcValues, actions, theValues);
    this.writeRawAlloc(lines);

    rawSheet.getRange(1, 1, lines.length, lines[0].length).setValues(lines);
  }

  this.getPivotTabelConfig = function(rawSheet, srcColumnLabels) {
    
    // API details at https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/pivot-tables
    var cfg = {
      "source": {
        sheetId: rawSheet.getSheetId(),
        endRowIndex: rawSheet.getDataRange().getNumRows(),
        endColumnIndex: rawSheet.getDataRange().getNumColumns()
      },
      "rows": [{
        sourceColumnOffset: 0,
        showTotals: true,
        //valueBucket: {}
        //label: "Something else instead of Theme",
        sortOrder: "ASCENDING" /*,
        valueMetadata: [{
          "value": { "formulaValue": "true" },   // 2BD
          "collapsed": true
        }] */        
      }],
      "columns": [],
      "values": [{
        sourceColumnOffset: srcColumnLabels.length + 1,
        summarizeFunction: "SUM"
      }]
    };
    
    for (var ii = 1; ii <= srcColumnLabels.length; ii++) {
      switch (ii) {
        case srcColumnLabels.length:
          cfg.columns.push({
            sourceColumnOffset: ii,
            showTotals: true,
            sortOrder: "ASCENDING"
          });
          break;
        default:
          cfg.rows.push({
            sourceColumnOffset: ii,
            showTotals: true,
            sortOrder: "ASCENDING"
          });
      }
    }
    return cfg;
  }    

  this.pivotTableUpdate = function(cfg, pivotTblSheet, spreadsheet) {
    var request = {
      "updateCells": {
        "rows": {
          "values": [{
            "pivotTable": cfg
          }]
        },
        "start": {
          "sheetId": pivotTblSheet.getSheetId()
        },
        "fields": "pivotTable"
      }
    };
    return Sheets.Spreadsheets.batchUpdate({"requests": [request]}, spreadsheet.getId());
    // Enable the Sheets API, or you get: Reference error: sheets is not defined
    // https://stackoverflow.com/questions/45625971/referenceerror-sheets-is-not-defined
  }    
  
  this.createPivotTable = function(spreadsheet, srcColumns, srcColumnLabels, rawSheet, pvtSheetName) {

    if (srcColumnLabels.length < 3) {
      return;
    }
    
    var cfg = this.getPivotTabelConfig(rawSheet, srcColumnLabels);
    var pivotTblSheet = spreadsheet.insertSheet(pvtSheetName);

    return this.pivotTableUpdate(cfg, pivotTblSheet, spreadsheet);
  }

  this.updatePivotTable = function(spreadsheet, rawSheet, pvtSheet) {

    var pvtTableSheetId = pvtSheet.getSheetId();

    // https://sites.google.com/site/scriptsexamples/learn-by-example/google-sheets-api/pivot
    // instead of supplying a whole new pivot table, use the API to get the configuration
    // of the exising Pivot Table, and update the source range
    //    
    var fields = "sheets(properties.sheetId,data.rowData.values.pivotTable)";
    try {
      var response = Sheets.Spreadsheets.get(spreadsheet.getId(), {ranges: "role-alloc", fields: fields});
    } catch (e) {
      if (response == undefined) {  // Internal Error? > you exceeded your quota? (dflt 100 reads/day)
        throw("Google Sheets API read quota exceeded");
      }
    }
    
    var cfg = response.sheets[0].data[0].rowData[0].values[0].pivotTable;
    cfg.source.endRowIndex = rawSheet.getDataRange().getNumRows();

    var request = {
      "updateCells": {
        "rows": {
          "values": [{
            "pivotTable": cfg
          }]
        },
        "start": {
          "sheetId": pvtTableSheetId
        },
        "fields": "pivotTable"
      }
    };

    Sheets.Spreadsheets.batchUpdate({"requests": [request]}, spreadsheet.getId());
  }

  /***
   * @param {string[]} srcColumnLabels Labels of columns in the source sheet that will be output to the raw sheet (labels may include '*' as a wild chard character at the end)
   * @param {string}   srcSheetName    Name of the Source sheet that feeds the Raw sheet
   * @param {string}   rawSheetName    Name of the Raw sheet that feeds the Pivot Table
   * @param {string}   pvtSheetName    Name of the Pivot Table sheet
   * @param {string}   theSheetName    Optional parameter to map Project Names to overarching Theme Names
   */

  this.main = function(srcColumnLabels, srcSheetName, pvtSheetName, theSheetName) {

    // open sheets; copy values

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var srcSheet = Common.sheetOpen(spreadsheet, sheetName = srcSheetName, minNrOfDataRows = 2, requireLabelInA1 = true );
    var srcValues = Common.getFilteredDataRange(srcSheet);
    var srcHeader = srcValues.shift();
    var rawSheetName = pvtSheetName + "-raw";
    var rawSheet = Common.sheetCreate(spreadsheet, sheetName = rawSheetName, overwriteSheet = true).clear();
    var theValues = undefined;
    if (theSheetName != undefined) {
      var theSheet = Common.sheetOpen(spreadsheet, sheetName = theSheetName, minNrOfDataRows = 2, requireLabelInA1 = true );
      theValues = Common.getFilteredDataRange(theSheet);
    }

    // create an array of actions, that each result in a line witten to the Raw Sheet
    // each action is a list of cells to copy from the Source Sheet

    var srcColumns = this.getSrcColumns(srcColumnLabels, srcHeader);
    var actions = this.walkSrcColumns(srcColumns, [], []);

    // write the raw sheet (that drives the pivot table later)

    this.writeRaw(srcColumns, actions, srcValues, rawSheet, theValues);

    // update the pivot table (create if necessary)

    if (spreadsheet.getSheetByName(pvtSheetName) == null) {
      this.createPivotTable(spreadsheet, srcColumns, srcColumnLabels, rawSheet, pvtSheetName);
    }
    var pvtSheet = spreadsheet.getSheetByName(pvtSheetName);
    this.updatePivotTable(spreadsheet, rawSheet, pvtSheet);
  }

}).apply(OnProjectRoleAlloc);


function onProjectRoleAlloc() {

  OnProjectRoleAlloc.main(srcColumnLabels = ["Project Allocation*", "Username", "Role" ],
                          srcSheetName = "persons",
                          pvtSheetName = "role-alloc",
                          theSheetName = "themes");
}
