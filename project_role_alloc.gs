/*                              CREATE A PIVOT TABLE
*                                      FOR
*                      ROLES ALLOCATIONS IN MULTIPLE PROJECTS
*
*                                   Coert Vonk
*                                 February 2019
*              https://github.com/cvonk/gas-sheets-project_role_alloc/
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
  
  // create an associative array that maps the "column label" to the "column index" in the src
  //   key = the column label,
  //   value = the index of the column with that label in the src
  // note: when a "column label" ends in '*', it matches any column that starts with that label
  // e.g.
  //   labels = ["Project Allocation*", "Username", "Role"]
  //   srcHeader = ["Username", "Preferred Name", "Person Type", "Role", "Project Allocation 1", "Project Allocation 1 %", "Project Allocation 2"]
  // will output
  //   // [{label:"Project Allocation", src:[4, 5]}, {label:"Username", src:[0]}, {label:"Role", src:[3]}]
  //   [{label:"Project Allocation", idx:[{val:4, pct:5}, {val:6, pct:-1}]}, {label:"Username", idx:[{val:0, pct:-1}]}, {label:"Role", idx:[{val:3, pct:-1}]}]
  
  this.getSrcColumns = function(labels, srcHeader) {
    
    String.prototype.startsWith = function(prefix) { 
      return this.indexOf(prefix) === 0; 
    } 
    String.prototype.endsWith = function(suffix) { 
      return this.match(suffix+"$") == suffix; 
    };    
    
    var srcColumns = [];
    var percentageStr = " %";
    
    for each (var label in labels) {
      
      if (label.substr(-1) == "*") {
        label = label.slice(0, -1).trim();
      }
      var valueLabels = srcHeader.filter(function(srcLabel) {
        return srcLabel.startsWith(label) && !srcLabel.endsWith(percentageStr);
      });
      var srcColIdx = [];
      for each (var valueLabel in valueLabels) {
        srcColIdx.push({val:srcHeader.indexOf(valueLabel), pct:srcHeader.indexOf(valueLabel + percentageStr)});
      };
      srcColumns.push({label: label, idx:srcColIdx});
    }
    return srcColumns;
  }
  
  // recusively walk the Src Columns to determine the actions needed to create the Raw Sheet
  // e.g. at the root level
  //  srcColumns = [{label:"Project Allocation", idx:[{val:4, pct:5}, {val:6, pct:-1}]}, {label:"Username", idx:[{val:0, pct:-1}]}, {label:"Role", idx:[{val:3, pct:-1}]}]
  //  action = []
  //  actions = []
  // will output [[{val:4, pct:5}, {val:0, pct:-1}, {val:3, pct:-1}], [{val:6, pct:-1}, {}, {}]]
  //   
  this.walkSrcColumns = function(srcColumns, action, actions) {

    if (srcColumns.length == 0) {
      actions.push(action);
      return;
    }
    for each (var idx in srcColumns[0].idx) {
      var actioncopy = action.slice();
      actioncopy.push(idx);
      this.walkSrcColumns(srcColumns.slice(1), actioncopy, actions);
    }
    return actions;
  }

  this.getRawHeader = function(theValues, srcColumns) {

    var row = [];
    if (theValues != undefined) {
      row.push("Theme");
    }
    row.push("Ratio");  // 2BD should be hardcoded here, but depend on something like actionLvlHasAssignedPercentages
    for each (var column in srcColumns) {
      row.push(column.label);
    }
    return row;
  }

  this.getRatio = function(srcRow, actions, idx, idxNr) {
    
    var assignedCnt = 0, assignedVal = 0, totalCnt = 0;
    for each (var act in actions) {        
      if (act[idxNr].pct >= 0) {
        var val = srcRow[act[idxNr].pct];
        if (val) { // skip blank
          assignedCnt++;
          assignedVal += val;
          if (assignedVal > 1) {
            throw("over 100% assigned for (" + srcRow[0]  + ")");
          }
        }
      }
      if (srcRow[act[idxNr].val]) {
        totalCnt++; // only count the columns with values
      }
    }
    if (assignedCnt) {
      if (idx.pct >= 0) {
        return srcRow[idx.pct];
      } else {
        return (1 - assignedVal) / (totalCnt - assignedCnt);
      }            
    }
    return 1.0 / totalCnt;
  }  
  
  this.getActionIdxsThatHaveAssignedPercentages = function(srcValues, actions, idx) {
    
    var result = [];
    for each (var srcRow in srcValues) {
      for each (var action in actions) {  
        var actionLvl = 0;
        for each (var idx in action) {
          if (idx.pct >= 0 && srcRow[idx.pct]) {
            result.push(actionLvl);
          }
          actionLvl++;
        }
      }
    }
    return result;
  }
  
  // write the values to the raw sheet.
  // if users work on >1 prject, divvy up their allocation.
  //
  // e.g.
  //   lines = [["Theme", "Ratio", "Project Allocation", "Username", "Role"]]
  //   srcValues = [["jvonk", "Johan", "Employee", "Student", "School", 0.8, "Java"], ["svonk", "Sander", "Employee", "Student", "School", "", "Reading"], ["brlevins", "Barrie", "Employee", "Adult", "BoBo", "", ""]]
  //   actions = [[{val:4, pct:5}, {val:0, pct:-1}, {val:3, pct:-1}], [{val:6, pct:-1}, {}, {}]]
  // returns updated lines
  //   lines = [["Theme", "Ratio", "Project Allocation", "Username", "Role"], [, 0.8, "School", "jvonk", "Student"], [, 0.2, "Java", "jvonk", "Student"], [, 0.5, "School", "svonk", "Student"], [, 0.5, "Reading", ...

  this.writeRawValues = function(rawHeader, srcValues, actions) {

    var lines = [];
    var actionLvlsWithAssignedPercentages = this.getActionIdxsThatHaveAssignedPercentages(srcValues, actions);    
        
    var rowNr = 1;
    for each (var srcRow in srcValues) {
      
      var ratioColIdx = rawHeader.indexOf("Ratio");
      
      for each (var action in actions) {
        var row = [,], alloc = 1, idxNr = 0;
        
        for each (var idx in action) {
          var ratio = this.getRatio(srcRow, actions, idx, idxNr);
          
          if( actionLvlsWithAssignedPercentages.indexOf(idxNr) >= 0) {
            row.push(Number(ratio.toFixed(2)));  // hide math precision err
          }
          row.push(srcRow[idx.val]);
          idxNr++;
        }
        if (row[2]) {        // 2BD shouldn't be hardcoded
          lines.push(row);
          rowNr++;
        }
      }
    }
    return lines;
  }

  // first row is supposed to be a header, and is skipped
  // reads the project name from the column with index "thePrjIdx", and write the
  // corresponding theme name to the map it to the column with index "theColIdx"
  
  this.writeRawThemeColumn = function(lines, theValues, prjColIdx, theColIdx) {

    function _getTheme(theValues, prjName) {
      for each (var row in theValues) {
        if (row[0] == prjName) {
          return row[1];
        }
      }
      return undefined;
    }
    
    if (theValues == undefined) {
      return;
    }

    for (var ii = 1; ii < lines.length; ii++) {
      var line = lines[ii];
      var prjTheme = _getTheme(theValues, line[prjColIdx]);
      lines[ii][theColIdx] = prjTheme;
    }
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
        //label: "Something else instead of Theme",
        sortOrder: "ASCENDING"
      }],
      "columns": [],
      "values": [{
        sourceColumnOffset: 1,   // 2BD not hardcoded
        summarizeFunction: "SUM"
      }]
    };
    
    for (var ii = 1; ii <= srcColumnLabels.length; ii++) {
      switch (ii) {
        case srcColumnLabels.length:
          cfg.columns.push({
            sourceColumnOffset: ii + 1,
            showTotals: true,
            sortOrder: "ASCENDING"
          });
          break;
        default:
          cfg.rows.push({
            sourceColumnOffset: ii + 1,
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
    var fields = "sheets.data.rowData.values.pivotTable";
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

    // create an array of actions
    // each action is a list of cells to copy from the Source Sheet

    var srcColumns = this.getSrcColumns(srcColumnLabels, srcHeader);
    var actions = this.walkSrcColumns(srcColumns, [], []);

    // write the raw sheet (that drives the pivot table later)

    var rawHeader = this.getRawHeader(theValues, srcColumns);
    var rawValues = this.writeRawValues(rawHeader, srcValues, actions);
    var rawData = [rawHeader].concat(rawValues);
    this.writeRawThemeColumn(rawData, theValues, 2, 0);
    rawSheet.getRange(1, 1, rawData.length, rawData[0].length).setValues(rawData);
    
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
