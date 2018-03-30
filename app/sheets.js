var {google} = require('googleapis');
var {OAuth2Client} = require('google-auth-library');
var util = require('util');

var SheetsHelper = function (accessToken) {
    // var authClient = new googleAuth();
    var auth = new OAuth2Client();
    auth.credentials = {
        access_token: accessToken
    };
    this.service = google.sheets({ version: 'v4', auth: auth });
};

// helper.createSpreadsheet in routes.js
SheetsHelper.prototype.createSpreadsheet = function(title, callback) {
  var self = this;
  var request = {
    resource: {
      properties: {
        title: title
      },
      sheets: [
        {
          properties: {
            title: 'Data',
            gridProperties: {
              columnCount: 6,
              frozenRowCount: 1
            }
          }
        },
        {
          properties: {
            title: 'Pivot',
            gridProperties: {
              hideGridlines: true
            }
          }
        }
      ]
    }
  };
  self.service.spreadsheets.create(request, function(err, response) {
    if (err) {
      return callback(err);
    }
    var spreadsheet = response.data;
    // Complete: Add header rows.
    const dataSheetId = spreadsheet.sheets[0].properties.sheetId;
    const pivotSheetId = spreadsheet.sheets[1].properties.sheetId;
    let requests = [
      buildHeaderRowRequest(dataSheetId),
    ];
    // Complete: Add pivot table and chart.
    requests = requests.concat([
      buildPivotTableRequest(dataSheetId, pivotSheetId),
      buildFormatPivotTableRequest(pivotSheetId),
      buildAddChartRequest(pivotSheetId)
    ]);
    const request = {
      spreadsheetId: spreadsheet.spreadsheetId,
      resource: {
        requests: requests
      }
    };
    self.service.spreadsheets.batchUpdate(request, (err, response) => {
      if (err) {
        return callback(err);
      }
      return callback(null, spreadsheet);
    });
  });
};

const COLUMNS = [
  // These key-value pairs correspond to the fields in the Order object (AKA DB columns)
  { field: 'id', header: 'ID' },
  { field: 'customerName', header: 'Customer Name' },
  { field: 'productCode', header: 'Product Code' },
  { field: 'unitsOrdered', header: 'Units Ordered' },
  { field: 'unitPrice', header: 'Unit Price' },
  { field: 'status', header: 'Status' }
]

function buildHeaderRowRequest(sheetId) {
  var cells = COLUMNS.map((col) => {
    return {
      userEnteredValue: { stringValue: col.header },
      userEnteredFormat: { textFormat: { bold: true } }
    }
  });
  // console.log(cells);
  return {
    updateCells: {
      start: { 
        sheetId: sheetId, 
        rowIndex: 0,
        columnIndex: 0
      },
      rows: [
        {
          values: cells
        }
      ],
      fields: 'userEnteredValue,userEnteredFormat.textFormat.bold'
    }
  };
}

// helper.sync in routes.js
SheetsHelper.prototype.sync = function(spreadsheetId, sheetId, orders, callback) {
  const requests = [];
  // Resize the sheet
  requests.push({
    updateSheetProperties: {
      properties: {
        sheetId: sheetId,
        gridProperties: {
          rowCount: orders.length + 1, // To include header
          columnCount: COLUMNS.length
        }
      },
      fields: 'gridProperties(rowCount,columnCount)'
    }
  });
  // Set the cell values
  requests.push({
    updateCells: {
      start: {
        sheetId: sheetId,
        rowIndex: 1,
        columnIndex: 0
      },
      rows: buildRowsForOrders(orders),
      fields: '*'
    }
  });
  // Send the batchUpdate request
  const request = {
    spreadsheetId: spreadsheetId,
    resource: {
      requests: requests
    }
  };
  this.service.spreadsheets.batchUpdate(request, (err) => {
    if (err) {
      return callback(err);
    }
    return callback();
  });
};

// Convert the Order objects to cells
function buildRowsForOrders(orders) {
  return orders.map(order => {
    const cells = COLUMNS.map(col => {
      switch (col.field) {
        case 'unitsOrdered':
          return {
            userEnteredValue: {
              numberValue: order.unitsOrdered
            },
            userEnteredFormat: {
              numberFormat: {
                type: 'NUMBER',
                pattern: '#,##0'
              }
            }
          };
          break;
        case 'unitPrice':
          return {
            userEnteredValue: {
              numberValue: order.unitPrice
            },
            userEnteredFormat: {
              numberFormat: {
                type: 'CURRENCY',
                pattern: '"$"#,##0.00'
              }
            }
          };
          break;
        case 'status':
          return {
            userEnteredValue: {
              stringValue: order.status
            },
            dataValidation: {
              condition: {
                type: 'ONE_OF_LIST',
                values: [
                  { userEnteredValue: 'PENDING' },
                  { userEnteredValue: 'SHIPPED' },
                  { userEnteredValue: 'DELIVERED' }
                ]
              },
              strict: true,
              showCustomUi: true
            }
          };
          break;
        default: 
          return {
            userEnteredValue: {
              stringValue: order[col.field].toString()
            }
          };
      }
    });
    return {
      values: cells
    };
  });
}

// Create requests for building the pivot table, formatting the results, and adding the chart
function buildPivotTableRequest(sourceSheetId, targetSheetId) {
  return {
    updateCells: {
      start: { sheetId: targetSheetId, rowIndex: 0, columnIndex: 0 },
      rows: [
        {
          values: [
            {
              pivotTable: {
                source: {
                  sheetId: sourceSheetId,
                  startRowIndex: 0,
                  startColumnIndex: 0,
                  endColumnIndex: COLUMNS.length
                },
                rows: [
                  {
                    sourceColumnOffset: getColumnForField('productCode').index,
                    showTotals: false,
                    sortOrder: 'ASCENDING'
                  }
                ],
                values: [
                  {
                    summarizeFunction: 'SUM',
                    sourceColumnOffset: getColumnForField('unitsOrdered').index
                  },
                  {
                    summarizeFunction: 'SUM',
                    name: 'Revenue',
                    formula: util.format("='%s' * '%s'", 
                      getColumnForField('unitsOrdered').header, 
                      getColumnForField('unitPrice').header)
                  }
                ]
              }
            }
          ]
        }
      ],
      fields: '*'
    }
  };
}

function buildFormatPivotTableRequest(sheetId) {
  return {
    repeatCell: {
      range: { sheetId: sheetId, startRowIndex: 1, startColumnIndex: 2 },
      cell: {
        userEnteredFormat: {
          numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }
        }
      },
      fields: 'userEnteredFormat.numberFormat'
    }
  };
}

function buildAddChartRequest(sheetId) {
  return {
    addChart: {
      chart: {
        spec: {
          title: 'Revenue per Product',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'RIGHT_LEGEND',
            domains: [
              // Show a bar for each product code in the pivot table.
              {
                domain: { sourceRange: { sources: [{
                  sheetId: sheetId,
                  startRowIndex: 0,
                  startColumnIndex: 0,
                  endColumnIndex: 1
                }]}}
              }
            ],
            series: [
              {
                series: { sourceRange: { sources: [{
                  sheetId: sheetId,
                  startRowIndex: 0,
                  startColumnIndex: 2,
                  endColumnIndex: 3
                }]}}
              }
            ]
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: sheetId, rowIndex: 0, columnIndex: 3 },
            widthPixels: 600,
            heightPixels: 400
          }
        }
      }
    }
  };
}

function getColumnForField(field) {
  return COLUMNS.reduce((result, col, i) => {
    if (col.field == field) {
      col.index = i;
      return col;
    }
    return result;
  });
}

module.exports = SheetsHelper;

