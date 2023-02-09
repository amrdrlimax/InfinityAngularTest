/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import "zone.js"; // Required for Angular
import { platformBrowserDynamic } from "@angular/platform-browser-dynamic";
import { AppModule } from "./app/app.module";
//import { Message } from "@angular/compiler/src/i18n/i18n_ast";
/* global console, document, Office */

Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";

  // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
  if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
    console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
  }

  // Assign event handlers and other initialization logic.
  // document.getElementById("create-table").onclick = createTable;
  // document.getElementById("filter-table").onclick = filterTable;
  // document.getElementById("sort-table").onclick = sortTable;
  // document.getElementById("create-chart").onclick = createChart;
  // document.getElementById("freeze-header").onclick = freezeHeader;
  // document.getElementById("teste").onclick = testGet;
  

  // Bootstrap the app
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch((error) => console.error(error));
};

async function createTable() {
  await Excel.run(async (context) => {

    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A10:D10", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values =
    [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);

    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function filterTable() {
  await Excel.run(async (context) => {

    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function sortTable() {
  await Excel.run(async (context) => {

      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
      const sortFields = [
          {
              key: 1,            // Merchant column
              ascending: false,
          }
      ];
      
      expensesTable.sort.apply(sortFields);

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}


async function createChart() {
  await Excel.run(async (context) => {

    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const dataRange = expensesTable.getDataBodyRange();

    const chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'Auto');

    chart.setPosition("A20", "F35");
    chart.title.text = "Expenses";
    chart.legend.position = "Right";
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in \u20AC';

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function freezeHeader() {
  await Excel.run(async (context) => {

    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);

      await context.sync();
  }) 
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function testGet() {
    await Excel.run(async (context) => { 
      alert("teste");
      
      const binding = context.workbook.bindings.getItemAt(0);
      const text = binding.getText();
      binding.load('text');
      
      await context.sync();

      console.log(text);
    });
}