/* eslint-disable */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("create-table").onclick = createTable;
    document.getElementById("open-dialog").onclick = openDialog;
  }
});

let dialog = null;

function openDialog() {
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/popup.html', {
      height: 45,
      width: 45,
    },
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
  );
}

export async function processMessage(arg) {
  Excel.run(async (context) => {
    try {

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load();
      await context.sync();

      let expensesTable = sheet.tables.getItem('expensesTable');
      const value = [arg.message.split(',')];
      expensesTable.rows.add(null, value);
      dialog.close();
      await context.sync();
    } catch (error) {
      console.error(error);
    }
  })
}

export async function createTable() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/ );
      expensesTable.name = "ExpensesTable";

      expensesTable.getHeaderRowRange().values = [
        ["Date", "Merchant", "Category", "Amount"]
      ];

      const data = [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
      ]

      expensesTable.rows.add(null, data);

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

      sheet.activate();

    });
  } catch (error) {
    console.error(error);
  }
}