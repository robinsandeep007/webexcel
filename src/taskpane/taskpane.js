/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    document.getElementById("protect").onclick = protect;

    document.getElementById("unprotect").onclick = unprotect;
  }
});

export async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}


export async function protect() {
  try {
    let password = "password";
    await Excel.run(async (context) => {
      let activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("protection/protected");
    
      await context.sync();
    
      if (!activeSheet.protection.protected) {
          activeSheet.protection.protect(null, password);
      }
      console.log(`the workbook is protected.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function unprotect() {
  try {
    let password = "password";
    await Excel.run(async (context) => {
      let activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("protection/protected");
    
      await context.sync();
    
      if (activeSheet.protection.protected) {
          activeSheet.protection.unprotect(password);
      }
      console.log(`the workbook is unprotected.`);
    });
  } catch (error) {
    console.error(error);
  }
}


