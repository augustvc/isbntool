/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");
      range.load("text");
      range.load("values");
      //range.insert(Excel.InsertShiftDirection.right);

      //range.values = [[ "hi" ]];

      // Update the fill color
      //range.format.fill.color = "yellow";

      await context.sync();

      const range2 = range.getColumnsAfter(3);
      range2.load("address");
      range2.load("text");
      range.load("values");
      range.load("columnIndex");
      range.load("rowIndex");
      await context.sync();

      let finalValues = [];

      for (let isbn in range.values) {
        finalValues.push([]);
        finalValues[finalValues.length - 1].push("Auteur(s)");
        finalValues[finalValues.length - 1].push("Titel");
        finalValues[finalValues.length - 1].push("Details");
      }
      
      const CI = range.columnIndex;
      const RI = range.rowIndex;

      let counter = 0;
      for (let idx in range.values) {
        const isbn = range.values[idx];

        const consti = counter;
        counter++;
        
        if(isbn < 1000) {
          continue;
        }

        const http = new XMLHttpRequest();
        http.open("GET", "https://www.googleapis.com/books/v1/volumes?q=isbn:" + isbn);

        http.onreadystatechange = async function () {
          if (this.readyState == 4 && this.status == 200) {
            let responseObj = JSON.parse(http.responseText);

            if(responseObj.items) {
              //console.log("ISBN: " + isbnNumber + " found!");
              let fullTitle = responseObj.items[0].volumeInfo.title;
              if(responseObj.items[0].volumeInfo.subtitle) {
                  fullTitle = fullTitle + " " + responseObj.items[0].volumeInfo.subtitle;
              }
              
              const sheet = context.workbook.worksheets.getFirst();
              
              const authorSpot = sheet.getCell(RI + consti, CI + 1);
              authorSpot.load("values");
              context.sync();
              let authors = "";
              let l = 0;
              for(let autNr of responseObj.items[0].volumeInfo.authors) {
                  if(l > 0) {
                    authors += ", ";
                  }
                  l++;
                  
                  authors += autNr;
              }
              authorSpot.values = [[authors]];
              console.log("Authors: "+ authors);

              const titleSpot = sheet.getCell(RI + consti, CI + 2);
              titleSpot.load("values");
              context.sync();
              titleSpot.values = [[fullTitle]];

              const detailsSpot = sheet.getCell(RI + consti, CI + 3);
              detailsSpot.load("values");
              context.sync();
              let details = "ISBN: " + isbn;
              if(responseObj.items[0].volumeInfo.publisher) {
                details += ", Uitgever: " + responseObj.items[0].volumeInfo.publisher;
              }
              if(responseObj.items[0].volumeInfo.publishedDate) {
                details += ", " + responseObj.items[0].volumeInfo.publishedDate;
              }
              detailsSpot.values = [[details]];
            }
          }
        }
        http.send();
      }

      console.log(`The range address was ${range.address}.`);
      console.log(`The 2nd range address was ${range2.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
