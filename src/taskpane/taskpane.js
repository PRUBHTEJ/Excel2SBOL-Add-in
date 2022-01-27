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
      range.load(['address', 'values']);
      await context.sync();
      var firstSelectedCellValue = range.values[0][0];
      //Iterating through the JS 2-D Array
      var size = range.values.length;
      globalThis.fetch = fetch;

      for(let i = 0; i<size; i++){
        var name = range.values[i][0];
        console.log(name);
        var URI = range.values[i][1];
        console.log(URI);
        const query = `SELECT distinct ?label
            WHERE
            {{{{
                optional
                {{{{
                    <${URI}> rdfs:label ?label .
                     filter langMatches(lang(?label), "en")
                }}}}
                optional
                {{{{
                    <${URI}> rdfs:label ?label .
                }}}}
            }}}}`
        const result = await fetch("http://sparql.hegroup.org/sparql", {
          method: "POST", 
          headers: {
            Accept: "application/json"
          }
        });
        const data = await result.json();
        console.log(data);
      }
      //console.log(`First Selected Value ${firstSelectedCellValue}`);
      // Update the fill color
      range.format.fill.color = "yellow";
      
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
