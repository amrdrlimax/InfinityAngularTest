import { Component, NgZone, Output } from "@angular/core";
//import * as excel from "./excel.app.component";
//import * as selectionForm from "selection-form/selectionForm.component";
import { platformBrowserDynamic } from "@angular/platform-browser-dynamic";
//import * as OfficeHelpers from "@microsoft/office-js-helpers";
import { ExcelTableUtil } from "../../../src/utils/excelTableUtils";
import { AppModule } from "./app.module";
//import { FormGroup, FormControl } from '@angular/forms';


const ALPHAVANTAGE_APIKEY = "{{VVI23S5D0LRGBXOJ}}";



Office.initialize = () => {

  //this.formName = document.getElementById("username");

  console.log('line one',"nome");
  // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
  if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
    console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
  }
  

  // Bootstrap the app
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch((error) => console.error(error));
};



@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})



// export class ProfileEditorComponent {
//   profileForm = new FormGroup({
//     firstName: new FormControl(''),
//     lastName: new FormControl(''),
//   });
// }

export default class AppComponent {
  /*  welcomeMessage = "Ola";
 
   async run() {
     const excelComponent = new excel.default();
     return excelComponent.run();
   }
   async runSelectionForm() {
     const selectionForm = new selectionForm.de();
     return selectionForm.run();
   } */

  

  symbols = [];
  error = null;
  waiting = false;
  zone = new NgZone({});
  formName = "";

  tableUtil = new ExcelTableUtil("Portfolio", "A1:H1", [
    "Symbol",
    "Last Price",
    "Timestamp",
    "Quantity",
    "Price Paid",
    "Total Gain",
    "Total Gain %",
    "Value",
  ]);

  constructor() {
    this.syncTable().then(() => { });    
  }
  

  clickme(username:string, email:string, message:string) {
    console.log('it does nothing',username, email, message);
    this.waiting = true;
    var teste = username;

    console.log(teste);

    var data = [
      [teste], //Symbol
      [email], //Last Price
      [message], // Timestamp of quote,
      0, // quantity (manually entered)
      0, // price paid (manually entered)
      ["a"], //Total Gain $
      ["s"], //Total Gain %
      ["d"], //Value
    ];

    console.log(data);

    //this.tableUtil.addRow.(data);
    this.tableUtil.addRow(data).then(
      () => {      
        this.waiting = false;
      },
      (err) => {
        this.error = err;
      }
    );
  }

  testSend = async (username:string) => {

   // console.warn(newHero.value);

    this.waiting = true;

    //let nome = (document.getElementById("newHero") as HTMLInputElement).value;

    const data = [
      [username], //Symbol
      ["teste2"], //Last Price
      ["teste3"], // Timestamp of quote,
      0, // quantity (manually entered)
      0, // price paid (manually entered)
      ["a"], //Total Gain $
      ["s"], //Total Gain %
      ["d"], //Value
    ];


    this.tableUtil.addRow(data).then(
      () => {
        //this.symbols.unshift(symbol.toUpperCase());
        this.waiting = false;
        //this.syncTable().then(() => { });  
      },
      (err) => {
        this.error = err;
      }
    );

    this.syncTable().then(() => { }); 

    // this.tableUtil.updateCell().then(
    //   async () => {
    //     this.waiting = false;
    //   },
    //   (err) => {
    //     this.error = err;
    //     this.waiting = false;
    //   }
    // );

  }

  // Adds symbol
  addSymbol = async (symbol) => {
    this.waiting = true;

    // Get quote and add to Excel table
    this.getQuote(symbol).then(
      (res) => {
        let cnt = this.symbols.length;
        const data = [
          res["01. symbol"], //Symbol
          res["05. price"], //Last Price
          res["07. latest trading day"], // Timestamp of quote,
          0, // quantity (manually entered)
          0, // price paid (manually entered)
          `=(B${cnt + 2} * D${cnt + 2}) - (E${cnt + 2} * D${cnt + 2})`, //Total Gain $
          `=H${cnt + 2} / (E${cnt + 2} * D${cnt + 2}) * 100 - 100`, //Total Gain %
          `=B${cnt + 2} * D${cnt + 2}`, //Value
        ];
        this.tableUtil.addRow(data).then(
          () => {
            this.symbols.unshift(symbol.toUpperCase());
            this.waiting = false;
          },
          (err) => {
            this.error = err;
          }
        );
      },
      (err) => {
        this.error = err;
        this.waiting = false;
      }
    );
  };

  // Delete symbol
  deleteSymbol = async (index) => {
    // Delete from Excel table by index number
    const symbol: string = this.symbols[index];
    this.waiting = true;
    this.tableUtil.getColumnData("Symbol").then(
      async (columnData: string[]) => {
        // Ensure the symbol was found in the Excel table
        if (columnData.indexOf(symbol) !== -1) {
          this.tableUtil.deleteRow(columnData.indexOf(symbol)).then(
            async () => {
              this.symbols.splice(index, 1);
              this.waiting = false;
            },
            (err) => {
              this.error = err;
              this.waiting = false;
            }
          );
        } else {
          this.symbols.splice(index, 1);
          this.waiting = false;
        }
      },
      (err) => {
        this.error = err;
        this.waiting = false;
      }
    );
  };

  // Refresh symbol
  refreshSymbol = async (index) => {
    // Refresh stock quote and update Excel table
    const symbol = this.symbols[index];
    this.waiting = true;
    this.tableUtil.getColumnData("Symbol").then(
      async (columnData: string[]) => {
        // Ensure the symbol was found in the Excel table
        const rowIndex = columnData.indexOf(symbol);
        if (rowIndex !== -1) {
          this.getQuote(symbol).then((res) => {
            // "last trade" is in column B with a row index offset of 2 (row 0 + the header row)
            this.tableUtil.updateCell(`B${rowIndex + 2}:B${rowIndex + 2}`, res["05. price"]).then(
              async () => {
                this.waiting = false;
              },
              (err) => {
                this.error = err;
                this.waiting = false;
              }
            );
          });
        } else {
          this.error = `${symbol} not found in Excel`;
          this.symbols.splice(index, 1);
          this.waiting = false;
        }
      },
      (err) => {
        this.error = err;
        this.waiting = false;
      }
    );
  };

  // Reads symbols from an existing Excel workbook and pre-populates them in the add-in
  syncTable = async () => {
    this.waiting = true;
    this.tableUtil.getColumnData("Symbol").then(
      async (columnData: string[]) => {
        this.symbols = columnData;
        this.waiting = false;
      },
      (err) => {
        this.error = err;
        this.waiting = false;
      }
    );
  };

  // Gets a quote by calling into the stock service
  getQuote = async (symbol) => {
    return new Promise((resolve, reject) => {
      const queryEndpoint = `https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=${escape(
        symbol
      )}&apikey=${ALPHAVANTAGE_APIKEY}`;

      fetch(queryEndpoint)
        .then((res) => {
          if (!res.ok) {
            reject("Error getting quote");
          }
          return res.json();
        })
        .then((jsonResponse) => {
          const quote = jsonResponse["Global Quote"];
          resolve(quote);
        });
    });
  };
}
