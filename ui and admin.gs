/** SA Legal Solutions Accounts Receivable Automation Station Codebase
* GitHub repo https://github.com/ydax/SA-Legal-Solutions-AR-Automations
* @dev Davis Jones | github.com/ydax | davis@eazl.co
* SA Legal Solutions POC Blake Boyd | bboyd@salegalsolutions.com
* Color Palette
  Use              HEX     MaterializeCSS
  Primary          #c62828 red darken-3
  Cell Background  #ffebee red lighten-5
  Confirmation     #00796b teal darken-2
  Error            #e65100 orange darken-4
  Primary -1       #e53935 red darken-1
  Primary -2       #e57373 red lighten-2
  Primary +1       #b71c1c red darken-4
*/

////////////////////////////////////////////////////////////////////////////////////
////////////// CREATION OF SPREADSHEET MENU PLUS USER INTERFACE CALLS //////////////
////////////////////////////////////////////////////////////////////////////////////

/** Creates the SA Legal Solutions menu.
@param {e} object Sheet load event object.
*/
function onOpen (e) {
  var ui = SpreadsheetApp.getUi();  
  ui.createMenu("⚖️ SA Legal Solutions")
  .addItem("🔁 Repeat Orderer", "sayHello")
  .addToUi();
};

function sayHello() {
  console.log('hello')
}


//////////////////////////////////////////////
/////////// DEVELOPMENT ROADMAP //////////////
//////////////////////////////////////////////

/** Development plan
Get the data pulling in from CSV properly
Send function combines all the reminders for a paralegal into one email
When data is imported, if an outstanding balance isn't in the new dump, it removes it from the sheet
Midnight function that finds emails of people in the AR sheet from Automation Station
Send an email or make a pdf / letter based on what Blake wants.
*/


















