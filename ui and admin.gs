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

function showLastRow() {
  var lastRow = SpreadsheetApp.getActive().getSheetByName('30-40').getLastRow()
  console.log(lastRow)
}

////////////////////////////////////////////////////////////////////////////////////
////////////// CREATION OF SPREADSHEET MENU PLUS USER INTERFACE CALLS //////////////
////////////////////////////////////////////////////////////////////////////////////

/** Creates the SA Legal Solutions menu.
@param {e} object Sheet load event object.
*/
function onOpen (e) {
  var ui = SpreadsheetApp.getUi();  
  ui.createMenu("⚖️ SA Legal Solutions")
  .addItem("📊 Import Data", "cleanData")
  .addToUi();
};

/** Adds an error to the developer tab.
* @param {message} string Error message generated by the program.
*/
function addError(message){
  const devSheet = SpreadsheetApp.getActive().getSheetByName('Developer')
  const timestamp = new Date().toISOString()
  const array = [timestamp, message]
  devSheet.getRange(devSheet.getLastRow() + 1, 1, 1, 2).setValues([array])
}


//////////////////////////////////////////////
/////////// DEVELOPMENT ROADMAP //////////////
//////////////////////////////////////////////

/** Development plan
Done July 11, 2020
X Get the data pulling in from CSV properly

TODO July 18, 2020
Have invoices NOT found in new data deleted
Create email status column. When changed to "Send", have it send an email to the POC, log a date, then change status to "Sent".
X Finds emails of people in the AR sheet from Automation Station on data import
X Menu function that cleans out the "Data Drop" sheet
AS: "I’m not finding the depos on the current list that the emails are sending to change them. Where do they go after they’re removed from the current list?"
Out of curiosity, I just checked on this. If the rows are hidden, then a General Find function won’t actually locate the depo / be able to search the cells.

Send function combines all the reminders for a paralegal into one email
Send an email or make a pdf / letter based on what Blake wants.

Hey Davis -  Hope your travels are going well!  I’m sending this email so that I don’t forget. No need for you to reply or even think about it at this point. 

If attorney John Smith has 3 invoices overdue I need to have a way to send all three invoices in one email.  If I sent three separate emails they might just block my emails like it’s spam. 

Then at the same time at 60 days, if 2 of the invoices are paid and one isn’t I don’t want to send them the two that are paid.

Figured this was a big enough issue to address before you start! 
*/


















