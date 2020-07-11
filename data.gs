/** Transforms raw invoice data in 'Data Drop' tab into usable data throughout the AR Automation Sheet */
function cleanData() {
  const ss = SpreadsheetApp.getActive()
  const dataTab = ss.getSheetByName('Data Drop')
  
}


///////////////////////////////////////////////
//////////////// UTILITIES ////////////////////
///////////////////////////////////////////////

/** Returns the age (in # of days) of an invoice. */
function getInvoiceAge() {
  const ss = SpreadsheetApp.getActive()
  const dataTab = ss.getSheetByName('Data Drop')
  
  const invoiceDate = dataTab.getRange(3, 19).getValue()
  
  const invoice = new Date(invoiceDate)
  const now = new Date()
  const diffTime = Math.abs(now - invoice)
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24))
  console.log(diffDays + " days")
}


/** Returns a value for a given month.
@param {month} Three-letter month string (e.g. 'Mar').
*/
function getMonthValue(month) {
  switch(month) {
    case 'Jan':
      return 1;
    case 'Feb':
      return 2;
    case 'Mar':
      return 3;
    case 'Apr':
      return 4;
    case 'May':
      return 5;
    case 'Jun':
      return 6;
    case 'Jul':
      return 7;
    case 'Aug':
      return 8;
    case 'Sep':
      return 9;
    case 'Oct':
      return 10;
    case 'Nov':
      return 11;
    case 'Dec':
      return 12;
  }
}