/** Transforms raw invoice data in 'Data Drop' tab into usable data throughout the AR Automation Sheet */
function cleanData() {
  const ss = SpreadsheetApp.getActive()
  const dataTab = ss.getSheetByName('Data Drop')
  const bucket1 = ss.getSheetByName('30-40')
  const bucket2 = ss.getSheetByName('41-60')
  const bucket3 = ss.getSheetByName('61-100')
  const bucket4 = ss.getSheetByName('100+')
 
  // Finds rows of invoices that need to be processed. They'll always have five digits and a balance (vs. a credit).
  const nums = dataTab.getRange(1, 1, dataTab.getLastRow()).getValues()
  let invoiceRows = []
  for (var i = 0; i < nums.length; i++) {
    if (nums[i].toString().length == 5) {
      (dataTab.getRange(i + 1, 23).getValue() > 0) ? invoiceRows.push(i + 1) : null
    }
  }
  
  //////// FINDING WHICH INVOICES AREN'T IN THE SHEET YET ////////
  // Builds an array of existing invoices in the Sheet
  const existingInvoicesInBucket1 = bucket1.getRange(2, 1, bucket1.getLastRow()).getValues()
  const existingInvoicesInBucket2 = bucket2.getRange(2, 1, bucket2.getLastRow()).getValues()
  const existingInvoicesInBucket3 = bucket3.getRange(2, 1, bucket3.getLastRow()).getValues()
  const existingInvoicesInBucket4 = bucket4.getRange(2, 1, bucket3.getLastRow()).getValues()
  const existingInvoices = existingInvoicesInBucket1.concat(existingInvoicesInBucket2, existingInvoicesInBucket3, existingInvoicesInBucket4).filter(invoice => invoice[0] !== '')
  
  // Finds invoices not in the Sheet yet.
  invoiceRows.forEach(function(row) {
    let invoiceNo = dataTab.getRange(row, 1).getValue()[0]
    if (!existingInvoices.some(number => invoiceNo == number)) { 
      // Gets data for the invoice.
      let data = dataTab.getRange(row, 1, 1, dataTab.getLastColumn()).getValues()[0]
      let number = data[0]
      let amount = data[22]
      let dueDate = data[18]
      let daysOverdue = getInvoiceAge(dueDate)
      let firmName = data[4]
      let pocName = data[2]
      let firstName = destructureName(pocName)[0] ? destructureName(pocName)[0] : null
      let lastName = destructureName(pocName)[1] ? destructureName(pocName)[1] : null
      let address1 = data[6]	
      let address2 = data[8]
      let city = data[10]
      let state	= data[12]
      let zip = data[14]
      
      // Finds the correct Sheet, then prints to the data to the final + 1 row in the correct columns.
      const bucket = findBucket(daysOverdue)
      let printRow = bucket.getLastRow() + 1
      bucket.getRange(printRow, 1).setValue(number)
      bucket.getRange(printRow, 2).setValue(amount)
      bucket.getRange(printRow, 3).setValue(dueDate)
      bucket.getRange(printRow, 4).setValue(daysOverdue)
      bucket.getRange(printRow, 5).setValue(firmName)
      bucket.getRange(printRow, 6).setValue(firstName)
      bucket.getRange(printRow, 7).setValue(lastName)
      bucket.getRange(printRow, 24).setValue(address1)
      bucket.getRange(printRow, 25).setValue(address2)
      bucket.getRange(printRow, 26).setValue(city)
      bucket.getRange(printRow, 27).setValue(state)
      bucket.getRange(printRow, 28).setValue(zip)
    }
  })  
  
  SpreadsheetApp.getActiveSpreadsheet().toast('✔️️ Invoice data organization complete.');
  
  // Look for invoices that were in the AR hub that aren't in this new list, removes them (this indicates they were paid).
}
                      
// Looks for emails of matching POC and firm names in Automation Station


///////////////////////////////////////////////
//////////////// UTILITIES ////////////////////
///////////////////////////////////////////////

/** Returns the age (in # of days) of an invoice.
* @param {invoiceDate} obj Date object.
*/
function getInvoiceAge(invoiceDate) {
  const ss = SpreadsheetApp.getActive()
  const dataTab = ss.getSheetByName('Data Drop')
  
  const now = new Date()
  const diffTime = Math.abs(now - invoiceDate)
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24))
  
  return diffDays
}

/** Returns (array) first and last name from strings.
* @param {fullName} string Full name e.g. José Valdez.
* @return array Array of results.
*/
function destructureName(fullName) {
  if (!fullName) {
    return false
  }
  let names = fullName.split(/\s+/)
  if (names.length === 0) {
    return [fullName]
  } else if (names.length === 1) {
    return names
  } else if (names.length === 2) {
    return [names[0], names[1]]
  } else {
    return [names[0], names[names.length - 1]]
  }
}

/** Determines which tab an invoice should go to based on its age.
* @param {age} number How old (by the number of days) an invoice is.
* @return {obj} Reference to the correct Sheet.
*/
function findBucket(age) {
  const ss = SpreadsheetApp.getActive()
  const bucket1 = ss.getSheetByName('30-40')
  const bucket2 = ss.getSheetByName('41-60')
  const bucket3 = ss.getSheetByName('61-100')
  const bucket4 = ss.getSheetByName('100+')
  if (age <= 40) {
    console.log('1')
    return bucket1
  } else if (age > 40 && age <= 60) {
    console.log('2')
    return bucket2 
  } else if (age > 60 && age <= 100) {
    console.log('3')
    return bucket3
  } else {
    console.log('4')
    return bucket4
  }
  
}