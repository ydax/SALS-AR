/** Transforms raw invoice data in 'Data Drop' tab into usable data throughout the AR Automation Sheet */
function cleanData() {
  SpreadsheetApp.getActiveSpreadsheet().toast('🟢️️ Started organizing invoice data.')
  try {
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
    const existingInvoices = getExistingInvoiceNumbers()
    
    // Finds invoices not in the Sheet yet.
    invoiceRows.forEach(function(row) {
      let invoiceNo = dataTab.getRange(row, 1).getValue()
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
        bucket.getRange(printRow, 10).setValue('Not Sent')
      }
    })
    
    // Attempts to match POCs with previous orderers in AS.
    matchEmails()
    
    // Sets data validation for all rows in invoice tabs.
    SpreadsheetApp.getActiveSpreadsheet().toast('🕹️️ Setting up data validation.')
    const invoiceTabs = [bucket1, bucket2, bucket3, bucket4]
    const validationRange = ss.getSheetByName('Developer').getRange(2, 1, 3, 1)
    invoiceTabs.forEach(function(tab) {
      const emailCells = tab.getRange(2, 10, tab.getLastRow(), 1)
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build()
      emailCells.setDataValidation(rule)
    })
    
    // Look for invoices that were in the AR hub that aren't in this new list, removes them (this indicates they were paid).
    const incomingInvoiceNumbers = nums.filter(elem => typeof elem[0] === 'number' && elem[0].toString().length === 5).map(elem => elem[0])
    removePaidInvoices(incomingInvoiceNumbers, existingInvoices)
    
    SpreadsheetApp.getActiveSpreadsheet().toast('✔️️ Invoice data organization complete.')
  } catch (error) {
    addError(error)
    SpreadsheetApp.getActiveSpreadsheet().toast('⚠️️️ There was an error organizing the data. This error has been recorded.')
  }
}

/** Removes invoices that aren't in the new data drop, indicating they're paid.
* @param {incomingInvoiceNumbers} array An array of incoming invoice numbers from the Data Drop Sheet.
* @param {existingInvoices} array An array of all the invoice numbers currently in the AR station.
*/
function removePaidInvoices(incomingInvoiceNumbers, existingInvoices) {
  SpreadsheetApp.getActiveSpreadsheet().toast(`🟢️️ Started locating and removing paid invoices.`)
  const ss = SpreadsheetApp.getActive()
  // Properly formats the existing invoices array into a 2D array of invoice numbers (same as the incomingInvoiceNumbers).
  existingInvoices = existingInvoices.map(elem => elem[0])
  // Pushes any invoice numbers that are in the existingInvoices, but not in the incomingInvoices, into a new array.
  let existingInvoicesNotInIncomingInvoicesArray = []
  existingInvoices.forEach(function(number) {
    if (!incomingInvoiceNumbers.some(existing => existing === number)) {
      existingInvoicesNotInIncomingInvoicesArray.push(number)
    }
  })
  // Creates array of arrays of rows to remove. Format [[bucketName, row]]. 
  let rowsToRemove = []
  let removalCount = 0
  for (let i = 0; i < existingInvoicesNotInIncomingInvoicesArray.length; i++) {
    // Looks for the row with the invoice, Sheet by Sheet.
    const bucket1 = ss.getSheetByName('30-40').getRange(2, 1, ss.getSheetByName('30-40').getLastRow()).getValues()
    for (let j = 0; j < bucket1.length; j++) {
      if (bucket1[j][0] === existingInvoicesNotInIncomingInvoicesArray[i]) {
        rowsToRemove.push(['bucket 1', j + 2])
        removalCount++
        break
      }
    }
    const bucket2 = ss.getSheetByName('41-60').getRange(2, 1, ss.getSheetByName('41-60').getLastRow()).getValues()
    for (let j = 0; j < bucket2.length; j++) {
      if (bucket2[j][0] === existingInvoicesNotInIncomingInvoicesArray[i]) {
        rowsToRemove.push(['bucket 2', j + 2])
        removalCount++
        break
      }
    }
    const bucket3 = ss.getSheetByName('61-100').getRange(2, 1, ss.getSheetByName('61-100').getLastRow()).getValues()
    for (let j = 0; j < bucket3.length; j++) {
      if (bucket3[j][0] === existingInvoicesNotInIncomingInvoicesArray[i]) {
        rowsToRemove.push(['bucket 3', j + 2])
        removalCount++
        break
      }
    }
    const bucket4 = ss.getSheetByName('100+').getRange(2, 1, ss.getSheetByName('100+').getLastRow()).getValues()
    for (let j = 0; j < bucket4.length; j++) {
      if (bucket4[j][0] === existingInvoicesNotInIncomingInvoicesArray[i]) {
        rowsToRemove.push(['bucket 4', j + 2])
        removalCount++
        break
      }
    }
  }
  // Sorts rows to remove from largest row to smallest so that deletion is done properly, then deletes each row.
  rowsToRemove.sort((a, b) => b[1] - a[1])
  rowsToRemove.forEach(row => removeRow(row))
  SpreadsheetApp.getActiveSpreadsheet().toast(`✔️️ Removed ${removalCount} paid invoices from AR station.`)
}

                      
/** Looks for matching name and firm name in Contacts Sheet. */
function matchEmails() {
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`🟢️️ Started finding matching email addresses for invoice POCs.`)
  
  try {
    const ss = SpreadsheetApp.getActive()
    const dataTab = ss.getSheetByName('Data Drop')
    const bucket1 = ss.getSheetByName('30-40')
    const bucket2 = ss.getSheetByName('41-60')
    const bucket3 = ss.getSheetByName('61-100')
    const bucket4 = ss.getSheetByName('100+')
    const array = [bucket1, bucket2, bucket3, bucket4]
    const contacts = ss.getSheetByName('Contacts')
    const contactsData = contacts.getRange(2, 1, contacts.getLastRow(), contacts.getLastColumn()).getValues()
    array.forEach(function(bucket) {
      // Tries to find matching POC, adds email address if successful.
      const invoices = bucket.getRange(2, 1, bucket.getLastRow(), bucket.getLastColumn()).getValues()
      for (var i = 0; i < invoices.length; i++) {
        let invoice = invoices[i]
        const POCEmail = invoice[7]
        if (POCEmail === '') {
          const firstName = invoice[5]
          const lastName = invoice[6]
          const firmName = invoice[4]
          let match = contactsData.filter(elem => elem[0] === firstName && elem[1] === lastName && elem[2] === firmName)
          if (match.length) {
            match = match[0]
            const matchingEmail = match[3]
            const row = i + 2
            bucket.getRange(row, 8).setValue(matchingEmail)
          }
        }
      }
    })
    SpreadsheetApp.getActiveSpreadsheet().toast(`✔️️ Finished matching emails for invoice POCs where possible.`)
  } catch (error) {
    addError(error)
    SpreadsheetApp.getActiveSpreadsheet().toast('⚠️️️ There was an error finding matching emails. This error has been recorded.')
  }
}


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
  try {
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
  } catch (error) {
    addError(error)
  }
}

/** Determines which tab an invoice should go to based on its age.
* @param {age} number How old (by the number of days) an invoice is.
* @return {obj} Reference to the correct Sheet.
*/
function findBucket(age) {
  try {
    const ss = SpreadsheetApp.getActive()
    const bucket1 = ss.getSheetByName('30-40')
    const bucket2 = ss.getSheetByName('41-60')
    const bucket3 = ss.getSheetByName('61-100')
    const bucket4 = ss.getSheetByName('100+')
    if (age > 30 && age <= 40) {
      return bucket1
    } else if (age > 40 && age <= 60) {
      return bucket2 
    } else if (age > 60 && age <= 100) {
      return bucket3
    } else {
      return bucket4
    }
  } catch (error) {
    addError(error)
  }
}

/** Removes an invoice (by deleting the row) 
* @param {element} array 2d array with bucket name, then row number.
*/
function removeRow(element) {
  
  const bucket = element[0]
  const row = element[1]

  const ss = SpreadsheetApp.getActive()
  const bucket1 = ss.getSheetByName('30-40')
  const bucket2 = ss.getSheetByName('41-60')
  const bucket3 = ss.getSheetByName('61-100')
  const bucket4 = ss.getSheetByName('100+')
  
  try {
    switch(bucket) {
    case 'bucket 1':
      bucket1.deleteRow(row);
      break;
    case 'bucket 2':
      bucket2.deleteRow(row);
      break;
    case 'bucket 3':
      bucket3.deleteRow(row);
      break;
    case 'bucket 4':
      bucket4.deleteRow(row);
      break;
    default:
      addError('Got to default case in removeRow function.')
    }
  } catch (error) {
    addError(`Error in removeRow: ${error}`)
  }
}

/** Returns an array of existing invoice numbers. */
function getExistingInvoiceNumbers() {
  const ss = SpreadsheetApp.getActive()
  const bucket1 = ss.getSheetByName('30-40')
  const bucket2 = ss.getSheetByName('41-60')
  const bucket3 = ss.getSheetByName('61-100')
  const bucket4 = ss.getSheetByName('100+')
  const existingInvoicesInBucket1 = bucket1.getRange(2, 1, bucket1.getLastRow()).getValues()
  const existingInvoicesInBucket2 = bucket2.getRange(2, 1, bucket2.getLastRow()).getValues()
  const existingInvoicesInBucket3 = bucket3.getRange(2, 1, bucket3.getLastRow()).getValues()
  const existingInvoicesInBucket4 = bucket4.getRange(2, 1, bucket4.getLastRow()).getValues()
  const existingInvoices = existingInvoicesInBucket1.concat(existingInvoicesInBucket2, existingInvoicesInBucket3, existingInvoicesInBucket4).filter(invoice => invoice[0] !== '') 
  return existingInvoices
}

/** Toggles the email send status value to 'Sent' for all invoices associated with a POC after successfully sending an email.
* @param {matches} array Array of bucket name (string) and row number matches for a POC.
*/
function toggleEmailSendStatus(matches) {
  const ss = SpreadsheetApp.getActive()
  const bucket1 = ss.getSheetByName('30-40')
  const bucket2 = ss.getSheetByName('41-60')
  const bucket3 = ss.getSheetByName('61-100')
  const bucket4 = ss.getSheetByName('100+')
  
  matches.forEach(function(match) {
    try {
       const bucket = match[0]
    const row = match[1]
    switch(bucket) {
      case 'bucket 1':
        bucket1.getRange(row, 10).setValue('Sent')
        break;
      case 'bucket 2':
        bucket2.getRange(row, 10).setValue('Sent')
        break;
      case 'bucket 3':
        bucket3.getRange(row, 10).setValue('Sent')
        break;
      case 'bucket 4':
        bucket4.getRange(row, 10).setValue('Sent')
        break;
      }
    } catch (error) {
      console.log(error)
      addError(error)
    }
  })
}