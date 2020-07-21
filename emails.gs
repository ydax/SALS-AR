/** Sends an email with details about all unpaid invoices associated with this POC in the AR station.
@param {Event} e Event object created when an edit is made.
@dev Called by onEdit(e) trigger
*/
function sendPOCEmail(e) {
  try {
    // Sets a mutual-exclusion lock to prevent code collisions if user is making multiple quick edits.
    const lock = LockService.getDocumentLock()
    lock.waitLock(6000)
    
    // Collects information about where the edit was made, returns if not in a relevant cell and value.
    const sheetName = e.source.getSheetName()
    const editRow = e.range.getRow()
    const editColumn = e.range.getColumn()
    if (!(sheetName == '30-40' || sheetName == '41-60' || sheetName == '61-100' || sheetName == '100+')) return;
    if (editColumn !== 10) return;
    
    const oldValue = e.oldValue
    const newValue = e.value
    if (newValue !== 'Sending') return;
    SpreadsheetApp.getActiveSpreadsheet().toast('📧️ Email automation started.')
    
    // Instantiates references to individual Sheets
    const ss = SpreadsheetApp.getActive()
    const dataTab = ss.getSheetByName('Data Drop')
    const bucket1 = ss.getSheetByName('30-40')
    const bucket2 = ss.getSheetByName('41-60')
    const bucket3 = ss.getSheetByName('61-100')
    const bucket4 = ss.getSheetByName('100+')
    
    // Gets POC details
    const activeSheet = ss.getSheetByName(sheetName)
    const rowData = activeSheet.getRange(editRow, 1, 1, activeSheet.getLastColumn()).getValues()[0]
    const POCEmail = rowData[7]
    if (POCEmail === '') {
      SpreadsheetApp.getActiveSpreadsheet().toast('⚠️ No email address for this POC. Automation cancelled.')
      activeSheet.getRange(editRow, editColumn).setValue(oldValue)
      return
    }
    const firstName = rowData[5]
    const lastName = rowData[6]
    const firmName = rowData[4]
    
    // Matches all unpaid invoices from this POC.
    const matches = getPOCInvoiceMatches(firstName, lastName, firmName) // Returns array of arrays with bucket reference string, row #
    
    // Gathers details from each match, then sends an email with these details.
    let collectiveInvoiceData = []
    matches.forEach(function(match) {
      let invoiceData
      const bucket = match[0]
      const row = match[1]
      switch(bucket) {
        case 'bucket 1':
          invoiceData = bucket1.getRange(row, 1, 1, bucket1.getLastColumn()).getValues()
          break;
        case 'bucket 2':
          invoiceData = bucket2.getRange(row, 1, 1, bucket1.getLastColumn()).getValues()
          break;
        case 'bucket 3':
          invoiceData = bucket3.getRange(row, 1, 1, bucket1.getLastColumn()).getValues()
          break;
        case 'bucket 4':
          invoiceData = bucket4.getRange(row, 1, 1, bucket1.getLastColumn()).getValues()
          break;
      }
      invoiceData = invoiceData[0]
      const invoiceNo = invoiceData[0]
      const amount = invoiceData[1]
      const daysOverdue = invoiceData[3]
      collectiveInvoiceData.push([invoiceNo, amount, daysOverdue])
    })
    
    console.log(collectiveInvoiceData) // works
    /** Looks like this: 
    [ [ 12755, 308.75, 26 ],
    [ 12780, 146.25, 19 ],
    [ 12796, 130, 17 ],
    [ 12803, 130, 13 ],
    [ 12755, 308.75, 27 ],
    [ 12780, 146.25, 20 ],
    [ 12796, 130, 18 ],
    [ 12803, 130, 14 ],
    [ 12668, 97.5, 60 ],
    [ 12755, 308.75, 26 ],
    [ 12780, 146.25, 19 ],
    [ 12796, 130, 17 ],
    [ 12803, 130, 13 ],
    [ 12668, 97.5, 60 ],
    [ 12668, 97.5, 60 ],
    [ 12629, 81.25, 80 ] ]
    */
    
    // Releases the mutual exclusion lock.
    lock.releaseLock()
  } catch (error) {
    addError('Error in sendPOCEmail: ' + error);
  }
}


///////////////////////////////////////////////////
/////////////////// UTILITIES /////////////////////
///////////////////////////////////////////////////

/** Returns an array of invoice numbers associated with the same POC.
* @param {firstName} string
* @param {lastName} string
* @param {firmName} string
*/
function getPOCInvoiceMatches(firstName, lastName, firmName) {
  SpreadsheetApp.getActiveSpreadsheet().toast('🔍️ Looking for all invoices associated with this POC.')
  const ss = SpreadsheetApp.getActive()

  // Gets invoice data
  const bucket1Data = ss.getSheetByName('30-40').getRange(2, 1, ss.getSheetByName('30-40').getLastRow(), ss.getSheetByName('30-40').getLastColumn()).getValues()
  const bucket2Data = ss.getSheetByName('41-60').getRange(2, 1, ss.getSheetByName('41-60').getLastRow(), ss.getSheetByName('41-60').getLastColumn()).getValues()
  const bucket3Data = ss.getSheetByName('61-100').getRange(2, 1, ss.getSheetByName('61-100').getLastRow(), ss.getSheetByName('61-100').getLastColumn()).getValues()
  const bucket4Data = ss.getSheetByName('100+').getRange(2, 1, ss.getSheetByName('100+').getLastRow(), ss.getSheetByName('100+').getLastColumn()).getValues()
  
  // Does a regex test of stringified data to look for preliminary match in each bucket, if found, loops over data, finds rows, and pushes into results array.
  let results = []
  const test1 = JSON.stringify(bucket1Data)
  if (test1.match(firstName) && test1.match(lastName) && test1.match(firmName)) {
    for (var i = 0; i < bucket1Data.length; i++) {
      if (firstName == bucket1Data[i][5] && lastName == bucket1Data[i][6] && firmName == bucket1Data[i][4]) {
        results.push(['bucket 1', i + 2])
      }
    }
  }
  const test2 = JSON.stringify(bucket2Data)
  if (test2.match(firstName) && test2.match(lastName) && test2.match(firmName)) {
    for (var i = 0; i < bucket2Data.length; i++) {
      if (firstName == bucket2Data[i][5] && lastName == bucket2Data[i][6] && firmName == bucket2Data[i][4]) {
        results.push(['bucket 2', i + 2])
      }
    }
  }
  const test3 = JSON.stringify(bucket3Data)
  if (test3.match(firstName) && test3.match(lastName) && test3.match(firmName)) {
    for (var i = 0; i < bucket3Data.length; i++) {
      if (firstName == bucket3Data[i][5] && lastName == bucket3Data[i][6] && firmName == bucket3Data[i][4]) {
        results.push(['bucket 3', i + 2])
      }
    }
  }
  const test4 = JSON.stringify(bucket4Data)
  if (test4.match(firstName) && test4.match(lastName) && test4.match(firmName)) {
    for (var i = 0; i < bucket4Data.length; i++) {
      if (firstName == bucket4Data[i][5] && lastName == bucket4Data[i][6] && firmName == bucket4Data[i][4]) {
        results.push(['bucket 4', i + 2])
      }
    }
  }
  
  return results
}