/** Goes into Automation Station (AS), collects previous orderers (POC full name, firm name, and email address),
* pushes each into an array, then stringify and stores an array of arrays of these previous POCs.
*/
function getPOCsFromAS() {
  SpreadsheetApp.getActiveSpreadsheet().toast(`✔️ Contacts pulled from Automation Station.`)
  try {
    // Goes into AS and collects data from 'Schedule a depo' Sheet
    const ASRef = 'https://docs.google.com/spreadsheets/d/1aEPtrPDMyGnIE1C49dLtmzj8hl18vVvOxzUkVP2dlak/'
    const AS = SpreadsheetApp.openByUrl(ASRef) 
    const deposSheet = AS.getSheetByName('Schedule a depo')
    const deposSheetData = deposSheet.getRange(2, 1, deposSheet.getLastRow(), deposSheet.getLastColumn()).getValues()
    
    // Instantiates an array, then pushes the POC full name, firm name, and email address as an array into the array.
    let array = []
    deposSheetData.forEach(function(elem) {
      const data = [elem[3], elem[7],elem[4]]
      addContact(data)
    })
    SpreadsheetApp.getActiveSpreadsheet().toast(`🟢️️ Started pulling contacts from Automation Station.`)
  } catch (error) {
    addError(error)
  }
}

/** Looks for contact in Contacts Sheet, adds if not found.
* @param {array of strings} [fullName, firmName, emailAddress]
*/
function addContact(array) {
  const ss = SpreadsheetApp.getActive()
  const contactsSheet = ss.getSheetByName('Contacts')
  
  // Instantiates required data.
  const existing = contactsSheet.getRange(2, 1, contactsSheet.getLastRow(), contactsSheet.getLastColumn()).getValues()
  const names = destructureName(array[0])
  const firstName = names[0]
  const lastName = names[1]
  const firmName = array[1]
  const email = array[2]
  
  // Looks for email in existing contacts, returns if found.
  const testString = JSON.stringify(existing)
  if (testString.match(email)) return;
  
  // Adds contact to Contacts Sheet.
  const row = contactsSheet.getLastRow() + 1
  contactsSheet.getRange(row, 1).setValue(firstName)
  contactsSheet.getRange(row, 2).setValue(lastName)
  contactsSheet.getRange(row, 3).setValue(firmName)
  contactsSheet.getRange(row, 4).setValue(email)
  SpreadsheetApp.getActiveSpreadsheet().toast(`🧑🏻 ${email} added to contacts.`)
}

/** Does a fresh pull of contacts from AS, then attempts to match to invoice POCs. 
* Note: Callable from Spreadsheet SA Legal Solutions Menu.
*/
function syncContacts() {
  getPOCsFromAS()
  matchEmails()
}