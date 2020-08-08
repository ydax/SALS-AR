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
      const dueDate = invoiceData[2]
      const daysOverdue = invoiceData[3]
      const firmName = invoiceData[4]
      const firstName = invoiceData[5]
      const lastName = invoiceData[6]
      const POCEmail = invoiceData[7]
      collectiveInvoiceData.push([invoiceNo, amount, dueDate, daysOverdue, firmName, firstName, lastName, POCEmail])
    })
    
    // Gets Drive file ids of invoices associated with this POC.
    for (let i = 0; i < collectiveInvoiceData.length; i++) {
      const invoiceNo = collectiveInvoiceData[i][0]
      const id = findInvoice(invoiceNo)
      if (id) {
        collectiveInvoiceData[i].push(id)
      } 
      // If invoice wasn't found, send an email telling Blake it wasn't in the folder.
      else {
        // Lets user know via toast.
        SpreadsheetApp.getActiveSpreadsheet().toast(`⚠️ Could not find invoice no. ${invoiceNo}. Sent you a reminder email.`)
        // Sends reminder email.
        GmailApp.sendEmail('bboyd@salegalsolutions.com', `Could not find invoice no. ${invoiceNo}`, 'Howdy,\n\nYou were trying to send AR emails, and when AR station was looking for invoice number ' + invoiceNo + ' associated with ' + collectiveInvoiceData[i][5] + ' ' + collectiveInvoiceData[i][6] + ', it wasn\'t in the Invoices folder at https://drive.google.com/drive/folders/1fheoM0D86nabYbw0CYy9HGnoqs8d004M.\n\nYou may want to upload it to the Invoices folder. When that\'s done and when you send another email to this person, the invoice will then be included in the AR / collections email.', { name: 'AR Station Bot' })
        // Splices invoice out of collective invoices array and matches array.
        collectiveInvoiceData.splice(i, 1)
        matches.splice(i, 1)
      }
    }
    
    // Sends email to POC.
    sendCollectionEmail(collectiveInvoiceData)
    
    // Toggles email send status for invoices included in the email.
    toggleEmailSendStatus(matches)
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Toggled email send status of invoices associated with this POC in AR Station. Automation complete.')

    
    // Releases the mutual exclusion lock.
    lock.releaseLock()
  } catch (error) {
    addError('Error in sendPOCEmail: ' + error);
  }
}

/** Sends a collections email to POC.
@params {array} array Array of invoice data arrays associated with a POC.
*/
function sendCollectionEmail(array) {
  const templateId = getRandomTemplateId()
  const pocEmail = array[0][7]
  const firstName = array[0][5]
  try {
    //////// GENERATING MESSAGE FROM ID ////////////
    // Gets message from ID
    const id = Gmail.Users.Drafts.get('me', templateId).message.id
    const message = GmailApp.getMessageById(id)
    let template = message.getRawContent()
    let subject = GmailApp.getMessageById(id).getSubject()
    console.log(id, subject)
    
    // Structures plural / singular phrasing based on number of invoices.
    let invoiceSubject = 'an outstanding invoice'
    let invoicePlural = 'an invoice'
    if (array.length > 1) {
      invoiceSubject = 'a couple oustanding invoices'
      invoicePlural = 'a couple of invoices'
    }
    
    // Replaces template variables with custom ones for the user using RegExes.
    subject = subject.replace(/invoiceSubject/g, invoiceSubject)
    template = template.replace(/templates@salegalsolutions.com/g, pocEmail)
    template = template.replace(/invoiceSubject/g, invoiceSubject)
    template = template.replace(/invoicePlural/g, invoicePlural)
    
    // Creates an array of attachments.
    let attachments = []
    array.forEach(function(invoice) {
      const fileId = invoice[8]
      attachments.push(DriveApp.getFileById(fileId))
    })
    
    // Creates the new message
    GmailApp.sendEmail(
      pocEmail, 
      subject, 
      'Hello ' + firstName + ',\n\nI found ' + invoicePlural + ' that I wanted to bring to your attention (please see attached). Would you mind checking on this for me?\n\nThanks!\n\nBlake', 
      {
        attachments: attachments,
        // htmlBody: template,
        name: 'Blake Boyd',
        bcc: 'bboyd@salegalsolutions.com',
      })
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅️ Reminder email successfully sent to ${pocEmail}.`)
  } catch (error) {
    Logger.log(error)
    SpreadsheetApp.getActiveSpreadsheet().toast(`⚠️ There was an error sending this email. It has been logged for the developer.`)
    addError(error)
  }
}


/** Generates a deposition confirmation PDF to include in confirmation email
* @param {array} array Invoice data created inside sendPOCEmail.
* @return {pdfUrl} string URL (file hosted on Google Drive) where the confirmation PDF can be found.
* @dev 200805 This isn't being used currently. Instead, we're going into a Drive folder, finding the pre-existing PDFs, and attaching them.
*/
function createInvoicePDF (array) {
  
  const invoiceNo = array[0]
  const dueDate = array[1]
  const invoiceAmount = array[2]
  const daysOverdue = array[3]
  const firmName = array[4]
  const firstName = array[5]
  const lastName = array[6]
  const fullName = createFullName(firstName, lastName)
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`📝 Started Creating PDF for Invoice Number ${invoiceNo}`)
  
  try {
    // setup
    var template = DocumentApp.openByUrl('https://docs.google.com/document/d/1mZBA0evQ6qvDM03rvIM9FYWspHFE1L36MY7LT7RLBdo/edit')
    var templateId = '1mZBA0evQ6qvDM03rvIM9FYWspHFE1L36MY7LT7RLBdo'
    var automatedInvoiceFolder = '1Va-Yl6d3h-zKogDIaer55aNP_4RFB6kb'
    
    // Generates the Google Doc version of the confirmation PDF.
    var fileName = 'SA Legal Solutions Invoice No. ' + invoiceNo
    var folder = DriveApp.getFolderById(automatedInvoiceFolder)
    var generatedDocCertUrl = DriveApp.getFileById(templateId).makeCopy(fileName, folder).getUrl()
    
    // Generates the URL of the newly-generated Google Docs version of the confirmation PDF (without copying the template fresh).
    var newUrl = ''
    var files = DriveApp.getFilesByName(fileName)
    while (files.hasNext()) {
      var file = files.next()
      newUrl = file.getUrl()
    }
    SpreadsheetApp.getActiveSpreadsheet().toast('✔️ New Invoice Template Created 📝')
    
    // Adds deposition information to template.  
    var confirmationBody = DocumentApp.openByUrl(newUrl).getBody()
    
    confirmationBody.replaceText('invoiceDate', dueDate)
    confirmationBody.replaceText('invoiceNumber', invoiceNo)
    confirmationBody.replaceText('POCName', fullName)
    confirmationBody.replaceText('amountDue', invoiceAmount)
    
    DocumentApp.openByUrl(newUrl).saveAndClose()
    
    // Converts the Google Doc version to PDF and updates sharing settings.
    var pdfUrl = convertToPDF(newUrl).slice(0, -13)
    var pdfId = getIdFromUrl(pdfUrl)
    moveFile(pdfId, automatedInvoiceFolder)
    folder.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ Invoice No. ${invoiceNo} PDF Creation Successful`)
    
    // remove the doc version of the generated cert
    DriveApp.getFileById(getIdFromUrl(newUrl)).setTrashed(true)
    
    return pdfUrl
  } catch (error) {
    console.log(error)
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

/** Creates a full name from seperated name strings.
* @param {firstName} string
* @param {lastName} string
*/
function createFullName(firstName, lastName) {
  let fullName = firstName
  if (lastName !== '') {
    fullName = firstName + ' ' + lastName
  }
  return fullName
}

/** Finds the id of the invoice uploaded by Blake in Drive Invoices folder.
* @dev Folder link: https://drive.google.com/drive/folders/1fheoM0D86nabYbw0CYy9HGnoqs8d004M
*/
function findInvoice(invoiceNo) {
  try {
    // Looks for invoice in the Invoices folder.
    const invoice = DriveApp.getFolderById('1fheoM0D86nabYbw0CYy9HGnoqs8d004M').getFilesByName(`${invoiceNo}.pdf`)
    let id
    while (invoice.hasNext()) {
      const file = invoice.next()
      id = file.getId()
    }
    if (id) return id 
    
    // Looks for the invoice in the entire Drive folder if not yet found.
    const search = DriveApp.searchFiles(`title contains "${invoiceNo}"`)
    while (search.hasNext()) { 
      const file = search.next()
      const name = file.getName()
      // Make sure the file has the characters 'inv', the way Blake stores invoices originally.
      const match = name.match('Inv.')
      if (match) {
        return file.getId()
      }
    } 
  } catch (error) {
    console.log(error)
    addError(error)
  }
}

// Displays a list of template IDs in the logs
function seeTemplateIds () {
  // Gets Gmail objects for all messages in drafts
  var response = Gmail.Users.Drafts.list('me')
  var drafts = response.drafts
  
  for (var i = 0; i < drafts.length; i++) {
    var gmailObjId = drafts[i].id // gets the GMail object ID
    // Gets the message id of from the Gmail object
    let draft = Gmail.Users.Drafts.get('me', gmailObjId).message.id
    let recipient = GmailApp.getMessageById(draft).getTo()
    if (recipient === 'templates@salegalsolutions.com') {
      let subject = GmailApp.getMessageById(draft).getSubject()
      Logger.log(`${gmailObjId}: ${subject}`)
    }
  }
}

// Returns a random template ID.
function getRandomTemplateId() {
  const templateArray = ['r-8391933797462466790', 'r-1139857297992043373']
  const id = templateArray[Math.floor(Math.random() * templateArray.length)]
  return id
}
