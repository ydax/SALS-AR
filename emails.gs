/** 
function manuallyUpdateCalendar(e) {
  
  try {
    // Sets a mutual-exclusion lock to prevent code collisions if user is making multiple quick edits.
    const lock = LockService.getDocumentLock();
    lock.waitLock(6000);
    
    var SACal = CalendarApp.getCalendarById('salegalsolutions.com_17vfv1akbq03ro6jvtsre0rv84@group.calendar.google.com');
    var ss = SpreadsheetApp.getActive();
    var depoSheet = ss.getSheetByName('Schedule a depo');
    
    // Check to see if the edit was made on the "Schedule a depo" sheet
    var sheetName = e.source.getSheetName();
    if (sheetName === 'Schedule a depo') {
      
      // If yes, get information about the edit made
      var editRow = e.range.getRow();
      var editColumn = e.range.getColumn();
      
      /////////////////////////////////////////////////
      // ROUTING BASED ON THE COLUMN THAT WAS EDITED //
      /////////////////////////////////////////////////
      
      switch(editColumn) {
          // Routing if it was made to Status Column.
        case (1):
          if (depoSheet.getRange(editRow, 1).getValue() === '🔴 Cancelled') {
            var eventId = depoSheet.getRange(editRow, 37).getValue();
            cancelDepo(eventId, editRow);
          };
          break;
          
          // Routing if it was made to event date. 2 because Date is in Column B.
        case (2):
          editDepoDate(e, ss, SACal, depoSheet, editColumn, editRow);
          updateSheetsOnTimeOrDateEdit(editRow);
          break;
          
          // Routing the edit if it was made to the event time. 7 because Start Time is in Column G.
        case (7):
          editDepoTime(e, ss, SACal, depoSheet, editColumn, editRow);
          break;
          
          // Routing if the edit is made to Columns recorded in Calendar events.
        case (3):
        case (4):
        case (6):
        case (8):
        case (9):
        case (10):
        case (11):
        case (12):
        case (13):
        case (14):
        case (17):
        case (18):
        case (19):
        case (20):
        case (21):
        case (22):
        case (24):
        case (25):
        case (26):
        case (27):
          editDepoGeneral(e, ss, SACal, depoSheet, editColumn, editRow);
          break;
          
          // NEXT: services information and the Services Calendar
          
        default:
          Logger.log('There are no more cases currently supported.');
      };
    };
    
    // Releases the mutual exclusion lock.
    lock.releaseLock();
  } catch (error) {
    addToDevLog('Error in manuallyUpdateCalendar: ' + error);
  }
};

*/