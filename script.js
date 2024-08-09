function generateBarcode(ticketNumber) {
    var ticketNumberStr = String(ticketNumber);
    var url = "https://bwipjs-api.metafloor.com/?bcid=code128&text=" + encodeURIComponent(ticketNumberStr) + "&scale=3&height=10&includetext=true";
    return url;
  }
  
  function generateTicketCode() {
    var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
    var code = '';
    for (var i = 0; i < 5; i++) {
      code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return code;
  }
  
  function generateTickets() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.clear();
    sheet.appendRow(['Ticket Number', 'Barcode URL', 'Status']);
    
    var codes = new Set();
    for (var i = 1; i <= 520; i++) {
      var ticketCode;
      do {
        ticketCode = generateTicketCode();
      } while (codes.has(ticketCode));
      codes.add(ticketCode);
      var barcodeUrl = generateBarcode(ticketCode);
      sheet.appendRow([ticketCode, barcodeUrl, '']);
    }
  }
  
  function doGet(e) {
    var ticketNumber = e.parameter.ticketNumber;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tickets');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == ticketNumber) {
        if (data[i][2] == 'Used') {
          return ContentService.createTextOutput(JSON.stringify({ valid: false, message: 'Invalid ticket. This ticket has already been used.' })).setMimeType(ContentService.MimeType.JSON);
        } else {
          sheet.getRange(i + 1, 3).setValue('Used');
          return ContentService.createTextOutput(JSON.stringify({ valid: true, message: 'ACCESS GRANTED.' })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ valid: false, message: 'Invalid ticket.' })).setMimeType(ContentService.MimeType.JSON);
  }
  
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Barcode Menu')
      .addItem('Generate 520 Barcodes', 'generateTickets')
      .addToUi();
  }
  
  
  