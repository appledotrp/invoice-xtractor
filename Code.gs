function onOpen(e) {
  addMenu();
}

function addMenu(){
  try{
    // Add Sync menu
    var syncMenu = SpreadsheetApp.getUi().createMenu('Invoice Parser Tool');
    syncMenu.addItem('Download Attachment', 'downloadAttachments');
    syncMenu.addItem('Parse Now', 'main');
    syncMenu.addItem('Clear', 'clearColumns');
    syncMenu.addItem('Force Re-run', 'forceRun');
    
    // Add Guide button
    syncMenu.addSeparator();
    syncMenu.addItem('Guide', 'howItWorks');

    syncMenu.addToUi();
  } catch(err){
    Logger.log(err);
  }
}

function howItWorks(){
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Guide')
    .setWidth(800)
    .setHeight(600)
    .setTitle('How It Works');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

function showCustomMessageAlert(message, imageUrl) {
  try {
    var imageStyle = 'max-width: 100%; max-height: 100%;';
    var html = HtmlService.createHtmlOutput('<img src="' + imageUrl + '" style="' + imageStyle + '"></img>');
    html.setHeight(200);
    html.setWidth(300);
    html.setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    SpreadsheetApp.getUi().showModelessDialog(html, message);
  }
  catch(err){
    Logger.log('Console Error');
  }  
}

function convertPDFToTextInDrive(driveFileId) {  
  var docFile = Drive.Files.insert(
    {
      title: DriveApp.getFileById(driveFileId).getName(),
      mimeType: 'application/pdf'},
    DriveApp.getFileById(driveFileId).getBlob(),
    { convert: true }
  );
  var text = DocumentApp.openById(docFile.id).getBody().getText();    
  Drive.Files.remove(docFile.id);
  return text;
}

function main() {  
  var spreadsheetId = "1L3mmua7RLMX6dQde1VifN_hJ1BDUBm1Am7RA7_LFpME";
  var sheetName = "Main";
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var brandInp = sheet.getRange("A2").getValue();

  var driveFolderId = "1oX1Ld-zJlPBGR4bKe_mHt2jRvgJjPe1D";
  var folder = DriveApp.getFolderById(driveFolderId);
  var files = folder.getFilesByType(MimeType.PDF);

  var invoiceDataArray = [];

  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() === "application/pdf") {
      try {
        var text = convertPDFToTextInDrive(file.getId());        

        if (brandInp === 'Altra') {
          var invoiceData = altraParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'ALPS') {
          var invoiceData = alpsParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'Helly Hansen') {
          var invoiceData = hellyHansenParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'Thorogood') {
          var invoiceData = thorogoodParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'Footmates') {
          var invoiceData = footmatesParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'Smoky Mountain') {
          var invoiceData = smokyMountainParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'Rocky') {
          var invoiceData = rockyParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'Dansko') {
          var invoiceData = danskoParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'Hypard') {
          var invoiceData = hypardParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'Nike Swim') {
          var invoiceData = nikeSwimParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'Weatherbeeta') {
          var invoiceData = weatherbeetaParser(text);
          invoiceDataArray.push(invoiceData);
        } else if (brandInp === 'Arborwear') {
          var invoiceData = arborwearParser(text);
          invoiceDataArray.push(invoiceData);
        }
        file.setTrashed(true);
      } catch (error) {
        if(error === "Exceeded maximum execution time"){
          Logger.log("Runtime Error: " + error)
        } else {
          Logger.log("Error Encountered: " + error)
        }         
      }
      try{
        displayInformation(invoiceDataArray);  
      } catch(err){
        Logger.log("Error Encountered for Display on this file: " + file.getName());
      }      
    }
  }  
  showCustomMessageAlert('Finished', 'https://uxpro.cc/media/publicationimage/header_02081e6d6f.png');
}

function displayInformation(invoiceDataArray){
    var spreadsheetId = getSheetID();
    var sheetName = "Main";
    var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    var c = invoiceDataArray.length-1;
    var lastRow = sheet.getLastRow();
    var columnFValues = sheet.getRange("F1:F" + lastRow).getValues();
            for (var m = lastRow; m > 0; m--) {
              if (columnFValues[m - 1][0] !== "") {
                lastRow = m;
                break;
              }
            }
    var lastRowEmpty = lastRow + 1;          
    
    var range1 = sheet.getRange(lastRowEmpty, 6);
    var range2 = sheet.getRange(lastRowEmpty, 7);
    var range3 = sheet.getRange(lastRowEmpty, 8);
    var range4 = sheet.getRange(lastRowEmpty, 9);
    var range5 = sheet.getRange(lastRowEmpty, 10);
    var range6 = sheet.getRange(lastRowEmpty, 11);
    
    range1.setValue(invoiceDataArray[c].invoiceNumber);
    range2.setValue(invoiceDataArray[c].invoiceDate);
    range3.setValue(invoiceDataArray[c].orderNumber);
    range4.setValue(invoiceDataArray[c].totalDues);
    range5.setValue(invoiceDataArray[c].payTerms);
    range6.setValue(invoiceDataArray[c].pageCheck);
  }

function getSheet(){
  var spreadsheetId = getSheetID();
  var sheetName = "Main";
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  return sheet;
}

function convertDateFrom(dateString) {
  var date = new Date(dateString);
  date.setDate(date.getDate());
  var year = date.getFullYear();
  var month = date.getMonth() + 1; // Months are zero-based, so we add 1
  var day = date.getDate();

  return year + '/' + month + '/' + day;
}

function convertDateTo(dateString) {
  var date = new Date(dateString);
  date.setDate(date.getDate() + 1);
  var year = date.getFullYear();
  var month = date.getMonth() + 1; // Months are zero-based, so we add 1
  var day = date.getDate();

  return year + '/' + month + '/' + day;
}

function clearColumns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange('F2:K');
  range.clearContent();
}

function markMessageAsRead(messageId) {
  var message = GmailApp.getMessageById(messageId);
  
  if (message) {
    message.markRead();
  }
}

function altraParser(text) {
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;
  var totalDues = null;
  var payTerms = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();    
    Logger.log(line);
    if (line.startsWith('invoice #')) {
      invoiceNumber = lines[i + 1].trim();
    } else if (line === 'purchase') {
      orderNumber = lines[i + 2].trim();
    } else if (line.includes('subtotal ')) {
      var pattern = /total due\s+([\d.,]+)/i;
      var match = line.match(pattern);
      if (match) {
        totalDues = match[1].replace(/,/g, ''); // Remove commas from the matched value
      }
    } else if (line === 'invoice') {      
      var invoiceDateLineIndex = i + 2;
      if (invoiceDateLineIndex < lines.length) {
        var invoiceDateLine = lines[invoiceDateLineIndex];
        if (invoiceDateLine.includes('/')) {
          invoiceDate = invoiceDateLine.trim();          
        }
      }
    } else if(line === 'terms') {
      payTerms = lines[i + 1].trim();
    }
  }
    Logger.log('Altra Selected');
    return {
      invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms
    };
}

function alpsParser(text){
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;
  var totalDues = null;
  var payTerms = null;

  var lines = text.split('\n');
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim().toLowerCase();              
      if (line.startsWith('invoice number')) {
        invoiceNumber = lines[i + 5].trim();
      } else if (line.includes('customer p.o.')) {
        orderNumber = line.replace('customer p.o. ', ""); 
      } else if(line.includes('order number: order date')) {       
        invoiceDate = lines[i + 5].trim();
      } else if (line.startsWith('invoice total') && lines[i].trim().split(" ")[2] !== 'USD') {
        totalDues = lines[i].trim().split(" ")[2];
      } else if (line.startsWith('terms net')) {
        payTerms = line.replace('terms net', "Net");
      } 
    }
    
    Logger.log('Alps Selected');
    return {
      invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms
    };
}

function hellyHansenParser(text){  
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;
  var totalDues = null;
  var payTerms = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();         
    if (line.startsWith('invoice no')) {
        invoiceNumber = lines[i].trim().split(" ").slice(2).join(" ");
    } else if (line === 'date') {
        invoiceDate = lines[i + 1].trim();
    } else if (line.includes('mike/ mary mayo')) {        
        orderNumber = line.replace('mike/ mary mayo ', "");
    } else if (line.includes('mike/mary mayo')) {        
        orderNumber = line.replace('mike/mary mayo ', "");
    } else if (line.includes('invoice total')) {
        totalDues = lines[i].trim().split(" ").slice(2).join(" ");
    } else if (line.includes('payment terms')) {
        payTerms = lines[i + 3].trim();
    }
  }    
    Logger.log('Helly Hansen Selected');
    return {
      invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms
    };
}

function thorogoodParser(text) {
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;
  var tempDue = null;
  var totalDues = null;
  var payTerms = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();        
    if (line === 'customer no order no') {      
      invoiceNumber = lines[i + 1].trim();
      invoiceDate = lines[i + 2].trim();
    } else if(line.includes("your order no")){
      orderNumber = lines[i + 5].trim();
    } else if (line.startsWith('to pay usd')) {        
        tempTotal =  lines[lines.length - 1].trim();          
        tempTotal.includes(' ') ? tempDue =  tempTotal.trim().split(" ").slice(1).join(" ") : tempDue =  tempTotal;
        tempDue.includes('.') ? totalDues = tempDue : totalDues = lines[i + 2].trim() || lines[i + 1].trim().split(" ")[1];
    } else if (line.includes('days net')) {        
        payTerms = lines[i].trim();
    }      
  }
  Logger.log('Thorogood Selected');
  return {
    invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms
  };
}

function footmatesParser(text) {
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;  
  var totalDues = null;
  var payTerms = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();
    Logger.log(line);
    if (line === 'invoice number:') {      
      invoiceNumber = lines[i + 3].trim();      
    } else if(line === 'invoice date:'){
      invoiceDate = lines[i + 3].trim();
    } else if(line === "customer p/o:"){
      orderNumber = lines[i + 6].trim();
    } else if (line === ('invoice total:') && lines[i+1].includes('$', ' ')) { 
      totalDues = lines[i + 1];
    } else if (line === 'terms:' && !lines[i+1]) {        
        payTerms = lines[i + 2].trim();
    }      
  }
  Logger.log('Footmates Selected');
  return {
    invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms
  };
}

function smokyMountainParser(text) {
  var invoiceNumber = null;  
  var invoiceDate = null;
  var orderNumber = null;  
  var totalDues = null;
  var payTerms = null;
  var pageCheck = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();
    Logger.log(line);            
    if(line.includes('page')){
      var tempInv1 = lines[i].trim().split(" ")
      var tempInv2 = tempInv1.slice(5, 8).join(" ");
      invoiceNumber = tempInv2.split(" ")[0];
      //pageCheck = tempInv2.slice(1, 2);
      pageCheck = tempInv2.split(" ").slice(1, 3).join(" ");

    } else if (line === 'invoiced date') {
      invoiceDate = lines[i + 1].trim();
    } else if (line === "customer po") {
      orderNumber = lines[i + 1].trim();
    } else if (line.startsWith('amount due')) { 
      totalDues = lines[i].trim().split(" ")[2];
    } else if (line === 'fedex home delivery terms') {        
      payTerms = lines[i + 1].trim();
    } 
  }
  Logger.log('Smoky Selected' + pageCheck);
  return {
    invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms, pageCheck: pageCheck
  };
}

function rockyParser(text) {
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;  
  var totalDues = null;
  var payTerms = null;
  var pageCheck = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();    
    Logger.log(line);
    if (line.includes('invoice no.')){
      invoiceNumber = (lines[i].trim().split(" ")[2]) || lines[i+1].trim();      
    } else if(line.includes('invoice date')){
      invoiceDate = lines[i].trim().split(" ").slice(-1).join(" ");
    } else if(line.includes("customer po no.") && lines[i + 1].includes('Order No.')){      
      orderNumber = lines[i].trim().split(" ").slice(-1).join(" ");
    } else if (line.includes('payment received after amount due')){
      totalDues = lines[i].trim().split(" ")[5]
    } else if (line === 'terms' && lines[i + 1].includes('NET ')) {
        payTerms = lines[i + 1].trim();
    } else if(line.includes('credit memo')){
        pageCheck = lines[i].replace('CREDIT MEMO', '');
    } else if(line.includes('invoicepage')){
        pageCheck = lines[i].replace('INVOICE', '');      
    }
  }
  Logger.log('Rocky Selected');
  return {
    invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms, pageCheck: pageCheck
  };
}

function danskoParser(text) {
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;  
  var totalDues = null;
  var payTerms = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();    
    if (line === 'invoice no.'){
      invoiceNumber = lines[i + 5].trim();
    } else if(line.includes('invoice date')){
      invoiceDate = (lines[i + 1].trim().split(" ")[1]);
    } else if (line.includes('amount due')){
      totalDues = lines[i + 5].trim();
    } else if (line.includes(' net ') && line.includes("p.o. number")) {
      // Both conditions are satisfied
      payTerms = lines[i].trim().split(" ").slice(-2).join(" ");
      lines[i + 1].includes('SO-') ? orderNumber =  lines[i + 2].trim() : orderNumber =  lines[i + 1].trim();
    } else if (line.includes(' net ')) {
      // Only the first condition is satisfied
      payTerms = lines[i].trim().split(" ").slice(-2).join(" ");
    } else if (line.includes("p.o. number")) {
      // Only the second condition is satisfied
      lines[i + 1].includes('SO-') ? orderNumber =  lines[i + 2].trim() : orderNumber =  lines[i + 1].trim();
    }    
  }
  Logger.log('Dansko Selected');
  return {
    invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms    
  };
}

function hypardParser(text) {
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;  
  var totalDues = null;
  var payTerms = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();
    Logger.log(line);
    if (line.includes('invoice no.:')){
      invoiceNumber = lines[i].trim().split(" ")[0];
    } else if(line==='date:'){
      invoiceDate = lines[i + 1].trim();
    } else if(line.includes("customer p/o no.")){
      orderNumber = lines[i + 12].trim();
    } else if (line.includes('total amount:')){
      totalDues = lines[i].trim().split(" ")[4];
    } else if (line.includes('terms')) {
      payTerms = lines[i + 15].trim();
    }      
  }
  Logger.log('Hypard Selected');
  return {
    invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms    
  };

}
function nikeSwimParser(text) {
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;
  var totalDues = null;
  var payTerms = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();    
    if (line.includes('invoice no: ')) {
      invoiceNumber = lines[i].trim().split(" ")[2];
    } else if (line.includes('invoice date:') && line.includes("order no.:")) {
      var tempDate = lines[i].trim().split(" ").slice(-6).join(" ");
      invoiceDate = tempDate.trim().split(" ")[0];
      var tempPO = lines[i].trim().split(" ").slice(-3).join(" ");
      orderNumber = tempPO.trim().split(" ")[0];
    } else if (line.includes('invoice date:')) {
      var tempDate = lines[i].trim().split(" ").slice(-6).join(" ");
      invoiceDate = tempDate.trim().split(" ")[0];
    } else if (line.includes("cust po ref:")) {     
      orderNumber = lines[i].trim().split(" ").slice(-1).join(" ");      
    } else if (line.includes('total amount:')) {
      totalDues = lines[i].trim().replace('Total Amount: ', '');
    } else if (line.includes('pay terms:')) {
      payTerms = lines[i + 1].trim();
    }
  }
  Logger.log('Nike Swim Selected');
  return {
    invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms
  };
}

function weatherbeetaParser(text) {
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;
  var totalDues = null;
  var payTerms = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();

    if (line.includes('invoice no.:')) {
      invoiceNumber = lines[i + 1].trim();
    } else if (line === 'invoice date:') {
      invoiceDate = lines[i + 1].trim();
    } else if (line.includes('total invoice:')) {
      totalDues = line.replace('total invoice: ', '');
    } else if (line.startsWith('order no:') && line.includes('eom')) {
      var tempOrder = line.split(' ');
      orderNumber = tempOrder[2].slice(0, -2);
      var tempTerms = lines[i].split(' ');
      payTerms = tempTerms[2].substr(-2) + ' ' + tempTerms[3];
    } else if (line.startsWith('order no:')) {
      var tempOrder = lines[i].split(' ');
      orderNumber = tempOrder[2].slice(0, -2);
    } else if (line.includes('eom')) {
      var tempTerms = lines[i].split(' ');
      payTerms = tempTerms[2].substr(-2) + ' ' + tempTerms[3];
    }
  }
  Logger.log('Weatherbeeta Selected');

  return {
    invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms
  };
}

function arborwearParser(text) {
  var invoiceNumber = null;
  var invoiceDate = null;
  var orderNumber = null;
  var totalDues = null;
  var payTerms = null;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim().toLowerCase();    
    if (line === 'invoice') {
      invoiceNumber = lines[i + 3].trim();
    } else if (line === 'invoice date') {
      invoiceDate = lines[i + 3].trim();
    } else if (line.startsWith('total: ')) {
      totalDues = lines[i].replace('Total: ', '');
    } else if (line.includes('customer po:')) {
      orderNumber = lines[i+1].trim();      
    } else if (line === 'terms') {
      payTerms = lines[i + 2].trim();
    }
  }
  Logger.log('Arborwear Selected');

  return {
    invoiceNumber: invoiceNumber, invoiceDate: invoiceDate, orderNumber: orderNumber, totalDues: totalDues, payTerms: payTerms
  };
}

function brandSwitches(brandInp, text) {
  var invoiceDataArray = [];
  switch (brandInp) {
          case 'Altra':
            var invoiceData = altraParser(text);
            invoiceDataArray.push(invoiceData);            
            displayInformation(invoiceDataArray);
            break;

          case 'ALPS':
            var invoiceData = alpsParser(text);
            invoiceDataArray.push(invoiceData);            
            displayInformation(invoiceDataArray);
            break;

          case 'Helly Hansen':
            var invoiceData = hellyHansenParser(text);
            invoiceDataArray.push(invoiceData);            
            displayInformation(invoiceDataArray);
            break;

          case 'Thorogood':
            var invoiceData = thorogoodParser(text);
            invoiceDataArray.push(invoiceData);
            displayInformation(invoiceDataArray);
            break;

          case 'Footmates':
            var invoiceData = footmatesParser(text);
            invoiceDataArray.push(invoiceData);
            displayInformation(invoiceDataArray);            
            break;

          case 'Smoky Mountain':
            var invoiceData = smokyMountainParser(text);
            invoiceDataArray.push(invoiceData);
            displayInformation(invoiceDataArray);            
            break;

          case 'Rocky':
            var invoiceData = rockyParser(text);
            invoiceDataArray.push(invoiceData);
            displayInformation(invoiceDataArray);            
            break;

          case 'Dansko':
            var invoiceData = danskoParser(text);
            invoiceDataArray.push(invoiceData);
            displayInformation(invoiceDataArray);            
            break;

          case 'Hypard':
            var invoiceData = hypardParser(text);
            invoiceDataArray.push(invoiceData);
            displayInformation(invoiceDataArray);            
            break;
          
          case 'Nike Swim':
            var invoiceData = nikeSwimParser(text);
            invoiceDataArray.push(invoiceData);
            displayInformation(invoiceDataArray);            
            break;

          case 'Weatherbeeta':
            var invoiceData = weatherbeetaParser(text);
            invoiceDataArray.push(invoiceData);
            displayInformation(invoiceDataArray);            
            break;
            
          case 'Arborwear':
            var invoiceData = arborwearParser(text);
            invoiceDataArray.push(invoiceData);
            displayInformation(invoiceDataArray);            
            break;

            default:
            break;
  }
}

function getSheetID(){
  var spreadsheetId = "1L3mmua7RLMX6dQde1VifN_hJ1BDUBm1Am7RA7_LFpME"; 
  return spreadsheetId;
}

function downloadAttachments() {
  try {
    deleteAllFilesInDrive();

    var spreadsheetId = "1L3mmua7RLMX6dQde1VifN_hJ1BDUBm1Am7RA7_LFpME"; 
    var sheetName = "Main";
    var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);    

    var label = sheet.getRange("C2").getValue();
    var date1 = sheet.getRange("D2").getValue();
    var date2 = convertDateTo(date1);
    var afterDate = convertDateFrom(date1);
    Logger.log(afterDate);
    var beforeDate = date2;
    var driveFolderId = "1oX1Ld-zJlPBGR4bKe_mHt2jRvgJjPe1D"; // Replace with your Drive folder ID// ...

    threadString = ' label:' + label + ' after:' + afterDate + ' before:' + beforeDate;

    Logger.log(threadString);
    var threads = GmailApp.search(threadString);
    
    if (threads.length === 0) {
      throw new Error("No e-mail found.");
    }
    
    Logger.log(threads);    
    threads.forEach(function(thread) {
      var messages = thread.getMessages();

      messages.forEach(function(message) {
        var attachments = message.getAttachments();

        attachments.forEach(function(attachment) {
          var file = DriveApp.getFolderById(driveFolderId).createFile(attachment);
          Logger.log("Saved attachment: " + file.getName());        
        });
        var receivedTime = message.getDate();
        Logger.log("Received time: " + receivedTime);

        //message.markRead();
      });
    });
    var brandInp = sheet.getRange("A2").getValue();
    if (brandInp === 'Smoky Mountain' || brandInp === 'Rocky' || brandInp === 'Arborwear') { // Brands which PDF File needs to be splitted
      splitPDFDocuments();
    }
  } catch (error) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("An error occurred: " + error + "\nPlease check your filter. Note: Emails received earlier than 8:00 a.m. are considered yesterday's mails.", ui.ButtonSet.OK);
  }
}

function splitPDFDocuments() {
  
  var googleDriveFolderId = '1oX1Ld-zJlPBGR4bKe_mHt2jRvgJjPe1D';
  //var pdfCoAPIKey = 'aj.rapirap@outdoorequipped.com_ce84cc35958422a3f2530e8b6cc94803d6d05aadccbb4ca9e0b6f8a9b034eb40b8cc08f1';
  var pdfCoAPIKey = 'acordovan@outdoorequipped.com_7dfc69bede6271038176694aaea64b4c73cd4a50981b76bcea5c106e3f25d3f2bf647bc0'; // Needs to be updated when API Key Expires

  var folder = DriveApp.getFolderById(googleDriveFolderId);
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Prepare Payload
    var data = {
      "async": false,
      "encrypt": false,
      "inline": true,
      "name": "result",
      "url": file.getDownloadUrl(),
      "pages": "*"
    };

    var options = {
      'method': 'post',
      'muteHttpExceptions' : true,
      'contentType': 'application/json',
      'headers': {
        "x-api-key": pdfCoAPIKey
      },

      'payload': JSON.stringify(data)
    };

    var pdfCoResponse = UrlFetchApp.fetch('https://api.pdf.co/v1/pdf/split', options);

    var pdfCoRespContent = pdfCoResponse.getContentText();
    var pdfCoRespJson = JSON.parse(pdfCoRespContent);

    var resultUrls = pdfCoRespJson.urls;
  
    // Save Split PDFs in Google Drive Folder
    for (let i = 0; i < resultUrls.length; i++) {
      var splitFile = UrlFetchApp.fetch(resultUrls[i]).getBlob();
      folder.createFile(splitFile);
    }
    // Remove Original File
    file.setTrashed(true);
  }
}

function deleteAllFilesInDrive() {
  var folderId = '1oX1Ld-zJlPBGR4bKe_mHt2jRvgJjPe1D';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  while (files.hasNext()) {
    var file = files.next();
    Drive.Files.remove(file.getId());
  }
}

function forceRun(){
  deleteAllFilesInDrive();
  downloadAttachments();
}