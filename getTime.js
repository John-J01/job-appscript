function getTime(){
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
  var targetBook = SpreadsheetApp.openById("1tFBc18O8cRlyPT2_14zt8Tfor1Jcsz7xJmbrXUUUZFw"); //target workbook
    var target = targetBook.getSheets()[0]; //Sheet1
  
    var readingGender = target.getRange("A3:A").getValues();
    var lastRow = readingGender.length +10;
    for(var i = 2000 ;i <=  lastRow ; i++){
      if (target.getRange(i,1,i,1).getValue() != "" && target.getRange(i,67,i,67).getValue() == ""){
      Logger.log(lastRow);
          target.getRange(i,67).setValue(Utilities.formatDate(new Date(),"GMT+7", "dd/MM/yyyy HH:mm")) ;
      }
    }
  }