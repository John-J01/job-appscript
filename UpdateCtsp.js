
// Last update 16/03/2022

function updateMSP() {

    var ui = SpreadsheetApp.getUi();
    //var response = ui.alert(
    // "Vui lòng xác nhận !!!",
     //"Bạn có chắc copy dòng phiếu tư vấn này sang thông tin sản phẩm hay không? ", 
     //ui.ButtonSet.YES_NO);
     //if (response == ui.Button.YES) {
      var tvSp = SpreadsheetApp.getActive().getSheetByName("Phiếu TVSP");
        var maSp = SpreadsheetApp.getActive().getSheetByName("Phiếu TVSP").getRange("N2").getValue();
      var ttSp = SpreadsheetApp.getActive().getSheetByName("Thông tin sản phẩm");
      var dataTvsp = ttSp.getRange("B1:B").getValues();
      var lastRow = dataTvsp.filter(String).length;
           for(var i = 1 ;i <=  lastRow ; i++){
                if (ttSp.getRange(i,2,i,2).getValue()== maSp) {
                    columnB = i;
                    var notification = "Đã update mã "+maSp;
                }else{
                    var columnB = lastRow+1;
                    var notification = "Đã tạo dòng mới "+maSp;
                } 
            }
      tvSp.getRange("N2:DG2").copyTo(ttSp.getRange(columnB,2),SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      ui.alert(
        'Hoàn Thành',
         notification,
         ui.ButtonSet.OK);
     }
     //if (response == ui.Button.NO) {
      //"Chương trình không thực hiện lệnh"
    // }
     //} 
    


    function rePlace() {
      var sh1 = SpreadsheetApp.getActive().getSheetByName("Thông tin sản phẩm");
      var sh2 = SpreadsheetApp.getActive().getSheetByName("Mực in");
      var sh3 =  SpreadsheetApp.getActive().getSheetByName("Bản in");
      var sh4 =  SpreadsheetApp.getActive().getSheetByName("File thiết kế");
      var lastRow = sh1.getRange("B1:B").getValues().filter(String).length;
      //Logger.log(lastRow);
      for(var i = 2 ;i <=  lastRow ; i++){
    
    //Logger.log(sh1.getRange(i,100,i,100).getValue());
    
        if(sh1.getRange(i,102,i,102).getValue() !=""){
          //Logger.log(SpreadsheetApp.getActive().getSheetByName("File thiết kế").getRange(i,2,i,2).getValue());
          oldValue = sh1.getRange(i,2,i,2).getValue();
          newValue = sh1.getRange(i,102,i,102).getValue();    
           sh2.getRange("B1:B").createTextFinder(oldValue).replaceAllWith(newValue);
           sh3.getRange("B1:B").createTextFinder(oldValue).replaceAllWith(newValue);
           sh4.getRange("B1:B").createTextFinder(oldValue).replaceAllWith(newValue);
           sh1.getRange("B1:B").createTextFinder(oldValue).replaceAllWith(newValue);
         }  
         }    
    };
    
    
    
    
    