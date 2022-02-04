# Google-Apps-Script
用以將google sheet的內容定期更新到依時間命名的新sheet內



語法 :

```
function myFunction() {


  　// 1. 指定URL開啟試算表(請輸入Google Spreadsheet文件的實際URL--複製到"#"前)
    var url = "https://docs.google.com/spreadsheets/d/1fPMuSFJI_2sgILG-w2T51teYbC5oHa3tzEjKWV9I0Ds/edit";   
    var spredSheet = SpreadsheetApp.openByUrl(url);
 
    // 2. 取得日期 (以取得時間命名)
    var d = new Date();
    var targetName = Utilities.formatDate(d, "GMT+8", "yyyy-MM-dd HH:mm:ss");

    // 3. 開啟工作表和目標工作表
    var baseSheet = spredSheet.getSheetByName("工作表1");
    var targetSheet = createSheet(spredSheet, targetName, 0);

    // 4. 複製取得值至目標工作表
    // 指定複製範圍 (注意不要複製到有程式碼的地方!!)
    var rangeToCopy = baseSheet.getRange('A3:D102');    
    
   //指定複製目的地的起始儲存格
    var targetToCopy = targetSheet.getRange('A2:D101');
    rangeToCopy.copyTo(targetToCopy);
    targetSheet.getRange(1,1).setValue("Title");
    targetSheet.getRange(1,2).setValue("Url");
    targetSheet.getRange(1,3).setValue("Date Created");
    targetSheet.getRange(1,4).setValue("Summary");
    rangeToCopy.copyTo(targetToCopy);

    // 5. 再次使用ImportFeed函數
    baseSheet.getRange(1,1).setValue("執行時間");
    baseSheet.getRange(1,2).setValue(targetName); 
    baseSheet.getRange(2,1).setValue("");
    baseSheet.getRange(2,1).setFormula("IMPORTfeed(\"https://rss.applemarketingtools.com/api/v2/us/music/most-played/10/albums.rss\",\"items\",true,100)");  //反斜線為跳字元
    
}

function createSheet(spredSheet, sheetName, index) {
    var sheet = spredSheet.getSheetByName(sheetName);
    if ( sheet == null) {
        spredSheet.insertSheet(sheetName, index);
        sheet = spredSheet.getSheetByName(sheetName);
    }
    return sheet;
}

```
