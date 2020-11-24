// Googleフォームのファイルアップロードデータを抽出し、ファイル名を変更して特定フォルダに一括保存する
// ＊ただし、ファイルアップロードは「１ファイルのみ」であり、複数ファイルには対応していない。
//
// 実際の処理：Googleスプレットシートの特定列（INPUT_SpreadSheet_num）に入力されたGoogleドライブ内のファイルへのリンクを抽出し、ファイル名を変更して特定フォルダ（OUTPUT_FOLDER_ID）に一括保存する Google Apps Script
//
function getAttachedDocuments() {
// == 値変更が必要な環境設定（ここから） ==
  var OUTPUT_FOLDER_ID = '=== FOLDER ID ==='; 
  var INPUT_SpreadSheet_ID ='=== Google SpreadSheet ID ===';
  var INPUT_SpreadSheet_num = 5; // Item of Attached File Link (start is 0)
  var OUTOUT_FOLDER_name_num = 2; // Item of Name (start is 0)
// == 値変更が必要な環境設定（ここまで） ==

// 以下、プログラム処理
  var saveFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID); //フォルダを取得
  var sheet = SpreadsheetApp.openById(INPUT_SpreadSheet_ID).getSheets()[0]; // 1つ目のシートをとる。
  
  // 全セルを保存
  var data = sheet.getDataRange().getValues();
  
  if(data){
    for(var num_row in data){
     //タイムスタンプ（申請日時）を取得
     var timestamp = new Date(data[num_row][0]);

     if(!isNaN(timestamp.getTime())){
       //名前の取得
       var name = data[num_row][OUTOUT_FOLDER_name_num];
       //添付ファイルのURL
       var urls = data[num_row][INPUT_SpreadSheet_num];
       
       if(urls){
       	var urls_data = urls.split(',');
       	var urls_id = [];
       	for(var i in urls_data){
       		urls_id[i] =  urls_data[i].split('=')[1].trim();
       	}
       }
      
       // タイムスタンプから年月日を取り出す関数
       var toDoubleDigits = function(num) {
         num += "";
         if (num.length === 1) {
           num = "0" + num;
         }
         return num;     
       };
   
       var yyyy = timestamp.getFullYear();
       var mm   = toDoubleDigits(timestamp.getMonth()+1);
       var dd   = toDoubleDigits(timestamp.getDate());
       
       var date_str = yyyy+"-"+mm+"-"+dd;
       
       // OUTPUT_FOLDER_ID で指定したフォルダへ、年-月-日_名前_添付ファイル名 で保存する。
       if(urls_id){
       	 for(var i in urls_id){
       	 	if(urls_id[i]){
       	 		var attached_file = DriveApp.getFileById(urls_id[i]);
       	 		var attached_file_name = attached_file.getName();
       	 		if (attached_file){
       	 			var attached_f= date_str +"_"+name+"_"+attached_file_name;
       	 			attached_file.makeCopy(attached_f, saveFolder);          
       	 		}
       	 	}
         } 
       }
     } // END - if(!isNaN(timestamp.getTime()))
    } // END -  for(var num_row in data){
  } // END -  if (data)
}