function KatakanaToRomaji() {
  // 基本的な変換マップ
  var translateMap = {
    'あ':'a','い':'i','う':'u','え':'e','お':'o',
    'か':'ka','き':'ki','く':'ku','け':'ke','こ':'ko',
    'さ':'sa','し':'shi','す':'su','せ':'se','そ':'so',
    'た':'ta','ち':'chi','つ':'tsu','て':'te','と':'to',
    'な':'na','に':'ni','ぬ':'nu','ね':'ne','の':'no',
    'は':'ha','ひ':'hi','ふ':'hu','へ':'he','ほ':'ho',
    'ま':'ma','み':'mi','む':'mu','め':'me','も':'mo',
    'や':'ya','ゆ':'yu','よ':'yo',
    'ら':'ra','り':'ri','る':'ru','れ':'re','ろ':'ro',
    'わ':'wa','を':'wo','ん':'nn',
    'が':'ga','ぎ':'gi','ぐ':'gu','げ':'ge','ご':'go',
    'ざ':'za','じ':'zi','ず':'zu','ぜ':'ze','ぞ':'zo',
    'だ':'da','ぢ':'di','づ':'du','で':'de','ど':'do',
    'ば':'ba','び':'bi','ぶ':'bu','べ':'be','ぼ':'bo',
    'ぱ':'pa','ぴ':'pi','ぷ':'pu','ぺ':'pe','ぽ':'po',
  };
  var exceptionTranslateMap = {
    'ゃ':'ya','ゅ':'yu','ょ':'yo'
  }
  // アクティブなシートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // アクティブな範囲の配列を取得
  var rangeList = sheet.getActiveRangeList().getRanges();
  var result = '';
  var value;
  var tmp;
  var cell;
  var undefinedFlag = false;
  // 取得したアクティブな範囲をループで回す
  for(var i = 0; i < rangeList.length ; i++){
    // アクティブばセルをループで回す
    for(var j=rangeList[i].getRow();j<=rangeList[i].getLastRow();j++) {
      // セルのひらがなをローマ字に変換してセットする
      cell = sheet.getRange(j,rangeList[i].getColumn());
      value = cell.getValue();
      for(k=0;k<value.length;k++) {
        // 通常変換
        if(translateMap[value[k]] == undefined) {
          // 小文字対応
          if (exceptionTranslateMap[value[k]] == undefined) {
            undefinedFlag=true;
          } else {
            tmp = result.slice(0,result.length-1);
            result = tmp + exceptionTranslateMap[value[k]];
          }
        } else {
          result += translateMap[value[k]];
        }
        // 伸ばし棒の発音は次の母音を1文字削除
        if((result[result.length-1]=='a' && value[k+1]=='あ') ||
        (result[result.length-1]=='i' && value[k+1]=='い') ||
        (result[result.length-1]=='u' && value[k+1]=='う') ||
        (result[result.length-1]=='e' && value[k+1]=='え') ||
        (result[result.length-1]=='o' && value[k+1]=='う')
        ) 
        {
          k++;
        }
      }
      if(!undefinedFlag) {
        cell.setValue(result);
      }
      undefinedFlag=false;
      result = '';
    }
  }
}
