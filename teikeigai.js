//100件以下で処理 100件以上はタイムアウトになる
function CreatePost_8Cases() {
    /// シートを取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("住所入力")
    const temSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("テンプレ")
    
    //データのある行からマイナス1のところを取得
    const aLastRow = sheet.getLastRow() -1 //何故かマイナス1

    //データがある行までの値を取得
    let amountDate = sheet.getRange(2,3,aLastRow,1).getValues()

    //情報量を取得
    let amountLength = amountDate.length

    //情報量を8で割って1プラスする(テンプレが8件ずつ入力してpdfを作る為,割った余り分まで処理する為にプラス1)
    let amountNum = amountLength / 8 + 1

  // 読み取り範囲（表の始まり行と終わり列）
  const topRow = 2
  const lastCol = 5
  const statusCellCol = 1

  // 予定の一覧バッファ内の列(0始まり)
  const statusNum = 0
  const originPostCode = 1
  const addressName_1 = 2 
  const addressName_2 = 3
  const firstAndLastName = 4      

  // 住所が20文字以上の場合の割り振り
  //C2～D2のデータがある行までの値を取得
  let colData = sheet.getRange(2,3,aLastRow,2).getValues();
  //C2のデータがある行までの一次元配列を取得
  let colData_2 = sheet.getRange(2,3,aLastRow,1).getValues();

  //空の配列を準備して繰り返し処理を格納
  let overAddress = []
  let originallyAddress = []

for (let ii = 0; ii < colData.length; ii++ ) {//配列は1字と認識されるのでi++で配列を展開
  if (String(colData[ii]).length >= 20) {//C列20文字以上の時は
    overAddress[ii] = [String(colData[ii]).slice(20)]//20文字以降の文字を切り取って配列に代入
    originallyAddress[ii] = [String(colData[ii]).slice(0,20)]//20文字以前の文字を切り取って配列に代入
  } else if(String(colData[ii]).length <= 20) {//C列20文字以下の場合は
    overAddress[ii] = [""]//空を配列に代入
    originallyAddress[ii] = colData_2[ii]//元の一次元配列を二次元配列の一列目に代入
  } else {
    ;
  }
}

//切り取った配列を各列にセット
      sheet.getRange(2,4,aLastRow,1).setValues(overAddress)
      sheet.getRange(2,3,aLastRow,1).setValues(originallyAddress)

//カンマがある際に削除（置換）する
    sheet
    .getRange("D2:D" + sheet.getLastRow())  // D列を選択
    .createTextFinder(',')  // 複数の文字も検索可
    .useRegularExpression(true)  // 正規表現を使用
    .replaceAllWith("");  // 空に置換

  //D列(住所2)の空白部分に半角スペースをセット ※address_2に空白無視して文字列が代入される為
  const spaceHunter = sheet.getRange(2,4,aLastRow,1).getValues()
    for (let row = 0; row < spaceHunter.length; row++) { 
       for (let col = 0; col < 1; col++) // 配列へのインデックスrowとcolは0から
      if (spaceHunter[row][col]== "") {
        spaceHunter[row][col] = " ";
      }
    }
  sheet.getRange(2,4,aLastRow,1).setValues(spaceHunter);

  // 予定の最終行を取得
  const lastRow = sheet.getLastRow()

  //一覧をバッファに取得
  let contents = sheet.getRange(topRow, 1, sheet.getLastRow(), lastCol).getValues()

  //データがある行まで処理を繰り返す
  for(let num = 0; num < amountNum; num++){

    //済で無視するために一回目の処理以降contentsにバッファの代入を繰り返す
    if(num >= 1) {
      contents = sheet.getRange(topRow, 1, sheet.getLastRow(), lastCol).getValues()
      }

  // バッファの内容に従ってpdfを作成
  for (let i = 0; i <= lastRow - topRow; i++) {

    //「済」の場合は無視する
    if (contents[i][statusNum] === '済') {
      continue
    }

    // 値をセット
    let originPostCodeNum = contents[i][originPostCode]
    let address_1 = contents[i][addressName_1]
    let address_2 = contents[i][addressName_2]
    let faln = contents[i][firstAndLastName]

    try {
      
      //テンプレ先が空だったらそれぞれの値をセットする
      let temPostCord

      if (temSheet.getRange("C2").getValue() === ""){
       temPostCord = temSheet.getRange("C2").setValue(originPostCodeNum)
      } else if (temSheet.getRange("C15").getValue() === ""){
       temPostCord = temSheet.getRange("C15").setValue(originPostCodeNum)
      } else if (temSheet.getRange("C28").getValue() === ""){
       temPostCord = temSheet.getRange("C28").setValue(originPostCodeNum)
      } else if (temSheet.getRange("C41").getValue() === ""){
       temPostCord = temSheet.getRange("C41").setValue(originPostCodeNum)
      } else if (temSheet.getRange("J2").getValue() === ""){
       temPostCord = temSheet.getRange("J2").setValue(originPostCodeNum)
      } else if (temSheet.getRange("J15").getValue() === "") {
       temPostCord = temSheet.getRange("J15").setValue(originPostCodeNum)
      } else if (temSheet.getRange("J28").getValue() === "") {
       temPostCord = temSheet.getRange("J28").setValue(originPostCodeNum)
      } else if (temSheet.getRange("J41").getValue() === "") {
       temPostCord = temSheet.getRange("J41").setValue(originPostCodeNum)
      } else {
        continue
      }
         
      let temAddress_1

      if (temSheet.getRange("C5").getValue() === ""){
       temAddress_1 = temSheet.getRange("C5").setValue(address_1)
      } else if (temSheet.getRange("C18").getValue() === ""){
       temAddress_1 = temSheet.getRange("C18").setValue(address_1)
      } else if (temSheet.getRange("C31").getValue() === ""){
       temAddress_1 = temSheet.getRange("C31").setValue(address_1)
      } else if (temSheet.getRange("C44").getValue() === ""){
       temAddress_1 = temSheet.getRange("C44").setValue(address_1)
      } else if (temSheet.getRange("J5").getValue() === ""){
       temAddress_1 = temSheet.getRange("J5").setValue(address_1)
      } else if (temSheet.getRange("J18").getValue() === "") {
       temAddress_1 = temSheet.getRange("J18").setValue(address_1)
      } else if (temSheet.getRange("J31").getValue() === "") {
       temAddress_1 = temSheet.getRange("J31").setValue(address_1)
      } else if (temSheet.getRange("J44").getValue() === "") {
       temAddress_1 = temSheet.getRange("J44").setValue(address_1)
      } else {
        continue
      }     
      let temAddress_2

      if (temSheet.getRange("C6").getValue() === ""){
       temAddress_2 = temSheet.getRange("C6").setValue(address_2)
      } else if (temSheet.getRange("C19").getValue() === ""){
       temAddress_2 = temSheet.getRange("C19").setValue(address_2)
      } else if (temSheet.getRange("C32").getValue() === ""){
       temAddress_2 = temSheet.getRange("C32").setValue(address_2)
      } else if (temSheet.getRange("C45").getValue() === ""){
       temAddress_2 = temSheet.getRange("C45").setValue(address_2)
      } else if (temSheet.getRange("J6").getValue() === ""){
       temAddress_2 = temSheet.getRange("J6").setValue(address_2)
      } else if (temSheet.getRange("J19").getValue() === "") {
       temAddress_2 = temSheet.getRange("J19").setValue(address_2)
      } else if (temSheet.getRange("J32").getValue() === "") {
       temAddress_2 = temSheet.getRange("J32").setValue(address_2)
      } else if (temSheet.getRange("J45").getValue() === "") {
       temAddress_2 = temSheet.getRange("J45").setValue(address_2)
      } else {
        continue
      }

      let temFaln

      if (temSheet.getRange("C8").getValue() === ""){
       temFaln = temSheet.getRange("C8").setValue(faln)
      } else if (temSheet.getRange("C21").getValue() === ""){
       temFaln = temSheet.getRange("C21").setValue(faln)
      } else if (temSheet.getRange("C34").getValue() === ""){
       temFaln = temSheet.getRange("C34").setValue(faln)
      } else if (temSheet.getRange("C47").getValue() === ""){
       temFaln = temSheet.getRange("C47").setValue(faln)
      } else if (temSheet.getRange("J8").getValue() === ""){
       temFaln = temSheet.getRange("J8").setValue(faln)
      } else if (temSheet.getRange("J21").getValue() === "") {
       temFaln = temSheet.getRange("J21").setValue(faln)
      } else if (temSheet.getRange("J34").getValue() === "") {
       temFaln = temSheet.getRange("J34").setValue(faln)
      } else if (temSheet.getRange("J47").getValue() === "") {
       temFaln = temSheet.getRange("J47").setValue(faln)
      } else {
        continue
      }
    
      //予定が作成されたら「済」にする
      sheet.getRange(topRow + i, statusCellCol).setValue('済')

      // エラーの場合ログ出力する
    } catch (e) {
      Logger.log(e)
    }
  }

    //テンプレ空の場合はpdfを作らない,空じゃなかったら作る
    if (temSheet.getRange("C2:C8").getValue() === "" && temSheet.getRange("C15:C21").getValue() === ""
        && temSheet.getRange("C28:C34").getValue() === "" && temSheet.getRange("C41:C47").getValue() === ""
        && temSheet.getRange("J2:J8").getValue() === "" && temSheet.getRange("J15:J21").getValue() === ""
        && temSheet.getRange("J28:J34").getValue() === "" && temSheet.getRange("J41:J47").getValue() === "") {
          continue
        } else {
      //PDFを格納するフォルダを指定
      let folderId = "16j7rsfeq2yHWiYrUWpmuInEjerZ1GJK8"
      let folder = DriveApp.getFolderById(folderId)
      //スプレッドシートの個別シートをPDF化するために新規のスプレッドシートを作成
      let newSs = SpreadsheetApp.create("削除するシート");
      //シートを格納するフォルダを指定
      let file = DriveApp.getFileById(newSs.getId());
      let sheetFolderId = "1AeDVrrneFZjKP667z86t9fKKOe0jAa-s"
      let sheetFolder = DriveApp.getFolderById(sheetFolderId).addFile(file)
      //PDF化したい個別シートを新規作成したスプレッドシートにコピー
      temSheet.copyTo(newSs);
      //スプレッドシート新規作成でデフォルト作成されるシートを削除
      newSs.deleteSheet(newSs.getSheets()[0]);
      //PDFとしてgetAsメソッドでblob形式で取得
      let pdf = newSs.getAs('application/pdf');
      //pdfファイルの名前を設定
      pdf.setName("定形外郵便PDF " + num);
      //GoogleドライブにPDFに変換したデータを保存
      folder.createFile(pdf).getId()
        }

      //テンプレ入力部分の削除
      temSheet.getRange("C2:C8").setValue("")
      temSheet.getRange("C15:C21").setValue("")
      temSheet.getRange("C28:C34").setValue("")
      temSheet.getRange("C41:C47").setValue("")

      temSheet.getRange("J2:J8").setValue("")
      temSheet.getRange("J15:J21").setValue("")
      temSheet.getRange("J28:J34").setValue("")
      temSheet.getRange("J41:J47").setValue("")

      delFileByName()
  }
}
