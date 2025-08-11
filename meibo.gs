function 集計して別シートに出力() {
  const sheetName = "名簿";  // 元データのシート名になる
  const outputSheetName = "会費集計";  // 出力用シート名になる
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();

  const headers = data[2]; // 3行目がヘッダー
  const rows = data.slice(3); // データ本体（4行目から）

  const 保護者名Index = headers.indexOf("保護者名");
  if (保護者名Index === -1) {
    throw new Error("「保護者名」列が見つかりません。");
  }

  // 集計マップ
  const grouped = {};//保護者名 → 児童数 の対応を保存する連想配列

  rows.forEach(row => {
    let guardianRaw = row[保護者名Index];//○○から
    if (!guardianRaw) return;//値なしで終了する

    const guardian = guardianRaw.toString().trim().replace(/　/g, "");  // 前後トリム＋全角スペース除去
    if (!grouped[guardian]) {//保護者オブジェクトがないときは
      grouped[guardian] = 0;
    }
    grouped[guardian]++;//保護者の児童をカウント ○○：0
  });

    // 出力用データ作成
  const output = [["保護者名", "児童人数", "児童会費", "保護者会費", "合計会費"]];
  let total = 0;
  for (let guardian in grouped) {
    const children = grouped[guardian];
    const childrenFee = children * 1200;
    const guardianFee = 1200;
    const totalFee = childrenFee + guardianFee;
    total += totalFee;
    output.push([guardian, children, childrenFee, guardianFee, totalFee]);
  }

  // 合計行を追加（空行は挿入しない）
  output.push(["", "", "", "合計", total]);

  // 出力先シートを作成またはクリア
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clear();
  }

  // 出力
  outputSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  //1,1は書き込み位置,output.lengthは行数、output[0].lengthは列数,setValuesは行ごとに順番に出力
}


function 世帯ごとの領収書を作成する() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("会費集計");
  const outputSheetName = "領収書";

  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1); // 実データ

  let outputSheet = ss.getSheetByName(outputSheetName);
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet(outputSheetName);
  }

  let currentRow = 1;
  const 会費 = 1200;

  rows.forEach(entry => {
    const 保護者名 = entry[headers.indexOf("保護者名")];
    const 人数 = entry[headers.indexOf("児童人数")];

    if (!保護者名 || !人数) return;

    const 合計金額 = 会費 * (Number(人数) + 1);

    // 出力先：8行 × 6列の領域を確保
    const blockRange = outputSheet.getRange(currentRow, 1, 8, 6);
    blockRange.setBorder(true, true, true, true, false, false);
    for (let i = 0; i < 8; i++) {
      outputSheet.setRowHeight(currentRow + i, 24);
    }
    outputSheet.setColumnWidths(1, 6, 100);

    // 1行ずつ出力（セルごとに配置指定）
    const lines = [
      { text: "領収書", align: "center", bold: true, size: 14 },
      { text: `　　　${保護者名} 様`, align: "left", bold: false, size: 12 },
      { text: `　合計　${合計金額.toLocaleString()}円（1,200円 × ${人数}名）`, align: "center", bold: false, size: 12 },
      { text: "", align: "center", bold: false, size: 12 }, // 空行
      { text: "2025年度子ども会費を領収いたしました。", align: "center", bold: false, size: 12 },
      { text: "2025年4月", align: "right", bold: false, size: 12 },
      { text: "柳町子ども会", align: "right", bold: false, size: 12 },
    ];

    lines.forEach((line, index) => {
      const row = currentRow + index;
      const range = outputSheet.getRange(row, 1, 1, 6); //A列からF列までの横長セル1行分
      range.merge();                            //その行の6セル（A～F）を1つに結合
      range.setValue(line.text);                // テキストを入れる
      range.setFontSize(line.size);             // フォントサイズ設定
      range.setFontWeight(line.bold ? "bold" : "normal"); // 太字か普通か
      range.setVerticalAlignment("middle");     // 縦方向の中央揃え
      range.setHorizontalAlignment(line.align); // 横方向（left/center/right）の揃え

    });

    currentRow += 10; // 次の領収書へ（少し間隔を空ける）
  });

  SpreadsheetApp.flush();
}

