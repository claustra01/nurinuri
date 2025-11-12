function nurinuri(){
  const col_min = 2; //ぬりぬり最初の列数 (B列なら2)
  const row_min = 3; //ぬりぬり最初の行数

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = [];
  for (var i = 0; i < spreadSheet.getSheets().length; i++) {
    sheets.push(spreadSheet.getSheets()[i].getSheetName());
  }
  const activeSheet = spreadSheet.getSheets()[0]; // 「集計」シートが対象となる。

  const col_max = activeSheet.getMaxColumns();
  const row_max = activeSheet.getMaxRows();
  const col_num = col_max - col_min + 1;
  const row_num = row_max - row_min + 1;
  var r = [];
  var g = [];
  var b = [];
  var num = [];
  for (var col = 0; col < col_num; col++) {
    r.push(Array(row_num).fill(0.0));
    g.push(Array(row_num).fill(0.0));
    b.push(Array(row_num).fill(0.0));
    num.push(Array(row_num).fill(0));
  }

  sheets = spreadSheet.getSheets()
  for (var s = 1; s < sheets.length; s++) {
    const sheet = sheets[s];
    // 一括読み込み
    const bg = sheet.getRange(row_min, col_min, row_max, col_max).getBackgrounds(); // bg[row][col]

    for (var row = 0; row < row_num; row++) {
      for (var col = 0; col < col_num; col++) {
        var bgcolor = bg[row][col];
        if (bgcolor != "#ffffff") {
          num[col][row] = num[col][row] + 1;
        }
        if (bgcolor.indexOf("#") == 0 && bgcolor.length == 7) {
          bgcolor = bgcolor.replace('#', '');
          r[col][row] = r[col][row] + parseInt(bgcolor.substr(0, 2), 16);
          g[col][row] = g[col][row] + parseInt(bgcolor.substr(2, 2), 16);
          b[col][row] = b[col][row] + parseInt(bgcolor.substr(4, 2), 16);
        }
      }
    }
  }

  let bgout = [];
  let valout = [];
  let fontout = [];

  for (var row = 0; row < row_num; row++) {
    bgout[row] = [];
    valout[row] = [];
    fontout[row] = [];

    for (var col = 0; col < col_num; col++) {
      r[col][row] = r[col][row] / (sheets.length - 1);
      g[col][row] = g[col][row] / (sheets.length - 1);
      b[col][row] = b[col][row] / (sheets.length - 1);
      var min = Math.min(r[col][row], g[col][row], b[col][row]);
      var max = Math.max(r[col][row], g[col][row], b[col][row]);
      var minmax = min + max;
      var averagedRGB = '#' + Math.floor(r[col][row]).toString(16).padStart(2, '0');
      averagedRGB = averagedRGB + Math.floor(g[col][row]).toString(16).padStart(2, '0');
      averagedRGB = averagedRGB + Math.floor(b[col][row]).toString(16).padStart(2, '0');
      var complementRGB = '#' + Math.floor(minmax - r[col][row]).toString(16).padStart(2, '0');
      complementRGB = complementRGB + Math.floor(minmax - g[col][row]).toString(16).padStart(2, '0');
      complementRGB = complementRGB + Math.floor(minmax - b[col][row]).toString(16).padStart(2, '0');
      bgout[row][col] = averagedRGB;
      valout[row][col] = num[col][row];
      fontout[row][col] = complementRGB;
      activeSheet.getRange(row_min + row, col_min + col).setBackground(averagedRGB);
      activeSheet.getRange(row_min + row, col_min + col).setValue(num[col][row]);
      activeSheet.getRange(row_min + row, col_min + col).setFontColor(complementRGB);
    }
  }

  // 一括書き込み
  const range = activeSheet.getRange(row_min, col_min, row_num, col_num);
  range.setBackgrounds(bgout);
  range.setValues(valout);
  range.setFontColors(fontout);
}
