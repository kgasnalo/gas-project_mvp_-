/**
 * スプレッドシート完全出力スクリプト
 *
 * Google Apps Script用のスプレッドシート情報を完全にJSON形式で出力するコード
 *
 * 機能：
 * - スプレッドシート基本情報の取得
 * - 全タブの詳細情報取得
 * - 全セルの値・数式・メモの取得
 * - 全セルの書式情報の取得
 * - JSON形式での出力
 *
 * 使い方：
 * 1. このスクリプトをGASエディタにコピー
 * 2. スプレッドシートを開いた状態で実行
 * 3. 「実行ログ」（Ctrl+Enter）で出力を確認
 *
 * オプション：
 * - maxCellsPerSheet: シートあたりの最大セル数制限（デフォルト: 100000）
 * - includeEmptyCells: 空セルを含めるかどうか（デフォルト: false）
 *
 * @author GAS Project
 * @version 1.0.0
 */

/**
 * スプレッドシートのすべての情報をJSON形式で出力するメイン関数
 *
 * @param {Object} options - オプション設定
 * @param {number} options.maxCellsPerSheet - シートあたりの最大セル数（デフォルト: 100000）
 * @param {boolean} options.includeEmptyCells - 空セルを含めるか（デフォルト: false）
 * @returns {Object} スプレッドシートの完全な情報を含むオブジェクト
 */
function getUltraCompleteData(options = {}) {
  try {
    // オプションのデフォルト値を設定
    const config = {
      maxCellsPerSheet: options.maxCellsPerSheet || 100000,
      includeEmptyCells: options.includeEmptyCells !== undefined ? options.includeEmptyCells : false
    };

    Logger.log('=== スプレッドシート完全出力を開始 ===');
    const startTime = new Date();

    // アクティブなスプレッドシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (!spreadsheet) {
      throw new Error('アクティブなスプレッドシートが見つかりません');
    }

    // スプレッドシート基本情報を取得
    const spreadsheetInfo = getSpreadsheetInfo(spreadsheet);
    Logger.log(`スプレッドシート名: ${spreadsheetInfo.name}`);

    // 全シートの情報を取得
    const sheets = spreadsheet.getSheets();
    Logger.log(`シート数: ${sheets.length}`);

    const sheetsData = [];
    let totalCells = 0;

    // 各シートを処理
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      Logger.log(`\n[${i + 1}/${sheets.length}] シート「${sheet.getName()}」を処理中...`);

      try {
        const sheetData = getSheetCompleteData(sheet, config);
        sheetsData.push(sheetData);

        // セル数をカウント
        const cellCount = sheetData.data.reduce((sum, row) => sum + row.length, 0);
        totalCells += cellCount;
        Logger.log(`  - セル数: ${cellCount}`);

        // 大量データの警告
        if (cellCount > config.maxCellsPerSheet) {
          Logger.log(`  ⚠️ 警告: シート「${sheet.getName()}」のセル数が制限を超えています（${cellCount} > ${config.maxCellsPerSheet}）`);
        }
      } catch (error) {
        Logger.log(`  ❌ エラー: シート「${sheet.getName()}」の処理に失敗: ${error.message}`);
        // エラーが発生したシートもスキップせず、エラー情報を含める
        sheetsData.push({
          name: sheet.getName(),
          index: sheet.getIndex(),
          error: error.message,
          data: []
        });
      }
    }

    // 最終的なデータ構造を構築
    const result = {
      spreadsheet: spreadsheetInfo,
      sheets: sheetsData,
      metadata: {
        exportDate: new Date().toISOString(),
        totalSheets: sheets.length,
        totalCells: totalCells,
        options: config
      }
    };

    // 処理時間を計算
    const endTime = new Date();
    const processingTime = (endTime - startTime) / 1000;
    Logger.log(`\n=== 処理完了 ===`);
    Logger.log(`処理時間: ${processingTime.toFixed(2)}秒`);
    Logger.log(`総セル数: ${totalCells}`);

    // JSON文字列に変換して出力
    const jsonOutput = JSON.stringify(result, null, 2);

    // ログサイズの警告
    if (jsonOutput.length > 100000) {
      Logger.log(`\n⚠️ 警告: 出力データが非常に大きいです（${(jsonOutput.length / 1024).toFixed(2)} KB）`);
      Logger.log('Logger.logの出力制限により、すべてのデータが表示されない可能性があります');
      Logger.log('大量データの場合は、Google Driveへの出力を検討してください');
    }

    // JSONを出力
    Logger.log('\n=== JSON出力 ===');
    Logger.log(jsonOutput);

    return result;

  } catch (error) {
    // エラーハンドリング
    Logger.log(`\n❌ 致命的なエラーが発生しました`);
    Logger.log(`エラーメッセージ: ${error.message}`);
    Logger.log(`スタックトレース: ${error.stack}`);
    throw error;
  }
}

/**
 * スプレッドシートの基本情報を取得
 *
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @returns {Object} スプレッドシートの基本情報
 */
function getSpreadsheetInfo(spreadsheet) {
  return {
    name: spreadsheet.getName(),
    id: spreadsheet.getId(),
    url: spreadsheet.getUrl(),
    locale: spreadsheet.getSpreadsheetLocale(),
    timezone: spreadsheet.getSpreadsheetTimeZone()
  };
}

/**
 * シートの完全な情報を取得
 *
 * @param {Sheet} sheet - シートオブジェクト
 * @param {Object} config - 設定オブジェクト
 * @returns {Object} シートの完全な情報
 */
function getSheetCompleteData(sheet, config) {
  // シートの基本情報を取得
  const sheetInfo = {
    name: sheet.getName(),
    index: sheet.getIndex(),
    dimensions: getSheetDimensions(sheet),
    properties: getSheetProperties(sheet),
    data: []
  };

  // データ範囲を取得（空でない場合のみ）
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0 || lastColumn === 0) {
    Logger.log(`  - シートにデータがありません`);
    return sheetInfo;
  }

  // データ範囲のチェック
  const totalCells = lastRow * lastColumn;
  if (totalCells > config.maxCellsPerSheet) {
    Logger.log(`  ⚠️ セル数が制限を超えています。最初の${config.maxCellsPerSheet}セルのみ取得します`);
    // 制限に基づいて範囲を調整
    const adjustedRows = Math.min(lastRow, Math.floor(config.maxCellsPerSheet / lastColumn));
    return getSheetDataInRange(sheet, 1, 1, adjustedRows, lastColumn, config);
  }

  // 全データを取得
  return getSheetDataInRange(sheet, 1, 1, lastRow, lastColumn, config);
}

/**
 * 指定範囲のシートデータを取得
 *
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} startRow - 開始行（1始まり）
 * @param {number} startCol - 開始列（1始まり）
 * @param {number} numRows - 行数
 * @param {number} numCols - 列数
 * @param {Object} config - 設定オブジェクト
 * @returns {Object} シート情報とデータ
 */
function getSheetDataInRange(sheet, startRow, startCol, numRows, numCols, config) {
  const sheetInfo = {
    name: sheet.getName(),
    index: sheet.getIndex(),
    dimensions: getSheetDimensions(sheet),
    properties: getSheetProperties(sheet),
    data: []
  };

  // 範囲を取得
  const range = sheet.getRange(startRow, startCol, numRows, numCols);

  // 一括でデータを取得（パフォーマンス最適化）
  const values = range.getValues();
  const formulas = range.getFormulas();
  const notes = range.getNotes();
  const backgrounds = range.getBackgrounds();
  const fontColors = range.getFontColors();
  const fontFamilies = range.getFontFamilies();
  const fontSizes = range.getFontSizes();
  const fontWeights = range.getFontWeights();
  const horizontalAlignments = range.getHorizontalAlignments();
  const verticalAlignments = range.getVerticalAlignments();

  // 各行を処理
  for (let row = 0; row < numRows; row++) {
    const rowData = [];

    for (let col = 0; col < numCols; col++) {
      const value = values[row][col];
      const formula = formulas[row][col];
      const note = notes[row][col];

      // 空セルをスキップするオプション
      if (!config.includeEmptyCells && value === '' && formula === '' && note === '') {
        continue;
      }

      // セル情報を構築
      const cellData = {
        position: {
          row: startRow + row,
          column: startCol + col,
          a1: columnToLetter(startCol + col) + (startRow + row)
        },
        content: {
          value: value,
          formula: formula || null,
          note: note || null
        },
        formatting: {
          background: backgrounds[row][col] || '#ffffff',
          fontColor: fontColors[row][col] || '#000000',
          fontFamily: fontFamilies[row][col] || 'Arial',
          fontSize: fontSizes[row][col] || 10,
          fontWeight: fontWeights[row][col] || 'normal',
          horizontalAlignment: horizontalAlignments[row][col] || 'left',
          verticalAlignment: verticalAlignments[row][col] || 'bottom'
        }
      };

      rowData.push(cellData);
    }

    // 空でない行のみ追加
    if (rowData.length > 0) {
      sheetInfo.data.push(rowData);
    }
  }

  return sheetInfo;
}

/**
 * シートの寸法情報を取得
 *
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {Object} 寸法情報
 */
function getSheetDimensions(sheet) {
  return {
    rows: sheet.getLastRow(),
    columns: sheet.getLastColumn(),
    maxRows: sheet.getMaxRows(),
    maxColumns: sheet.getMaxColumns()
  };
}

/**
 * シートのプロパティ情報を取得
 *
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {Object} プロパティ情報
 */
function getSheetProperties(sheet) {
  const properties = {
    isHidden: sheet.isSheetHidden(),
    tabColor: null,
    isProtected: false
  };

  // タブの色を取得（存在する場合）
  try {
    const tabColor = sheet.getTabColor();
    properties.tabColor = tabColor;
  } catch (e) {
    // タブの色が設定されていない場合
    properties.tabColor = null;
  }

  // 保護設定を確認
  try {
    const protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    properties.isProtected = protection.length > 0;
  } catch (e) {
    properties.isProtected = false;
  }

  return properties;
}

/**
 * 列番号をA1形式の列文字に変換
 *
 * @param {number} column - 列番号（1始まり）
 * @returns {string} A1形式の列文字（A, B, C, ..., Z, AA, AB, ...）
 */
function columnToLetter(column) {
  let temp;
  let letter = '';

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

/**
 * 結果をGoogle Driveに保存する補助関数
 * 大量データの場合に使用することを推奨
 *
 * @param {Object} data - 保存するデータオブジェクト
 * @param {string} filename - ファイル名（デフォルト: 'spreadsheet_export.json'）
 * @returns {string} 作成されたファイルのURL
 */
function saveToGoogleDrive(data, filename = 'spreadsheet_export.json') {
  try {
    const jsonString = JSON.stringify(data, null, 2);
    const blob = Utilities.newBlob(jsonString, 'application/json', filename);
    const file = DriveApp.createFile(blob);

    Logger.log(`\n✅ ファイルをGoogle Driveに保存しました`);
    Logger.log(`ファイル名: ${filename}`);
    Logger.log(`ファイルURL: ${file.getUrl()}`);

    return file.getUrl();
  } catch (error) {
    Logger.log(`\n❌ Google Driveへの保存に失敗しました: ${error.message}`);
    throw error;
  }
}

/**
 * 簡易版: 基本情報のみを出力する関数
 * テスト用または概要確認用
 *
 * @returns {Object} スプレッドシートの基本情報
 */
function getBasicSpreadsheetInfo() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();

    const result = {
      spreadsheet: getSpreadsheetInfo(spreadsheet),
      sheets: sheets.map(sheet => ({
        name: sheet.getName(),
        index: sheet.getIndex(),
        dimensions: getSheetDimensions(sheet),
        properties: getSheetProperties(sheet)
      }))
    };

    Logger.log(JSON.stringify(result, null, 2));
    return result;
  } catch (error) {
    Logger.log(`❌ エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 使用例とテストコード
 *
 * 実行方法：
 * 1. 基本実行（デフォルト設定）
 *    getUltraCompleteData();
 *
 * 2. 空セルを含めて出力
 *    getUltraCompleteData({ includeEmptyCells: true });
 *
 * 3. セル数制限を変更
 *    getUltraCompleteData({ maxCellsPerSheet: 50000 });
 *
 * 4. Google Driveに保存
 *    const data = getUltraCompleteData();
 *    saveToGoogleDrive(data, 'my_spreadsheet_export.json');
 *
 * 5. 基本情報のみ取得（テスト用）
 *    getBasicSpreadsheetInfo();
 */
