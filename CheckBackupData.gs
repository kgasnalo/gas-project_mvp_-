/**
 * バックアップシートのデータを確認
 */
function checkBackupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupSheet = ss.getSheetByName('Candidates_Master_BACKUP_20251128_125348');

  if (!backupSheet) {
    Logger.log('❌ バックアップシートが見つかりません');
    return;
  }

  Logger.log('========================================');
  Logger.log('バックアップシート確認');
  Logger.log('========================================');
  Logger.log('シート名: ' + backupSheet.getName());
  Logger.log('列数: ' + backupSheet.getLastColumn());
  Logger.log('行数（データ）: ' + (backupSheet.getLastRow() - 1));
  Logger.log('');

  // ヘッダー行を取得
  const headers = backupSheet.getRange(1, 1, 1, backupSheet.getLastColumn()).getValues()[0];
  Logger.log('列構成:');
  headers.forEach((header, index) => {
    Logger.log(`  ${index + 1}. ${header}`);
  });

  Logger.log('');
  Logger.log('重要な列の確認:');
  const importantColumns = [
    'candidate_id',
    '氏名',
    '最新_合格可能性',
    '前回_合格可能性',
    '最新_承諾可能性（統合）',
    'コアモチベーション',
    '主要懸念事項'
  ];

  importantColumns.forEach(colName => {
    const index = headers.indexOf(colName);
    if (index !== -1) {
      Logger.log(`  ✅ ${colName} (列${index + 1})`);
    } else {
      Logger.log(`  ❌ ${colName} (見つかりません)`);
    }
  });

  Logger.log('');
  Logger.log('========================================');
}

/**
 * 旧構造バックアップから現在のCandidates_Masterにデータを復元
 * （21列構造はそのまま、データのみを復元）
 */
function restoreFromBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupSheet = ss.getSheetByName('Candidates_Master_BACKUP_20251128_125348');
  const currentMaster = ss.getSheetByName('Candidates_Master');

  if (!backupSheet) {
    Logger.log('❌ バックアップシートが見つかりません');
    return;
  }

  Logger.log('========================================');
  Logger.log('バックアップからの復元開始');
  Logger.log('========================================');
  Logger.log('');

  // 新しいバックアップを作成（現在の21列構造を保存）
  const newBackup = currentMaster.copyTo(ss);
  const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
  newBackup.setName('Candidates_Master_BACKUP_' + timestamp);
  newBackup.hideSheet();
  Logger.log('✅ 現在の状態をバックアップ: ' + newBackup.getName());
  Logger.log('');

  // バックアップシート（旧構造）のデータとヘッダーを取得
  const backupHeaders = backupSheet.getRange(1, 1, 1, backupSheet.getLastColumn()).getValues()[0];
  const backupDataRows = backupSheet.getLastRow() - 1;

  if (backupDataRows === 0) {
    Logger.log('⚠️ バックアップシートにデータがありません');
    return;
  }

  const backupData = backupSheet.getRange(2, 1, backupDataRows, backupSheet.getLastColumn()).getValues();

  Logger.log('バックアップデータ: ' + backupDataRows + '件');
  Logger.log('');

  // 現在のCandidates_Masterのヘッダーとデータをクリア（ヘッダーは保持）
  if (currentMaster.getLastRow() > 1) {
    currentMaster.getRange(2, 1, currentMaster.getLastRow() - 1, currentMaster.getLastColumn()).clearContent();
  }

  // バックアップシートの全データを現在のシートにコピー
  Logger.log('復元方法: 旧構造（57列）に完全復元');
  Logger.log('');

  // 現在のシートをクリアして、バックアップの構造とデータをコピー
  currentMaster.clear();

  const allBackupData = backupSheet.getRange(1, 1, backupSheet.getLastRow(), backupSheet.getLastColumn()).getValues();
  currentMaster.getRange(1, 1, allBackupData.length, allBackupData[0].length).setValues(allBackupData);

  Logger.log('✅ 旧構造（57列）に復元完了');
  Logger.log('');
  Logger.log('========================================');
  Logger.log('次のステップ');
  Logger.log('========================================');
  Logger.log('1. phase1_execute() を実行して新構造に移行');
  Logger.log('   - Candidate_ScoresとCandidate_Insightsが作成されます');
  Logger.log('   - データが移行されます');
  Logger.log('   - Candidates_Masterが21列に再構築されます');
  Logger.log('');
  Logger.log('phase1_execute(); を実行してください');
  Logger.log('========================================');
}
