/**
 * DriveManager.gs
 * 採用ドライブのフォルダ構造を自動管理
 *
 * フォルダ構造:
 * [企業名]_採用ドライブ/
 * ├─ 新卒/
 * │  ├─ 0_初回面談/
 * │  ├─ 1_1次面接/
 * │  ├─ 2_2次面接/
 * │  ├─ 3_最終面接/
 * │  ├─ 評価B以上/
 * │  ├─ 内定・承諾/
 * │  └─ エンゲージメントレポート/
 * └─ 中途/
 *    └─ （同じ構造）
 */

/**
 * ルートフォルダを取得または作成
 *
 * @param {string} companyName - 企業名（デフォルト: "テスト株式会社"）
 * @return {GoogleAppsScript.Drive.Folder} ルートフォルダ
 */
function getOrCreateRootFolder(companyName = "テスト株式会社") {
  const folderName = `${companyName}_採用ドライブ`;

  // 既存フォルダを検索
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    const folder = folders.next();
    Logger.log('✅ ルートフォルダ取得: ' + folder.getName());
    return folder;
  }

  // 新規作成
  const newFolder = DriveApp.createFolder(folderName);

  // ドメイン内全員に閲覧権限（エラー時はスキップ）
  try {
    newFolder.setSharing(
      DriveApp.Access.DOMAIN_WITH_LINK,
      DriveApp.Permission.VIEW
    );
    Logger.log('✅ 共有権限設定完了');
  } catch (error) {
    Logger.log('⚠️ 共有権限設定をスキップ（手動で設定してください）: ' + error.message);
  }

  Logger.log('✅ ルートフォルダ作成: ' + newFolder.getName());
  return newFolder;
}

/**
 * フォルダを取得または作成（ヘルパー関数）
 *
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - 親フォルダ
 * @param {string} folderName - フォルダ名
 * @return {GoogleAppsScript.Drive.Folder} フォルダ
 */
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  const newFolder = parentFolder.createFolder(folderName);

  // ドメイン内全員に閲覧権限（エラー時はスキップ）
  try {
    newFolder.setSharing(
      DriveApp.Access.DOMAIN_WITH_LINK,
      DriveApp.Permission.VIEW
    );
  } catch (error) {
    // 権限設定エラーは無視（手動で設定可能）
  }

  return newFolder;
}

/**
 * ドライブ構造を初期化（全フォルダ作成）
 *
 * @param {string} companyName - 企業名
 * @return {Object} 作成されたフォルダ情報
 */
function initializeDriveStructure(companyName = "テスト株式会社") {
  Logger.log('=== ドライブ構造初期化開始 ===');

  // ルートフォルダ作成
  const rootFolder = getOrCreateRootFolder(companyName);

  const result = {
    rootFolder: rootFolder.getName(),
    recruitTypes: []
  };

  // 新卒・中途フォルダ
  const recruitTypes = ['新卒', '中途'];

  recruitTypes.forEach(recruitType => {
    Logger.log(`--- ${recruitType}フォルダ作成開始 ---`);

    // 採用区分フォルダ
    const typeFolder = getOrCreateFolder(rootFolder, recruitType);

    // 選考フェーズフォルダ
    const phases = [
      '0_初回面談',
      '1_1次面接',
      '2_2次面接',
      '3_最終面接',
      '評価B以上',
      '内定・承諾',
      'エンゲージメントレポート'
    ];

    const createdPhases = [];
    phases.forEach(phase => {
      const phaseFolder = getOrCreateFolder(typeFolder, phase);
      createdPhases.push(phaseFolder.getName());
    });

    result.recruitTypes.push({
      type: recruitType,
      phases: createdPhases
    });

    Logger.log(`✅ ${recruitType}フォルダ作成完了`);
  });

  Logger.log('=== ドライブ構造初期化完了 ===');
  Logger.log(JSON.stringify(result, null, 2));

  return result;
}

/**
 * 選考フェーズフォルダを取得または作成
 *
 * @param {string} recruitType - 採用区分（新卒/中途）
 * @param {string} phase - 選考フェーズ（初回面談/1次面接等）
 * @param {string} companyName - 企業名（デフォルト: "テスト株式会社"）
 * @return {GoogleAppsScript.Drive.Folder} フェーズフォルダ
 */
function getOrCreatePhaseFolder(recruitType, phase, companyName = "テスト株式会社") {
  const rootFolder = getOrCreateRootFolder(companyName);
  const typeFolder = getOrCreateFolder(rootFolder, recruitType);

  // フェーズ名の正規化
  const phaseMap = {
    '初回面談': '0_初回面談',
    '1次面接': '1_1次面接',
    '2次面接': '2_2次面接',
    '最終面接': '3_最終面接'
  };

  const normalizedPhase = phaseMap[phase] || phase;
  const phaseFolder = getOrCreateFolder(typeFolder, normalizedPhase);

  Logger.log(`✅ フェーズフォルダ取得: ${recruitType}/${normalizedPhase}`);
  return phaseFolder;
}

/**
 * 候補者フォルダを取得または作成
 *
 * @param {string} recruitType - 採用区分
 * @param {string} phase - 選考フェーズ
 * @param {string} candidateId - 候補者ID（CAND_YYYYMMDD_HHMMSS）
 * @param {string} candidateName - 候補者名
 * @param {string} companyName - 企業名（デフォルト: "テスト株式会社"）
 * @return {GoogleAppsScript.Drive.Folder} 候補者フォルダ
 */
function getOrCreateCandidateFolder(recruitType, phase, candidateId, candidateName, companyName = "テスト株式会社") {
  const phaseFolder = getOrCreatePhaseFolder(recruitType, phase, companyName);

  // フォルダ名: CAND_20251218_山田太郎
  const folderName = `${candidateId}_${candidateName}`;

  const candidateFolder = getOrCreateFolder(phaseFolder, folderName);

  Logger.log(`✅ 候補者フォルダ取得: ${folderName}`);
  return candidateFolder;
}

/**
 * candidate_idで全候補者フォルダを検索
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} companyName - 企業名（デフォルト: "テスト株式会社"）
 * @return {Array<Object>} フォルダ情報の配列
 */
function findCandidateFolders(candidateId, companyName = "テスト株式会社") {
  Logger.log(`=== 候補者フォルダ検索: ${candidateId} ===`);

  const rootFolder = getOrCreateRootFolder(companyName);
  const results = [];

  // ルートフォルダ配下を再帰的に検索
  function searchInFolder(folder, path) {
    const subfolders = folder.getFolders();

    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      const folderName = subfolder.getName();
      const currentPath = path + '/' + folderName;

      // candidate_idで始まるフォルダを検出
      if (folderName.startsWith(candidateId)) {
        results.push({
          folder: subfolder,
          path: currentPath,
          id: subfolder.getId(),
          url: subfolder.getUrl()
        });
        Logger.log(`✅ 発見: ${currentPath}`);
      }

      // 再帰的に検索
      searchInFolder(subfolder, currentPath);
    }
  }

  searchInFolder(rootFolder, rootFolder.getName());

  Logger.log(`検索結果: ${results.length}件`);
  return results;
}

/**
 * フォルダを再帰的にコピー（ヘルパー関数）
 *
 * @param {GoogleAppsScript.Drive.Folder} sourceFolder - コピー元フォルダ
 * @param {GoogleAppsScript.Drive.Folder} destinationParent - コピー先親フォルダ
 * @return {GoogleAppsScript.Drive.Folder} コピーされたフォルダ
 */
function copyFolderRecursive(sourceFolder, destinationParent) {
  // 新しいフォルダ作成
  const newFolder = destinationParent.createFolder(sourceFolder.getName());

  // 権限設定（エラー時はスキップ）
  try {
    newFolder.setSharing(
      DriveApp.Access.DOMAIN_WITH_LINK,
      DriveApp.Permission.VIEW
    );
  } catch (error) {
    // 権限設定エラーは無視（手動で設定可能）
  }

  // ファイルをコピー
  const files = sourceFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    file.makeCopy(file.getName(), newFolder);
  }

  // サブフォルダを再帰的にコピー
  const subfolders = sourceFolder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    copyFolderRecursive(subfolder, newFolder);
  }

  return newFolder;
}

/**
 * 候補者フォルダを評価B以上フォルダにコピー
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} candidateName - 候補者名
 * @param {string} recruitType - 採用区分
 * @param {string} companyName - 企業名（デフォルト: "テスト株式会社"）
 * @return {Object} コピー結果
 */
function copyFolderToGradeB(candidateId, candidateName, recruitType, companyName = "テスト株式会社") {
  Logger.log('=== 評価B以上フォルダへコピー開始 ===');

  try {
    // 元フォルダを検索
    const sourceFolders = findCandidateFolders(candidateId, companyName);

    if (sourceFolders.length === 0) {
      throw new Error(`候補者フォルダが見つかりません: ${candidateId}`);
    }

    // 最新のフォルダを使用（通常は1件のはず）
    const sourceFolder = sourceFolders[0].folder;

    // 評価B以上フォルダ取得
    const gradeBFolder = getOrCreatePhaseFolder(recruitType, '評価B以上', companyName);

    // 同名フォルダが既に存在するか確認
    const folderName = `${candidateId}_${candidateName}`;
    const existing = gradeBFolder.getFoldersByName(folderName);

    if (existing.hasNext()) {
      Logger.log('⚠️ 既にコピー済み: ' + folderName);
      return {
        success: true,
        message: '既にコピー済み',
        folder: existing.next()
      };
    }

    // フォルダをコピー
    const copiedFolder = copyFolderRecursive(sourceFolder, gradeBFolder);

    Logger.log('✅ コピー完了: ' + copiedFolder.getUrl());

    return {
      success: true,
      message: 'コピー成功',
      folder: copiedFolder,
      url: copiedFolder.getUrl()
    };

  } catch (error) {
    Logger.log('❌ コピーエラー: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 候補者フォルダを内定・承諾フォルダに移動
 * 元フォルダは削除
 *
 * @param {string} candidateId - 候補者ID
 * @param {string} candidateName - 候補者名
 * @param {string} recruitType - 採用区分
 * @param {string} currentPhase - 現在の選考フェーズ
 * @param {string} companyName - 企業名（デフォルト: "テスト株式会社"）
 * @return {Object} 移動結果
 */
function moveFolderToAccepted(candidateId, candidateName, recruitType, currentPhase, companyName = "テスト株式会社") {
  Logger.log('=== 内定・承諾フォルダへ移動開始 ===');

  try {
    // 元フォルダを検索
    const sourceFolders = findCandidateFolders(candidateId, companyName);

    if (sourceFolders.length === 0) {
      throw new Error(`候補者フォルダが見つかりません: ${candidateId}`);
    }

    // 内定・承諾フォルダ取得
    const acceptedFolder = getOrCreatePhaseFolder(recruitType, '内定・承諾', companyName);

    // フォルダ名
    const folderName = `${candidateId}_${candidateName}`;

    // 既に存在するか確認
    const existing = acceptedFolder.getFoldersByName(folderName);

    if (existing.hasNext()) {
      Logger.log('⚠️ 既に移動済み: ' + folderName);

      // 元フォルダを削除
      sourceFolders.forEach(item => {
        if (!item.path.includes('内定・承諾')) {
          item.folder.setTrashed(true);
          Logger.log(`✅ 元フォルダ削除: ${item.path}`);
        }
      });

      return {
        success: true,
        message: '既に移動済み（元フォルダ削除）',
        folder: existing.next()
      };
    }

    // 最初のフォルダをコピーして移動
    const sourceFolder = sourceFolders[0].folder;
    const movedFolder = copyFolderRecursive(sourceFolder, acceptedFolder);

    // 全ての元フォルダを削除
    sourceFolders.forEach(item => {
      item.folder.setTrashed(true);
      Logger.log(`✅ 元フォルダ削除: ${item.path}`);
    });

    Logger.log('✅ 移動完了: ' + movedFolder.getUrl());

    return {
      success: true,
      message: '移動成功',
      folder: movedFolder,
      url: movedFolder.getUrl(),
      deletedFolders: sourceFolders.length
    };

  } catch (error) {
    Logger.log('❌ 移動エラー: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ドライブ構造管理のテスト
 */
function testDriveManager() {
  Logger.log('=== DriveManager テスト開始 ===');

  const companyName = 'テスト株式会社';  // 全テストで同じ企業名を使用

  // 1. 構造初期化
  Logger.log('--- テスト1: 構造初期化 ---');
  const initResult = initializeDriveStructure(companyName);
  Logger.log('結果: ' + JSON.stringify(initResult, null, 2));

  // 2. 候補者フォルダ作成
  Logger.log('--- テスト2: 候補者フォルダ作成 ---');
  const candidateId = 'CAND_20251218_120000';
  const candidateName = 'テスト太郎';
  const candidateFolder = getOrCreateCandidateFolder(
    '新卒',
    '1次面接',
    candidateId,
    candidateName,
    companyName  // 企業名を指定
  );
  Logger.log('フォルダURL: ' + candidateFolder.getUrl());

  // 3. フォルダ検索
  Logger.log('--- テスト3: フォルダ検索 ---');
  const found = findCandidateFolders(candidateId, companyName);  // 企業名を指定
  Logger.log(`検索結果: ${found.length}件`);

  // 4. 評価B以上にコピー
  Logger.log('--- テスト4: 評価B以上にコピー ---');
  const copyResult = copyFolderToGradeB(candidateId, candidateName, '新卒', companyName);  // 企業名を指定
  Logger.log('結果: ' + JSON.stringify(copyResult, null, 2));

  // 5. 内定・承諾に移動
  Logger.log('--- テスト5: 内定・承諾に移動 ---');
  const moveResult = moveFolderToAccepted(candidateId, candidateName, '新卒', '1次面接', companyName);  // 企業名を指定
  Logger.log('結果: ' + JSON.stringify(moveResult, null, 2));

  Logger.log('=== DriveManager テスト完了 ===');
}
