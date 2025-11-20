/**
 * ========================================
 * Phase 3.5: テストデータ拡充
 * ========================================
 *
 * 目的: 候補者5名 × 4フェーズ = 20パターンのテストデータを投入
 *
 * データパターン:
 * - C001: 理想的なケース（志望度上昇、接点密、高評価）
 * - C002: 志望度下降ケース（興味減少、懸念増加）
 * - C003: 高評価ケース（全スコア高い）
 * - C004: 低評価ケース（接点少ない、志望度低い）
 * - C005: 標準ケース（平均的なスコア）
 */

/**
 * Candidates_MasterのY列・Z列にデータを投入
 */
function insertCandidatesMasterExtended() {
  try {
    Logger.log('\n========================================');
    Logger.log('Candidates_Master拡張データ投入開始');
    Logger.log('========================================\n');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);

    if (!master) {
      Logger.log('❌ Candidates_Masterシートが見つかりません');
      return false;
    }

    const data = master.getDataRange().getValues();

    // Y列・Z列のデータ
    const extendedData = {
      'C001': {
        coreMotivation: '成長機会・キャリアアップ',
        topConcern: '業務内容の具体性'
      },
      'C002': {
        coreMotivation: '高年収・待遇',
        topConcern: '給与水準・評価制度'
      },
      'C003': {
        coreMotivation: '企業文化・働きやすさ',
        topConcern: '特になし'
      },
      'C004': {
        coreMotivation: '安定性・福利厚生',
        topConcern: '勤務地・転勤の可能性'
      },
      'C005': {
        coreMotivation: 'スキル開発・自己実現',
        topConcern: '成長機会の具体性'
      }
    };

    // ヘッダー行をスキップ（i=1から開始）
    for (let i = 1; i < data.length; i++) {
      const candidateId = data[i][0];

      if (extendedData[candidateId]) {
        // Y列（インデックス24）: コアモチベーション
        master.getRange(i + 1, 25).setValue(extendedData[candidateId].coreMotivation);

        // Z列（インデックス25）: 主要懸念事項
        master.getRange(i + 1, 26).setValue(extendedData[candidateId].topConcern);

        Logger.log(`✅ ${candidateId}: ${extendedData[candidateId].coreMotivation} / ${extendedData[candidateId].topConcern}`);
      }
    }

    Logger.log('\n✅ Candidates_Master拡張データ投入完了');
    Logger.log('========================================\n');

    return true;

  } catch (error) {
    Logger.log(`❌ エラー: ${error}`);
    return false;
  }
}

/**
 * 全アンケートデータの投入
 */
function insertExpandedSurveyResponses() {
  try {
    Logger.log('\n========================================');
    Logger.log('全アンケートデータ投入開始');
    Logger.log('========================================\n');

    // 初回面談アンケート
    insertSurvey初回面談();

    // 社員面談アンケート
    insertSurvey社員面談();

    // 2次面接アンケート
    insertSurvey2次面接();

    // 内定後アンケート
    insertSurvey内定後();

    Logger.log('\n✅ 全アンケートデータ投入完了');
    Logger.log('========================================\n');

    return true;

  } catch (error) {
    Logger.log(`❌ エラー: ${error}`);
    return false;
  }
}

/**
 * 初回面談アンケートデータの投入
 */
function insertSurvey初回面談() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('アンケート_初回面談');

  if (!sheet) {
    Logger.log('❌ アンケート_初回面談シートが見つかりません');
    return;
  }

  // 既存データをクリア（ヘッダー以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  const responses = [
    // C001: 理想的なケース
    [
      new Date('2025-10-02 14:30:00'), // A: タイムスタンプ
      '田中太郎', // B: Q1. お名前
      'tanaka@example.com', // C: Q2. メールアドレス
      8, // D: Q3. 初回面談の満足度（1-10）
      '大きく上がった', // E: Q4. PIGNUSへの興味の変化
      8, // F: Q5. PIGNUSで働きたい気持ち（1-10）
      '業務内容がより明確になった', // G: Q6. PIGNUSの魅力
      'スタートアップの安定性', // H: Q7. キャリアへの懸念
      '', // I: Q8. 他に気になること
      '業務の具体的な内容', // J: Q9. 不安・懸念事項
      2, // K: Q10-1. 選考中の企業数
      0, // L: Q10-2. 内定済みの企業数
      '非常に良かった。具体的な業務内容が聞けて安心した。', // M: Q12. その他コメント
      'ビジョンに共感した' // N: Q13. 志望理由
    ],

    // C002: 志望度下降ケース
    [
      new Date('2025-10-03 15:00:00'),
      '佐藤花子',
      'sato@example.com',
      6,
      'やや下がった', // 志望度が下がった
      6,
      '待遇面',
      '給与水準が想定より低い',
      '昇給の仕組みについて',
      '給与・評価制度',
      3, // 選考中企業が多い
      1, // 内定済みが1社
      '給与面での説明がもう少し欲しかった',
      '待遇重視'
    ],

    // C003: 高評価ケース
    [
      new Date('2025-10-04 16:00:00'),
      '鈴木一郎',
      'suzuki@example.com',
      10,
      '大きく上がった',
      10,
      'すべての面で魅力的',
      '特になし',
      '',
      'なし',
      0, // 選考中企業なし
      0, // 内定なし
      '非常に良い印象。ぜひ次に進みたい。',
      '企業文化とビジョンに強く共感'
    ],

    // C004: 低評価ケース
    [
      new Date('2025-10-05 17:00:00'),
      '高橋次郎',
      'takahashi@example.com',
      5,
      '変わらない',
      5,
      '勤務地',
      '業務内容が不明確',
      '具体的な仕事内容',
      '業務内容・役割',
      5, // 選考中企業が多い
      2, // 内定済みが2社
      '具体的な業務イメージが湧かなかった',
      '勤務地が近い'
    ],

    // C005: 標準ケース
    [
      new Date('2025-10-06 14:00:00'),
      '伊藤三郎',
      'ito@example.com',
      7,
      'やや上がった',
      7,
      '成長機会',
      'キャリアパスの具体性',
      '研修制度について',
      '成長機会の具体性',
      2,
      0,
      '概ね良かった。次のステップに進みたい。',
      'スキルアップ'
    ]
  ];

  // データを投入
  sheet.getRange(2, 1, responses.length, responses[0].length).setValues(responses);

  Logger.log(`✅ 初回面談アンケート: ${responses.length}件投入`);
}

/**
 * 社員面談アンケートデータの投入
 */
function insertSurvey社員面談() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('アンケート_社員面談');

  if (!sheet) {
    Logger.log('❌ アンケート_社員面談シートが見つかりません');
    return;
  }

  // 既存データをクリア
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  const responses = [
    // C001: 志望度さらに上昇
    [
      new Date('2025-10-10 15:00:00'),
      '田中太郎',
      'tanaka@example.com',
      9,
      '上がった',
      9,
      '社員の方の雰囲気が良かった',
      '業務内容がより明確になった',
      '',
      '特にない',
      '懸念が解消された',
      1,
      0,
      '',
      '非常に良い雰囲気だった',
      '社員の方の話が参考になった',
      'ぜひ次に進みたい'
    ],

    // C002: 志望度さらに下降
    [
      new Date('2025-10-11 16:00:00'),
      '佐藤花子',
      'sato@example.com',
      5,
      'やや下がった',
      5,
      '社員の雰囲気',
      '給与面での不安は残る',
      '評価制度の詳細',
      '給与・評価制度',
      '給与面の説明不足',
      3,
      1,
      '',
      '給与についてもう少し知りたい',
      '他社と比較検討中',
      '少し迷っている'
    ],

    // C003: 高評価維持
    [
      new Date('2025-10-12 14:30:00'),
      '鈴木一郎',
      'suzuki@example.com',
      10,
      '変わらない（元々高い）',
      10,
      'すべて',
      '特になし',
      '',
      'なし',
      'すべて解消',
      0,
      0,
      '',
      '完璧な面談だった',
      '非常に満足',
      'すぐにでも入社したい'
    ],

    // C004: 低評価のまま
    [
      new Date('2025-10-13 17:30:00'),
      '高橋次郎',
      'takahashi@example.com',
      5,
      '変わらない',
      5,
      '特になし',
      '業務内容が依然不明確',
      '具体的な業務フロー',
      '業務内容',
      '不安は残る',
      5,
      2,
      '',
      'まだ迷っている',
      '他社と比較中',
      '慎重に検討したい'
    ],

    // C005: やや上昇
    [
      new Date('2025-10-14 15:30:00'),
      '伊藤三郎',
      'ito@example.com',
      8,
      'やや上がった',
      8,
      '成長機会の具体性',
      '研修制度が明確になった',
      '',
      '特になし',
      '概ね解消',
      2,
      0,
      '',
      '良い印象',
      '前向きに検討',
      '次に進みたい'
    ]
  ];

  sheet.getRange(2, 1, responses.length, responses[0].length).setValues(responses);

  Logger.log(`✅ 社員面談アンケート: ${responses.length}件投入`);
}

/**
 * 2次面接アンケートデータの投入
 */
function insertSurvey2次面接() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('アンケート_2次面接');

  if (!sheet) {
    Logger.log('❌ アンケート_2次面接シートが見つかりません');
    return;
  }

  // 既存データをクリア
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  const responses = [
    // C001: 高評価
    [
      new Date('2025-10-20 16:00:00'),
      '田中太郎',
      'tanaka@example.com',
      9,
      '上がった',
      9,
      '代表の方のビジョンに共感',
      '特にない',
      '',
      '特になし',
      '入社を前向きに検討',
      9, // M: Q12. 承諾可能性（1-10）← 重要
      '',
      '非常に良い面接だった',
      'ビジョンに共感',
      '入社したい',
      '特になし',
      'ぜひ内定をいただきたい'
    ],

    // C002: やや改善
    [
      new Date('2025-10-21 17:00:00'),
      '佐藤花子',
      'sato@example.com',
      6,
      'やや上がった',
      7,
      '代表の熱意',
      '給与面は依然気になる',
      '昇給の仕組み',
      '給与',
      '慎重に検討',
      6, // 承諾可能性: やや低い
      '1ヶ月以上先',
      '給与面が気になる',
      '他社と比較中',
      '慎重に判断したい',
      '給与条件',
      '検討中'
    ],

    // C003: 最高評価
    [
      new Date('2025-10-22 15:30:00'),
      '鈴木一郎',
      'suzuki@example.com',
      10,
      '変わらない（高い）',
      10,
      'すべて',
      'なし',
      '',
      'なし',
      'すぐにでも入社したい',
      10, // 承諾可能性: 最高
      '',
      '完璧な面接',
      '非常に満足',
      '即答で承諾',
      'なし',
      '内定をお待ちしています'
    ],

    // C004: 低評価
    [
      new Date('2025-10-23 18:00:00'),
      '高橋次郎',
      'takahashi@example.com',
      5,
      '変わらない',
      5,
      '特になし',
      '業務内容が不明確',
      '具体的な業務',
      '業務内容',
      '迷っている',
      4, // 承諾可能性: 低い
      '2週間以内',
      '業務が不明確',
      '他社も検討',
      '迷っている',
      '業務内容',
      '慎重に判断'
    ],

    // C005: 標準
    [
      new Date('2025-10-24 16:30:00'),
      '伊藤三郎',
      'ito@example.com',
      8,
      'やや上がった',
      8,
      '成長機会',
      '研修制度',
      '',
      '特になし',
      '前向きに検討',
      7, // 承諾可能性: やや高い
      '',
      '良い面接',
      '成長できそう',
      '前向きに検討',
      '特になし',
      '検討します'
    ]
  ];

  sheet.getRange(2, 1, responses.length, responses[0].length).setValues(responses);

  Logger.log(`✅ 2次面接アンケート: ${responses.length}件投入`);
}

/**
 * 内定後アンケートデータの投入
 */
function insertSurvey内定後() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('アンケート_内定');

  if (!sheet) {
    Logger.log('❌ アンケート_内定シートが見つかりません');
    return;
  }

  // 既存データをクリア
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  const responses = [
    // C001: 高確率で承諾
    [
      new Date('2025-10-30 14:00:00'),
      '田中太郎',
      'tanaka@example.com',
      10, // D: Q3. 最終面接の満足度
      '非常に共感できた', // E: Q4. ビジョンへの共感
      '大きく上がった', // F: Q5. 働きたい気持ちの変化
      '特になし', // G: Q6. 不足しているスキル
      10, // H: Q7. PIGNUSで働くことへの前向きさ（1-10）← 重要
      '1週間以内', // I: Q8. 返答できるタイミング
      '2025-11-15', // J: Q9. 意思決定の期限
      0, // K: Q10. 他社の内定数
      0, // L: Q11. 選考中の企業数
      10, // M: Q12. PIGNUSの優位性（1-10）
      '成長機会・企業文化', // N: Q13. PIGNUSの魅力
      'ビジョンに共感・成長機会', // O: Q14. 入社を決める要因
      '給与条件の確認', // P: Q15. 入社を決断するために必要な情報
      '', // Q: Q16. 解消したい不安
      '特になし', // R: Q17. 他に懸念していること
      '非常に満足', // S: Q18. 選考プロセスの感想
      '非常に良かった' // T: Q19. 全体的な感想
    ],

    // C002: 慎重（他社と比較）
    [
      new Date('2025-10-31 15:00:00'),
      '佐藤花子',
      'sato@example.com',
      7,
      'やや共感できた',
      'やや上がった',
      '特になし',
      7, // 前向きさ: やや高い
      '2週間以内',
      '2025-11-20',
      1, // 他社内定あり
      2, // 選考中企業あり
      6, // PIGNUSの優位性: やや低い
      '勤務地',
      '待遇・成長機会',
      '給与詳細・昇給制度',
      '給与面での不安',
      '他社との比較',
      '概ね満足',
      '慎重に検討したい'
    ],

    // C003: 即承諾
    [
      new Date('2025-11-01 16:00:00'),
      '鈴木一郎',
      'suzuki@example.com',
      10,
      '非常に共感できた',
      '大きく上がった',
      'なし',
      10, // 前向きさ: 最高
      '即日',
      '2025-11-10',
      0,
      0,
      10, // PIGNUSの優位性: 最高
      'すべて',
      'ビジョン・文化・成長機会',
      'なし',
      '',
      'なし',
      '完璧',
      '即承諾します'
    ],

    // C004: 低確率
    [
      new Date('2025-11-02 17:00:00'),
      '高橋次郎',
      'takahashi@example.com',
      5,
      'あまり共感できなかった',
      '変わらない',
      '業務内容の理解',
      4, // 前向きさ: 低い
      '1ヶ月以内',
      '2025-11-30',
      2, // 他社内定2社
      3, // 選考中企業3社
      3, // PIGNUSの優位性: 低い
      '勤務地のみ',
      '待遇・安定性',
      '業務内容の詳細',
      '業務内容が不明確',
      '他社と大きく迷っている',
      '普通',
      '慎重に判断したい'
    ],

    // C005: やや高確率
    [
      new Date('2025-11-03 14:30:00'),
      '伊藤三郎',
      'ito@example.com',
      8,
      '共感できた',
      '上がった',
      '一部のスキル',
      8, // 前向きさ: 高い
      '1週間以内',
      '2025-11-15',
      0,
      1,
      8, // PIGNUSの優位性: 高い
      '成長機会',
      '成長機会・企業文化',
      '研修制度の詳細',
      '研修制度',
      '特になし',
      '良かった',
      '前向きに検討します'
    ]
  ];

  sheet.getRange(2, 1, responses.length, responses[0].length).setValues(responses);

  Logger.log(`✅ 内定後アンケート: ${responses.length}件投入`);
}

/**
 * 接点履歴の拡充（50件）
 */
function insertExpandedContactHistory() {
  try {
    Logger.log('\n========================================');
    Logger.log('接点履歴拡充データ投入開始');
    Logger.log('========================================\n');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONTACT_HISTORY);

    if (!sheet) {
      Logger.log('❌ Contact_Historyシートが見つかりません');
      return false;
    }

    // 既存データをクリア（ヘッダー以外）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }

    const contacts = [];
    let contactIdCounter = 1;

    // C001: 密な接点（10件）
    const c001Contacts = [
      ['2025-09-28', 'メール', '田中太郎'],
      ['2025-10-01', '電話', '田中太郎'],
      ['2025-10-02', '面談', '田中太郎'],
      ['2025-10-05', 'メール', '田中太郎'],
      ['2025-10-08', '面談', '田中太郎'],
      ['2025-10-10', '面談', '田中太郎'],
      ['2025-10-15', 'メール', '田中太郎'],
      ['2025-10-20', '面接', '田中太郎'],
      ['2025-10-25', 'メール', '田中太郎'],
      ['2025-10-30', '面接', '田中太郎']
    ];

    for (let contact of c001Contacts) {
      contacts.push([
        `CON${String(contactIdCounter).padStart(3, '0')}`,
        'C001',
        '田中太郎',
        new Date(contact[0]),
        contact[1],
        contact[2]
      ]);
      contactIdCounter++;
    }

    // C002: 標準的な接点（7件）
    const c002Contacts = [
      ['2025-10-01', 'メール', '佐藤花子'],
      ['2025-10-03', '面談', '佐藤花子'],
      ['2025-10-08', 'メール', '佐藤花子'],
      ['2025-10-11', '面談', '佐藤花子'],
      ['2025-10-17', 'メール', '佐藤花子'],
      ['2025-10-21', '面接', '佐藤花子'],
      ['2025-10-31', '面接', '佐藤花子']
    ];

    for (let contact of c002Contacts) {
      contacts.push([
        `CON${String(contactIdCounter).padStart(3, '0')}`,
        'C002',
        '佐藤花子',
        new Date(contact[0]),
        contact[1],
        contact[2]
      ]);
      contactIdCounter++;
    }

    // C003: 非常に密な接点（12件）
    const c003Contacts = [
      ['2025-09-30', 'メール', '鈴木一郎'],
      ['2025-10-02', '電話', '鈴木一郎'],
      ['2025-10-04', '面談', '鈴木一郎'],
      ['2025-10-06', 'メール', '鈴木一郎'],
      ['2025-10-08', '電話', '鈴木一郎'],
      ['2025-10-10', '面談', '鈴木一郎'],
      ['2025-10-12', '面談', '鈴木一郎'],
      ['2025-10-15', 'メール', '鈴木一郎'],
      ['2025-10-18', '電話', '鈴木一郎'],
      ['2025-10-22', '面接', '鈴木一郎'],
      ['2025-10-28', 'メール', '鈴木一郎'],
      ['2025-11-01', '面接', '鈴木一郎']
    ];

    for (let contact of c003Contacts) {
      contacts.push([
        `CON${String(contactIdCounter).padStart(3, '0')}`,
        'C003',
        '鈴木一郎',
        new Date(contact[0]),
        contact[1],
        contact[2]
      ]);
      contactIdCounter++;
    }

    // C004: 接点が少ない（4件）
    const c004Contacts = [
      ['2025-10-03', 'メール', '高橋次郎'],
      ['2025-10-05', '面談', '高橋次郎'],
      ['2025-10-13', '面談', '高橋次郎'],
      ['2025-10-23', '面接', '高橋次郎']
    ];

    for (let contact of c004Contacts) {
      contacts.push([
        `CON${String(contactIdCounter).padStart(3, '0')}`,
        'C004',
        '高橋次郎',
        new Date(contact[0]),
        contact[1],
        contact[2]
      ]);
      contactIdCounter++;
    }

    // C005: やや密な接点（8件）
    const c005Contacts = [
      ['2025-10-02', 'メール', '伊藤三郎'],
      ['2025-10-04', '電話', '伊藤三郎'],
      ['2025-10-06', '面談', '伊藤三郎'],
      ['2025-10-10', 'メール', '伊藤三郎'],
      ['2025-10-14', '面談', '伊藤三郎'],
      ['2025-10-18', 'メール', '伊藤三郎'],
      ['2025-10-24', '面接', '伊藤三郎'],
      ['2025-11-03', '面接', '伊藤三郎']
    ];

    for (let contact of c005Contacts) {
      contacts.push([
        `CON${String(contactIdCounter).padStart(3, '0')}`,
        'C005',
        '伊藤三郎',
        new Date(contact[0]),
        contact[1],
        contact[2]
      ]);
      contactIdCounter++;
    }

    // データを投入
    if (contacts.length > 0) {
      sheet.getRange(2, 1, contacts.length, 6).setValues(contacts);
      Logger.log(`✅ 接点履歴: ${contacts.length}件投入`);
    }

    Logger.log('\n========================================');
    Logger.log('接点履歴拡充データ投入完了');
    Logger.log('========================================\n');

    return true;

  } catch (error) {
    Logger.log(`❌ エラー: ${error}`);
    return false;
  }
}

/**
 * 全データ投入の実行
 */
function runExpandedDataInsertion() {
  Logger.log('\n========================================');
  Logger.log('Phase 3.5 全データ投入開始');
  Logger.log('========================================\n');

  // 1. Candidates_Master拡張データ
  Logger.log('>>> Step 1: Candidates_Master拡張データ投入');
  insertCandidatesMasterExtended();

  // 2. 全アンケートデータ
  Logger.log('\n>>> Step 2: 全アンケートデータ投入');
  insertExpandedSurveyResponses();

  // 3. 接点履歴拡充
  Logger.log('\n>>> Step 3: 接点履歴拡充');
  insertExpandedContactHistory();

  Logger.log('\n========================================');
  Logger.log('Phase 3.5 全データ投入完了');
  Logger.log('========================================');
  Logger.log('\n⚠️ 次のステップ:');
  Logger.log('1. runComprehensivePhase3Tests()を実行');
  Logger.log('2. 全候補者（C001-C005）×全フェーズ（4種類）= 20パターンのテスト');
  Logger.log('3. エラーゼロを確認');
  Logger.log('========================================\n');
}
