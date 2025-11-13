/**
 * 全初期データを投入
 */
function insertAllInitialData() {
  try {
    Logger.log('初期データの投入を開始します...');
    
    insertCompanyAssetsData();
    insertCompetitorDBData();
    insertSampleCandidateData();
    insertSampleEvaluationLogData();
    insertSampleEngagementLogData();
    insertSampleEvidenceData();
    insertSampleRiskData();
    insertSampleNextQData();
    insertSampleContactHistoryData();
    insertSampleSurveyResponseData();
    insertSampleAcceptanceStoryData();
    insertSampleCompetitorComparisonData();
    insertSampleArchiveData();
    
    Logger.log('✅ 全初期データの投入が完了しました');
    
  } catch (error) {
    logError('insertAllInitialData', error);
    throw error;
  }
}

/**
 * Company_Assetsの初期データを投入
 */
function insertCompanyAssetsData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.COMPANY_ASSETS);
  
  if (!sheet) return;
  
  const now = new Date();
  
  const data = [
    [
      'A001',
      '成長環境',
      '若手の裁量権が大きい',
      '入社3年で責任者になった事例あり。新規事業の立ち上げメンバーに抜擢される機会も多い。',
      '大手企業では下請け業務が中心だが、当社では入社1年目から責任者になれる環境がある。',
      '若手社員の成功事例を紹介する面談を設定。実際の責任者と話してもらう。',
      'https://example.com/career-path',
      now
    ],
    [
      'A002',
      '社風',
      'フラットな組織文化',
      '役職で呼ばない文化。失敗を許容し、チャレンジを推奨する雰囲気。週1回の全社MTGで誰でも発言可能。',
      '大手企業のような硬い雰囲気ではなく、自由に意見を言える環境。年齢や役職に関係なく、良いアイデアが採用される。',
      '社員面談で社風を体感してもらう。全社MTGの様子を動画で共有。',
      'https://example.com/culture',
      now
    ],
    [
      'A003',
      '事業安定性',
      '黒字経営（5期連続）',
      '売上成長率: 年平均20%。主要顧客との契約継続率95%以上。資金調達も完了済み。',
      'ベンチャーだが、安定した経営基盤がある。上場を視野に入れた経営体制。',
      '業績資料（IR資料）を提示。CFOとの面談を設定。',
      'https://example.com/ir',
      now
    ],
    [
      'A004',
      '給与待遇',
      '業界平均+10%の給与水準',
      '初任給: 新卒400万円、中途500万円〜。昇給実績: 入社3年で平均30%アップ。ストックオプションあり。',
      '初任給はやや低めだが、昇給率が高く、長期的には高給与。成果に応じたインセンティブ制度も充実。',
      '給与テーブルと昇給実績を提示。ストックオプションの詳細を説明。',
      'https://example.com/salary',
      now
    ],
    [
      'A005',
      '働き方',
      'リモートワーク可（週2日）',
      'フレックスタイム制。コアタイム: 11:00-15:00。残業時間: 月平均20時間。有給取得率80%以上。',
      '大手企業のような激務ではなく、ワークライフバランスが良い。家族との時間も大切にできる。',
      '働き方の実態を説明。実際の社員のスケジュールを共有。',
      'https://example.com/workstyle',
      now
    ],
    [
      'A006',
      '成長環境',
      '研修制度が充実',
      '入社時研修（2週間）、OJT、外部セミナー参加支援（年間10万円まで）、1on1（週1回）。',
      '大手企業並みの研修制度がある。自己成長を支援する文化。',
      '研修プログラムの詳細を説明。過去の研修実績を共有。',
      'https://example.com/training',
      now
    ],
    [
      'A007',
      '事業内容',
      '社会課題を解決する事業',
      '教育格差の解消、地方創生、環境問題など、社会的意義の高い事業を展開。',
      '単なる利益追求ではなく、社会に貢献できる仕事ができる。',
      '事業のミッション・ビジョンを説明。社会的インパクトを数値で示す。',
      'https://example.com/mission',
      now
    ],
    [
      'A008',
      '社風',
      '多様性を尊重',
      '女性管理職比率30%。外国籍社員10%。育休取得率（男性）50%。',
      '性別、国籍、年齢に関係なく活躍できる環境。ダイバーシティを重視。',
      'ダイバーシティの取り組みを説明。実際の女性管理職と面談。',
      'https://example.com/diversity',
      now
    ],
    [
      'A009',
      '福利厚生',
      '充実した福利厚生',
      '住宅手当（月3万円）、書籍購入補助（月1万円）、ランチ補助（月5千円）、部活動支援。',
      '給与以外の待遇も充実。生活コストを抑えられる。',
      '福利厚生の詳細を説明。実際の利用事例を共有。',
      'https://example.com/benefits',
      now
    ],
    [
      'A010',
      'キャリアパス',
      '複数のキャリアパスが選べる',
      'マネジメント職、スペシャリスト職、新規事業責任者など、多様なキャリアパスを用意。',
      '一つのキャリアに縛られない。自分の強みを活かせる道を選べる。',
      'キャリアパスの選択肢を説明。ロールモデルとなる社員を紹介。',
      'https://example.com/career',
      now
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ Company_Assetsに${data.length}件のデータを投入しました`);
}

/**
 * Competitor_DBの初期データを投入
 */
function insertCompetitorDBData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.COMPETITOR_DB);
  
  if (!sheet) return;
  
  const now = new Date();
  
  const data = [
    [
      'COMP001',
      'A社（大手コンサル）',
      'コンサルティング',
      '1000名以上',
      '知名度が高い',
      '給与水準が高い（業界平均+30%）',
      '研修制度が充実',
      '若手の裁量権が小さい',
      '残業時間が多い（月平均50時間）',
      '下請け業務が中心',
      JSON.stringify({
        description: '大手企業のため、研修制度は充実しているが、若手は下請け業務が中心。',
        strength: '体系的な研修プログラム',
        weakness: '実務経験が積みにくい'
      }),
      JSON.stringify({
        description: '業界平均+30%の高給与。ただし、激務のため時給換算すると低い。',
        base_salary: '新卒450万円、中途600万円〜',
        bonus: '年2回、業績連動'
      }),
      JSON.stringify({
        description: 'リモートワークは週1日のみ。残業時間が多く、ワークライフバランスは悪い。',
        remote_work: '週1日',
        overtime: '月平均50時間'
      }),
      JSON.stringify({
        description: '年功序列の文化が残る。若手の意見は通りにくい。',
        hierarchy: '階層的',
        communication: 'トップダウン'
      }),
      '自社は若手の裁量権が大きく、入社1年目から責任者になれる点で優位。給与は劣るが、ワークライフバランスは良い。',
      'https://www.a-company.com',
      now
    ],
    [
      'COMP002',
      'B社（メガベンチャー）',
      'IT/Web',
      '500名',
      '成長スピードが速い',
      'ストックオプションが魅力',
      '若手が活躍している',
      '給与水準が低い（業界平均-10%）',
      '離職率が高い（年間20%）',
      '激務（残業月40時間）',
      JSON.stringify({
        description: '成長スピードが速く、若手でも責任ある仕事を任される。ただし、激務。',
        strength: '若手の裁量権が大きい',
        weakness: '激務で離職率が高い'
      }),
      JSON.stringify({
        description: '給与は低めだが、ストックオプションが魅力。上場すれば大きなリターン。',
        base_salary: '新卒350万円、中途450万円〜',
        stock_option: 'あり（上場時に大きなリターン）'
      }),
      JSON.stringify({
        description: 'リモートワーク可（週2日）。ただし、残業時間は多め。',
        remote_work: '週2日',
        overtime: '月平均40時間'
      }),
      JSON.stringify({
        description: 'フラットな組織。ただし、競争が激しく、成果主義。',
        hierarchy: 'フラット',
        communication: 'オープン'
      }),
      '自社はB社と似た文化だが、残業時間が少なく、ワークライフバランスが良い点で優位。給与は同程度。',
      'https://www.b-company.com',
      now
    ],
    [
      'COMP003',
      'C社（外資系IT）',
      'IT/SaaS',
      '300名',
      '給与水準が高い（業界平均+50%）',
      'グローバルな環境',
      '最新技術に触れられる',
      '成果主義が厳しい',
      '雇用が不安定（リストラあり）',
      '日本法人は意思決定が遅い',
      JSON.stringify({
        description: 'グローバルな環境で最新技術に触れられる。ただし、成果主義が厳しい。',
        strength: '最新技術、グローバル',
        weakness: '成果主義、雇用不安定'
      }),
      JSON.stringify({
        description: '業界トップクラスの高給与。ただし、成果が出ないと減給やリストラのリスク。',
        base_salary: '新卒500万円、中途700万円〜',
        bonus: '年1回、業績連動（0-100%）'
      }),
      JSON.stringify({
        description: 'リモートワーク可（週3日）。フレックスタイム制。',
        remote_work: '週3日',
        overtime: '月平均30時間'
      }),
      JSON.stringify({
        description: 'グローバル基準の評価制度。成果主義が徹底。',
        hierarchy: 'フラット',
        communication: 'グローバル'
      }),
      '自社は給与では劣るが、雇用が安定しており、長期的なキャリア形成が可能な点で優位。',
      'https://www.c-company.com',
      now
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ Competitor_DBに${data.length}件のデータを投入しました`);
}

/**
 * Candidates_Masterのサンプルデータを投入
 */
function insertSampleCandidateData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.CANDIDATES_MASTER);
  
  if (!sheet) return;
  
  const now = new Date();
  const today = new Date();
  
  // 日付計算用のヘルパー関数
  function addDays(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }
  
  const data = [
    // 候補者1: 田中太郎（最終面接、承諾可能性高、CRITICAL）
    [
      'C001', '田中太郎', '最終面接', now, '新卒', '山田部長', addDays(today, -10),
      85, 80, 5, 0.85, 0.90, 0.85, 0.80, 0.85,
      75, 80, 77, 70, 7, '高', 85, 70, 80, '成長環境、裁量権', '給与水準がやや低い',
      'A社（大手コンサル）', 'B社（メガベンチャー）', '', '第二志望',
      addDays(today, -1), 1, 5, 3, today,
      '代表との1on1面談を設定し、成長環境の魅力を訴求', today, 'CRITICAL', '未実行', '山田部長',
      2.16, 3.0, 'tanaka@example.com',
      addDays(today, -60), '実施済', addDays(today, -55), '実施済',
      addDays(today, -50), '合格', 2, addDays(today, -45), '実施済',
      addDays(today, -35), '合格', addDays(today, -25), '合格'
    ],
    
    // 候補者2: 佐藤花子（2次面接、承諾可能性中、HIGH）
    [
      'C002', '佐藤花子', '2次面接', addDays(now, -1), '中途', '鈴木課長', addDays(today, -15),
      75, 70, 5, 0.75, 0.80, 0.75, 0.70, 0.75,
      55, 60, 57, 50, 7, '中', 65, 50, 60, '給与水準、福利厚生', '給与が低い、福利厚生が不明確',
      'C社（外資系IT）', '', '', '第三志望',
      addDays(today, -7), 7, 3, 5, addDays(today, 1),
      '給与テーブルと福利厚生の詳細を提示。人事との面談を設定', addDays(today, 1), 'HIGH', '未実行', '鈴木課長',
      1.28, 2.0, 'sato@example.com',
      addDays(today, -40), '実施済', addDays(today, -38), '実施済',
      addDays(today, -35), '合格', 1, addDays(today, -30), '実施済',
      addDays(today, -20), '合格', '', '未実施'
    ],
    
    // 候補者3: 鈴木一郎（最終面接、承諾可能性非常に高、NORMAL）
    [
      'C003', '鈴木一郎', '最終面接', addDays(now, -3), '新卒', '山田部長', addDays(today, -20),
      90, 85, 5, 0.90, 0.95, 0.90, 0.85, 0.90,
      85, 90, 87, 80, 7, '高', 95, 85, 90, '社会貢献、成長環境', '特になし',
      'B社（メガベンチャー）', '', '', '第一志望',
      addDays(today, -5), 5, 6, 3, addDays(today, 2),
      'フォローメールを送信し、入社意思を確認', addDays(today, 2), 'NORMAL', '未実行', '山田部長',
      2.61, 3.0, 'suzuki@example.com',
      addDays(today, -65), '実施済', addDays(today, -62), '実施済',
      addDays(today, -58), '合格', 3, addDays(today, -50), '実施済',
      addDays(today, -40), '合格', addDays(today, -28), '合格'
    ],
    
    // 候補者4: 高橋美咲（1次面接、承諾可能性低、HIGH）
    [
      'C004', '高橋美咲', '1次面接', addDays(now, -2), '新卒', '伊藤主任', addDays(today, -8),
      70, 65, 5, 0.70, 0.75, 0.70, 0.65, 0.70,
      40, 45, 42, 35, 7, '低', 50, 35, 45, '安定性', '事業の将来性に不安',
      'A社（大手コンサル）', 'C社（外資系IT）', '', '第四志望',
      addDays(today, -3), 3, 2, 4, addDays(today, 1),
      '事業計画と成長戦略を説明。CFOとの面談を設定', addDays(today, 1), 'HIGH', '未実行', '伊藤主任',
      0.88, 1.5, 'takahashi@example.com',
      addDays(today, -25), '実施済', addDays(today, -22), '実施済',
      addDays(today, -18), '合格', 0, '', '未実施',
      '', '未実施', '', '未実施'
    ],

    // 候補者5: 渡辺健太（初回面談、承諾可能性中、NORMAL）
    [
      'C005', '渡辺健太', '初回面談', addDays(now, -5), '中途', '鈴木課長', addDays(today, -5),
      65, 60, 5, 0.65, 0.70, 0.65, 0.60, 0.65,
      60, 55, 58, 55, 3, '中', 70, 55, 65, 'キャリアアップ', '転職理由が不明確',
      'B社（メガベンチャー）', '', '', '第二志望',
      addDays(today, -2), 2, 2, 2, addDays(today, 3),
      '1次面接を実施。キャリアビジョンを深掘り', addDays(today, 3), 'NORMAL', '未実行', '鈴木課長',
      0.62, 1.0, 'watanabe@example.com',
      addDays(today, -12), '実施済', '', '未実施',
      '', '未実施', 0, '', '未実施',
      '', '未実施', '', '未実施'
    ],
    
    // 候補者6: 山本さくら（最終面接、承諾可能性高、CRITICAL）
    [
      'C006', '山本さくら', '最終面接', now, '新卒', '山田部長', addDays(today, -12),
      80, 75, 5, 0.80, 0.85, 0.80, 0.75, 0.80,
      70, 75, 72, 65, 7, '高', 80, 65, 75, 'ワークライフバランス', '残業時間が気になる',
      'A社（大手コンサル）', '', '', '第二志望',
      addDays(today, -1), 1, 4, 3, today,
      '働き方の実態を説明。実際の社員と面談を設定', today, 'CRITICAL', '未実行', '山田部長',
      1.92, 3.0, 'yamamoto@example.com',
      addDays(today, -55), '実施済', addDays(today, -52), '実施済',
      addDays(today, -48), '合格', 1, addDays(today, -42), '実施済',
      addDays(today, -32), '合格', addDays(today, -22), '合格'
    ],

    // 候補者7: 中村大輔（2次面接、承諾可能性中、NORMAL）
    [
      'C007', '中村大輔', '2次面接', addDays(now, -4), '中途', '伊藤主任', addDays(today, -18),
      72, 68, 4, 0.72, 0.78, 0.72, 0.68, 0.72,
      58, 62, 60, 55, 5, '中', 68, 58, 62, '技術スキル向上', '技術スタックが合うか不安',
      'C社（外資系IT）', '', '', '第三志望',
      addDays(today, -6), 6, 3, 4, addDays(today, 2),
      '技術スタックの詳細を説明。CTOとの面談を設定', addDays(today, 2), 'NORMAL', '未実行', '伊藤主任',
      0.86, 2.0, 'nakamura@example.com',
      addDays(today, -45), '実施済', addDays(today, -42), '実施済',
      addDays(today, -38), '合格', 2, addDays(today, -33), '実施済',
      addDays(today, -24), '合格', '', '未実施'
    ],

    // 候補者8: 小林愛（1次面接、承諾可能性低、LOW）
    [
      'C008', '小林愛', '1次面接', addDays(now, -7), '新卒', '鈴木課長', addDays(today, -6),
      60, 58, 2, 0.60, 0.65, 0.60, 0.55, 0.60,
      35, 40, 37, 30, 7, '低', 45, 30, 40, '大手志向', '知名度が低い',
      'A社（大手コンサル）', 'B社（メガベンチャー）', 'C社（外資系IT）', '第五志望',
      addDays(today, -10), 10, 2, 5, addDays(today, 5),
      '企業ブランディング資料を提示。若手社員の成功事例を紹介', addDays(today, 5), 'LOW', '未実行', '鈴木課長',
      0.33, 1.5, 'kobayashi@example.com',
      addDays(today, -20), '実施済', addDays(today, -18), '実施済',
      addDays(today, -15), '合格', 0, '', '未実施',
      '', '未実施', '', '未実施'
    ],

    // 候補者9: 加藤翔太（初回面談、承諾可能性高、NORMAL）
    [
      'C009', '加藤翔太', '初回面談', addDays(now, -3), '新卒', '伊藤主任', addDays(today, -4),
      78, 75, 3, 0.78, 0.82, 0.78, 0.75, 0.78,
      72, 70, 71, 68, 3, '高', 78, 70, 75, '社会貢献、成長環境', '給与がやや低い',
      'B社（メガベンチャー）', '', '', '第一志望',
      addDays(today, -1), 1, 2, 2, addDays(today, 4),
      '1次面接を実施。給与テーブルと昇給実績を提示', addDays(today, 4), 'NORMAL', '未実行', '伊藤主任',
      1.08, 1.0, 'kato@example.com',
      addDays(today, -10), '実施済', '', '未実施',
      '', '未実施', 0, '', '未実施',
      '', '未実施', '', '未実施'
    ],

    // 候補者10: 伊藤美穂（2次面接、承諾可能性中、HIGH）
    [
      'C010', '伊藤美穂', '2次面接', addDays(now, -2), '中途', '山田部長', addDays(today, -14),
      68, 65, 3, 0.68, 0.72, 0.68, 0.65, 0.68,
      52, 55, 53, 48, 5, '中', 60, 50, 55, 'キャリアチェンジ', '未経験職種への不安',
      'A社（大手コンサル）', '', '', '第二志望',
      addDays(today, -4), 4, 3, 4, addDays(today, 1),
      '研修制度の詳細を説明。同じくキャリアチェンジした社員と面談', addDays(today, 1), 'HIGH', '未実行', '山田部長',
      0.90, 2.0, 'ito@example.com',
      addDays(today, -38), '実施済', addDays(today, -36), '実施済',
      addDays(today, -32), '合格', 1, addDays(today, -28), '実施済',
      addDays(today, -21), '合格', '', '未実施'
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ Candidates_Masterに${data.length}件のサンプルデータを投入しました`);
}

/**
 * Evaluation_Logのサンプルデータを投入
 */
function insertSampleEvaluationLogData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.EVALUATION_LOG);
  
  if (!sheet) return;
  
  const now = new Date();
  
  function addDays(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }
  
  const data = [
    // 田中太郎の評価（最終面接）
    [
      'EV001', 'C001', '田中太郎', addDays(now, -3), '最終面接', '山田部長',
      0.85, 0.90, '社会課題の解決に強い関心。NPOでのボランティア経験あり。',
      0.85, '課題を構造化し、戦略を立案できる。自分で意思決定した経験が豊富。',
      0.80, '当社を第一志望ではないが、成長環境に魅力を感じている。',
      0.85, '学生団体で新規プロジェクトを立ち上げ、100名規模のイベントを成功させた。',
      'Medium', 85, 'Pass', 'https://example.com/eval/EV001', 90
    ],

    // 佐藤花子の評価（2次面接）
    [
      'EV002', 'C002', '佐藤花子', addDays(now, -1), '2次面接', '鈴木課長',
      0.75, 0.80, '前職での経験から、顧客志向の重要性を理解している。',
      0.75, '戦略立案の経験はあるが、実行経験が少ない。',
      0.70, '給与水準が懸念。他社（C社）の最終面接も控えている。',
      0.75, '前職でプロジェクトマネージャーとして、5名のチームをリード。',
      'High', 75, 'Conditional', 'https://example.com/eval/EV002', 70
    ],

    // 鈴木一郎の評価（最終面接）
    [
      'EV003', 'C003', '鈴木一郎', addDays(now, -3), '最終面接', '山田部長',
      0.90, 0.95, '社会貢献への強い動機。教育格差の解消に関心。',
      0.90, '論理的思考力が高く、複雑な問題を構造化できる。',
      0.85, '当社を第一志望。他社選考は辞退する意向。',
      0.90, 'インターンで新規事業の立ち上げに貢献。売上100万円を達成。',
      'None', 90, 'Pass', 'https://example.com/eval/EV003', 95
    ],

    // 高橋美咲の評価（1次面接）
    [
      'EV004', 'C004', '高橋美咲', addDays(now, -2), '1次面接', '伊藤主任',
      0.70, 0.75, '安定志向が強い。リスクを取ることに消極的。',
      0.70, '分析力はあるが、自分で意思決定した経験が少ない。',
      0.65, '大手企業志向。当社の事業の将来性に不安を感じている。',
      0.70, 'ゼミでのリサーチプロジェクトで、データ分析を担当。',
      'High', 70, 'Conditional', 'https://example.com/eval/EV004', 60
    ],
    
    // 渡辺健太の評価（初回面談）
    [
      'EV005', 'C005', '渡辺健太', addDays(now, -5), '初回面談', '鈴木課長',
      0.65, 0.70, 'キャリアアップへの意欲はあるが、具体的なビジョンが不明確。',
      0.65, '前職での戦略立案経験はあるが、深さが不足。',
      0.60, '転職理由が曖昧。複数社を並行して選考中。',
      0.65, '前職で営業として、年間売上3000万円を達成。',
      'Medium', 65, 'Conditional', 'https://example.com/eval/EV005', 75
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ Evaluation_Logに${data.length}件のサンプルデータを投入しました`);
}

/**
 * Engagement_Logのサンプルデータを投入
 */
function insertSampleEngagementLogData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.ENGAGEMENT_LOG);
  
  if (!sheet) return;
  
  const now = new Date();
  const today = new Date();
  
  function addDays(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }
  
  const data = [
    // 田中太郎の承諾可能性分析
    [
      'EN001', 'C001', '田中太郎', addDays(now, -1), '面接',
      76, 75, 77, '高', 85, 70, 80,
      '成長環境、裁量権', '給与水準がやや低い', '給与・待遇',
      'A社（大手コンサル）、B社（メガベンチャー）', 70,
      '代表との1on1面談を設定し、成長環境の魅力を訴求', today, '高',
      'https://example.com/engagement/EN001'
    ],
    
    // 佐藤花子の承諾可能性分析
    [
      'EN002', 'C002', '佐藤花子', addDays(now, -1), '面接',
      57, 55, 57, '中', 65, 50, 60,
      '給与水準、福利厚生', '給与が低い、福利厚生が不明確', '給与・待遇',
      'C社（外資系IT）', 50,
      '給与テーブルと福利厚生の詳細を提示。人事との面談を設定', addDays(now, 1), '高',
      'https://example.com/engagement/EN002'
    ],
    
    // 鈴木一郎の承諾可能性分析
    [
      'EN003', 'C003', '鈴木一郎', addDays(now, -3), '面接',
      87, 85, 87, '高', 95, 85, 90,
      '社会貢献、成長環境', '特になし', 'なし',
      'B社（メガベンチャー）', 85,
      'フォローメールを送信し、入社意思を確認', addDays(now, 2), '中',
      'https://example.com/engagement/EN003'
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ Engagement_Logに${data.length}件のサンプルデータを投入しました`);
}

/**
 * Evidenceのサンプルデータを投入
 */
function insertSampleEvidenceData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.EVIDENCE);
  
  if (!sheet) return;
  
  const now = new Date();
  
  const data = [
    [
      'EVID001', 'C001', '田中太郎', 'EV001', 'Philosophy',
      '「教育格差の問題に取り組みたい。NPOでのボランティアを通じて、地方の子どもたちの可能性を広げる仕事がしたいと思いました」',
      now
    ],
    [
      'EVID002', 'C001', '田中太郎', 'EV001', 'Strategy',
      '「学生団体で新規イベントを企画する際、まず課題を3つに分解し、それぞれに対する施策を立案しました。結果として100名規模のイベントを成功させることができました」',
      now
    ],
    [
      'EVID003', 'C002', '佐藤花子', 'EV002', 'Motivation',
      '「C社の最終面接も控えていますが、御社の成長環境には魅力を感じています。ただ、給与面での不安があります」',
      now
    ],
    [
      'EVID004', 'C003', '鈴木一郎', 'EV003', 'Philosophy',
      '「社会課題を解決する事業に携わりたい。特に教育格差の問題は、自分が地方出身ということもあり、強い関心があります」',
      now
    ],
    [
      'EVID005', 'C003', '鈴木一郎', 'EV003', 'Execution',
      '「インターンで新規事業の立ち上げに参加し、3ヶ月で売上100万円を達成しました。顧客ヒアリングから施策立案、実行まで一貫して担当しました」',
      now
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ Evidenceに${data.length}件のサンプルデータを投入しました`);
}

/**
 * Riskのサンプルデータを投入
 */
function insertSampleRiskData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.RISK);
  
  if (!sheet) return;
  
  const now = new Date();
  
  const data = [
    [
      'RISK001', 'C001', '田中太郎', 'EV001', 'Medium',
      '給与水準が競合企業（A社、B社）より低い',
      '給与テーブルと昇給実績を提示。長期的な給与水準の高さを訴求',
      now
    ],
    [
      'RISK002', 'C002', '佐藤花子', 'EV002', 'High',
      'C社（外資系IT）の最終面接が控えている。給与水準が当社より50%高い',
      '福利厚生と働き方の良さを訴求。C社の雇用不安定性をリスクとして説明',
      now
    ],
    [
      'RISK003', 'C004', '高橋美咲', 'EV004', 'High',
      '大手企業志向が強く、当社の知名度の低さを懸念',
      '企業ブランディング資料を提示。若手の成長事例を紹介',
      now
    ],
    [
      'RISK004', 'C005', '渡辺健太', 'EV005', 'Medium',
      '転職理由が曖昧。複数社を並行して選考中',
      'キャリアビジョンを深掘りし、当社でのキャリアパスを提示',
      now
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ Riskに${data.length}件のサンプルデータを投入しました`);
}

/**
 * NextQのサンプルデータを投入
 */
function insertSampleNextQData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.NEXT_Q);
  
  if (!sheet) return;
  
  const now = new Date();
  
  const data = [
    [
      'Q001', 'C001', '田中太郎', 'EV001', 'キャリアビジョン',
      '5年後、10年後のキャリアビジョンを教えてください',
      '長期的なキャリアビジョンが当社とマッチするか確認',
      now
    ],
    [
      'Q002', 'C002', '佐藤花子', 'EV002', '給与・待遇',
      '給与面での希望を教えてください。C社との比較でどう考えていますか？',
      '給与への期待値と、C社との比較を明確化',
      now
    ],
    [
      'Q003', 'C004', '高橋美咲', 'EV004', 'モチベーション',
      '大手企業と当社の違いをどう考えていますか？',
      '大手企業志向の強さと、当社への関心度を確認',
      now
    ],
    [
      'Q004', 'C005', '渡辺健太', 'EV005', '転職理由',
      '転職を考えた理由を詳しく教えてください',
      '転職理由を深掘りし、当社でそれが解決できるか確認',
      now
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ NextQに${data.length}件のサンプルデータを投入しました`);
}

/**
 * Contact_Historyのサンプルデータを投入
 */
function insertSampleContactHistoryData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.CONTACT_HISTORY);
  
  if (!sheet) return;
  
  const now = new Date();
  const today = new Date();
  
  function addDays(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }
  
  const data = [
    [
      'CONT001', 'C001', '田中太郎', addDays(now, -10), '面談', '山田部長',
      '初回面談。経歴と志望動機をヒアリング', addDays(now, -7)
    ],
    [
      'CONT002', 'C001', '田中太郎', addDays(now, -7), '1次面接', '鈴木課長',
      '1次面接。4軸評価を実施', addDays(now, -5)
    ],
    [
      'CONT003', 'C001', '田中太郎', addDays(now, -5), '2次面接', '伊藤主任',
      '2次面接。技術スキルと思考様式を評価', addDays(now, -3)
    ],
    [
      'CONT004', 'C001', '田中太郎', addDays(now, -3), '最終面接', '山田部長',
      '最終面接。カルチャーフィットを確認', today
    ],
    [
      'CONT005', 'C001', '田中太郎', addDays(now, -1), 'メール', '山田部長',
      'フォローメール送信。次回アクションを提案', today
    ],
    [
      'CONT006', 'C002', '佐藤花子', addDays(now, -15), '面談', '鈴木課長',
      '初回面談。経歴と転職理由をヒアリング', addDays(now, -10)
    ],
    [
      'CONT007', 'C002', '佐藤花子', addDays(now, -10), '1次面接', '伊藤主任',
      '1次面接。前職での実績を確認', addDays(now, -7)
    ],
    [
      'CONT008', 'C002', '佐藤花子', addDays(now, -7), 'メール', '鈴木課長',
      'フォローメール送信。2次面接の日程調整', addDays(now, -1)
    ],
    [
      'CONT009', 'C002', '佐藤花子', addDays(now, -1), '2次面接', '鈴木課長',
      '2次面接。給与面での懸念をヒアリング', addDays(now, 1)
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ Contact_Historyに${data.length}件のサンプルデータを投入しました`);
}

/**
 * Survey_Responseのサンプルデータを投入
 */
function insertSampleSurveyResponseData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);

  if (!sheet) return;

  const now = new Date();

  function addDays(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }

  const data = [
    [
      'SURV001', 'C001', '田中太郎', addDays(now, -2), 8,
      '給与水準がやや低い', 'A社（大手コンサル）、B社（メガベンチャー）の最終面接を控えています',
      '成長環境に魅力を感じています', '初回面談' // ✅ I列: アンケート種別を追加
    ],
    [
      'SURV002', 'C002', '佐藤花子', addDays(now, -3), 6,
      '給与が低い、福利厚生が不明確', 'C社（外資系IT）の最終面接が来週です',
      '働き方には魅力を感じていますが、給与面での不安があります', '社員面談' // ✅ I列: アンケート種別を追加
    ],
    [
      'SURV003', 'C003', '鈴木一郎', addDays(now, -4), 9,
      '特になし', 'B社（メガベンチャー）の選考は辞退しました',
      '御社が第一志望です。早く内定をいただきたいです', '2次面接' // ✅ I列: アンケート種別を追加
    ],
    [
      'SURV004', 'C004', '高橋美咲', addDays(now, -5), 5,
      '事業の将来性に不安', 'A社（大手コンサル）、C社（外資系IT）の選考中です',
      '安定性を重視しています', '内定後' // ✅ I列: アンケート種別を追加
    ]
  ];

  // ✅ 9列のデータを書き込む（8列 → 9列に修正）
  sheet.getRange(2, 1, data.length, 9).setValues(data);

  Logger.log(`✅ Survey_Responseに${data.length}件のサンプルデータを投入しました`);
}

/**
 * Acceptance_Storyのサンプルデータを投入
 */
function insertSampleAcceptanceStoryData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.ACCEPTANCE_STORY);
  
  if (!sheet) return;
  
  const now = new Date();
  
  const data = [
    [
      'C001', '田中太郎',
      '現在、A社とB社の最終面接を控えている。当社への志望度は高いが、給与面での懸念がある。競合企業との比較で、成長環境では優位だが、給与では劣位。',
      '【ステップ1】代表との1on1面談で成長環境の魅力を訴求 → 【ステップ2】給与テーブルと昇給実績を提示 → 【ステップ3】内定を出し、入社意思を確認',
      '代表との1on1面談を設定し、成長環境の魅力を訴求',
      '当社の成長環境への理解が深まり、志望度が85%→90%に向上',
      '「入社1年目から責任者になれる」という発言に強い関心を示していた',
      '給与テーブルと昇給実績を提示。長期的な給与水準の高さを説明',
      '給与面での懸念が解消され、承諾可能性が77%→85%に向上',
      '「長期的には高給与」という説明に納得していた',
      '内定を出し、入社意思を確認。A社・B社の選考状況もヒアリング',
      '内定承諾。A社・B社の選考は辞退',
      '「成長環境が決め手」という発言',
      'A社・B社から先に内定が出た場合、当社を選ばないリスクがある',
      '【優先度1】代表との1on1面談（緊急度: CRITICAL）\n【優先度2】給与説明（緊急度: HIGH）\n【優先度3】内定出し（緊急度: HIGH）',
      now, now
    ],
    [
      'C002', '佐藤花子',
      '現在、C社の最終面接を控えている。給与面での懸念が大きく、C社の給与水準が当社より50%高い。当社への志望度は中程度。',
      '【ステップ1】給与テーブルと福利厚生の詳細を提示 → 【ステップ2】C社の雇用不安定性をリスクとして説明 → 【ステップ3】人事との面談で働き方の良さを訴求',
      '給与テーブルと福利厚生の詳細を提示',
      '給与面での懸念が部分的に解消され、承諾可能性が57%→65%に向上',
      '「福利厚生が充実している」という説明に関心を示していた',
      'C社の雇用不安定性（リストラのリスク）を説明',
      'C社へのリスク認識が高まり、当社への志望度が65%→70%に向上',
      '「安定性も重要」という発言',
      '人事との面談で働き方の良さを訴求。ワークライフバランスを重視',
      '内定承諾の可能性が高まる',
      '「働き方が魅力的」という発言',
      'C社から先に内定が出た場合、給与の高さで選ばれるリスクがある',
      '【優先度1】給与・福利厚生の説明（緊急度: HIGH）\n【優先度2】C社のリスク説明（緊急度: HIGH）\n【優先度3】人事面談（緊急度: NORMAL）',
      now, now
    ],
    [
      'C003', '鈴木一郎',
      '当社が第一志望。B社の選考は辞退済み。志望度が非常に高く、承諾可能性は87%。',
      '【ステップ1】フォローメールで入社意思を確認 → 【ステップ2】内定を出す → 【ステップ3】入社までのフォロー',
      'フォローメールを送信し、入社意思を確認',
      '入社意思が確認でき、承諾可能性が87%→95%に向上',
      '「早く内定をいただきたい」という発言',
      '内定を出す',
      '即座に承諾',
      '「第一志望です」という発言',
      '入社までのフォロー。定期的に連絡を取り、不安を解消',
      '入社まで円滑に進む',
      '「入社が楽しみ」という発言',
      '特になし。承諾可能性が非常に高い',
      '【優先度1】フォローメール（緊急度: NORMAL）\n【優先度2】内定出し（緊急度: NORMAL）\n【優先度3】入社フォロー（緊急度: LOW）',
      now, now
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ Acceptance_Storyに${data.length}件のサンプルデータを投入しました`);
}

/**
 * Competitor_Comparisonのサンプルデータを投入
 */
function insertSampleCompetitorComparisonData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.COMPETITOR_COMPARISON);
  
  if (!sheet) return;
  
  const now = new Date();
  
  const data = [
    [
      'COMP_C001_A社', 'C001', '田中太郎', 'A社（大手コンサル）',
      '中程度（成長中）', '高い（大手）', '競合',
      '業界平均+10%', '業界平均+30%', '競合',
      '若手の裁量権が大きい', '若手は下請け業務が中心', '自社',
      '入社1年目から責任者', '若手の裁量権は小さい', '自社',
      '月平均20時間', '月平均50時間', '自社',
      'フラットな組織', '年功序列', '自社',
      '黒字経営（5期連続）', '大手企業で安定', '互角',
      '自社', '成長環境と裁量権の大きさを訴求。給与は劣るが、ワークライフバランスが良い点を強調',
      now, now
    ],
    [
      'COMP_C001_B社', 'C001', '田中太郎', 'B社（メガベンチャー）',
      '中程度（成長中）', '高い（メガベンチャー）', '競合',
      '業界平均+10%', '業界平均-10%（ストックオプションあり）', '自社',
      '若手の裁量権が大きい', '若手の裁量権が大きい', '互角',
      '入社1年目から責任者', '入社1年目から責任者', '互角',
      '月平均20時間', '月平均40時間', '自社',
      'フラットな組織', 'フラットな組織', '互角',
      '黒字経営（5期連続）', '成長中（上場準備中）', '自社',
      '自社', 'ワークライフバランスの良さを訴求。B社は激務で離職率が高い点を説明',
      now, now
    ],
    [
      'COMP_C002_C社', 'C002', '佐藤花子', 'C社（外資系IT）',
      '中程度（成長中）', '高い（外資系）', '競合',
      '業界平均+10%', '業界平均+50%', '競合',
      '若手の裁量権が大きい', 'グローバル基準', '互角',
      '入社1年目から責任者', '成果主義が厳しい', '自社',
      '月平均20時間', '月平均30時間', '自社',
      'フラットな組織', 'グローバル', '互角',
      '黒字経営（5期連続）', '雇用が不安定（リストラあり）', '自社',
      '自社', '雇用の安定性とワークライフバランスを訴求。C社は成果主義が厳しく、リストラのリスクがある点を説明',
      now, now
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  Logger.log(`✅ Competitor_Comparisonに${data.length}件のサンプルデータを投入しました`);
}

/**
 * Archiveのサンプルデータを投入
 */
function insertSampleArchiveData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.ARCHIVE);
  
  if (!sheet) return;
  
  const now = new Date();
  
  function addDays(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }
  
  const data = [
    [
      'C999', '山田花子', addDays(now, -60), addDays(now, -30), 'マーケティング部',
      92, 88,
      '成長環境と社風に魅力を感じて承諾。代表との面談が決め手となった。',
      addDays(now, -60)
    ],
    [
      'C998', '佐々木太郎', addDays(now, -90), addDays(now, -60), '営業部',
      88, 85,
      '給与面での懸念があったが、昇給実績の説明で納得。若手の裁量権の大きさが魅力。',
      addDays(now, -90)
    ]
  ];
  
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

  Logger.log(`✅ Archiveに${data.length}件のサンプルデータを投入しました`);
}

/**
 * テストデータのバリデーション
 * ヘッダー数とデータ列数の一致をチェック
 *
 * @return {boolean} バリデーション成功/失敗
 */
function validateTestData() {
  try {
    Logger.log('\n========================================');
    Logger.log('テストデータバリデーション開始');
    Logger.log('========================================');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const results = [];

    // Survey_Responseのチェック
    const surveyResponseSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_RESPONSE);
    if (surveyResponseSheet) {
      const headers = surveyResponseSheet.getRange(1, 1, 1, surveyResponseSheet.getLastColumn()).getValues()[0];
      const dataRow = surveyResponseSheet.getRange(2, 1, 1, surveyResponseSheet.getLastColumn()).getValues()[0];

      const headerCount = headers.filter(h => h !== '').length;
      const dataCount = dataRow.filter(d => d !== '').length;

      Logger.log(`\n【Survey_Response】`);
      Logger.log(`  ヘッダー数: ${headerCount}列`);
      Logger.log(`  データ列数: ${dataCount}列`);
      Logger.log(`  ヘッダー: ${headers.filter(h => h !== '').join(', ')}`);

      if (headerCount !== 9) {
        results.push(`❌ Survey_Response: ヘッダー数が不正 (期待: 9列, 実際: ${headerCount}列)`);
        Logger.log(`  ❌ ヘッダー数が不正`);
      } else {
        results.push(`✅ Survey_Response: ヘッダー数が正しい (9列)`);
        Logger.log(`  ✅ ヘッダー数が正しい`);
      }

      if (dataCount !== 9) {
        results.push(`❌ Survey_Response: データ列数が不正 (期待: 9列, 実際: ${dataCount}列)`);
        Logger.log(`  ❌ データ列数が不正`);
      } else {
        results.push(`✅ Survey_Response: データ列数が正しい (9列)`);
        Logger.log(`  ✅ データ列数が正しい`);
      }
    } else {
      results.push(`❌ Survey_Response: シートが見つかりません`);
    }

    // Survey_Send_Logのチェック
    const sendLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SURVEY_SEND_LOG);
    if (sendLogSheet) {
      const headers = sendLogSheet.getRange(1, 1, 1, sendLogSheet.getLastColumn()).getValues()[0];
      const dataRow2 = sendLogSheet.getRange(2, 1, 1, sendLogSheet.getLastColumn()).getValues()[0];

      const headerCount = headers.filter(h => h !== '').length;
      const dataCount = dataRow2.filter(d => d !== '').length;

      Logger.log(`\n【Survey_Send_Log】`);
      Logger.log(`  ヘッダー数: ${headerCount}列`);
      Logger.log(`  データ列数: ${dataCount}列`);
      Logger.log(`  ヘッダー: ${headers.filter(h => h !== '').join(', ')}`);

      if (headerCount !== 8) {
        results.push(`❌ Survey_Send_Log: ヘッダー数が不正 (期待: 8列, 実際: ${headerCount}列)`);
        Logger.log(`  ❌ ヘッダー数が不正`);
      } else {
        results.push(`✅ Survey_Send_Log: ヘッダー数が正しい (8列)`);
        Logger.log(`  ✅ ヘッダー数が正しい`);
      }

      if (dataCount !== 8) {
        results.push(`❌ Survey_Send_Log: データ列数が不正 (期待: 8列, 実際: ${dataCount}列)`);
        Logger.log(`  ❌ データ列数が不正`);
      } else {
        results.push(`✅ Survey_Send_Log: データ列数が正しい (8列)`);
        Logger.log(`  ✅ データ列数が正しい`);
      }
    } else {
      results.push(`❌ Survey_Send_Log: シートが見つかりません`);
    }

    Logger.log('\n========================================');
    Logger.log('バリデーション結果');
    Logger.log('========================================');
    results.forEach(r => Logger.log(r));
    Logger.log('========================================\n');

    // エラーがあればアラート表示
    const errors = results.filter(r => r.startsWith('❌'));
    if (errors.length > 0) {
      const message = '【⚠️ テストデータにエラーがあります】\n\n' +
        errors.join('\n') +
        '\n\nInitialData.gsまたはTestDataGenerator.gsを修正してください。';

      SpreadsheetApp.getUi().alert(
        'テストデータ検証エラー',
        message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }

    const message = '【✅ テストデータが正常です】\n\n' +
      results.join('\n') +
      '\n\n全てのバリデーションチェックに合格しました。';

    SpreadsheetApp.getUi().alert(
      'テストデータ検証成功',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return true;

  } catch (error) {
    Logger.log(`❌ validateTestDataエラー: ${error.message}`);
    Logger.log(error.stack);

    SpreadsheetApp.getUi().alert(
      'エラー',
      `バリデーション中にエラーが発生しました:\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
}
