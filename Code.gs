/**
 * 案件2: GAS + AI ブログ記事自動生成
 *
 * スプレッドシートにキーワードを入力 → Claude APIで記事下書きを生成
 * → WordPressに「下書き」として投稿、という一連の流れを自動化する。
 *
 * 【安全のための大原則】
 * 生成した記事は必ず status: 'draft' でWordPressに送る。
 * 自動公開(publish)は絶対にしない。クライアントとの約束(下書き→人間確認→公開)を
 * コードの中で強制することで、ヒアリングで決めた運用ルールを担保している。
 */

// ============================================================
// 設定値(列の位置や固定のステータス文字列など)
// ============================================================

const DATA_SHEET_NAME = 'キーワード管理';
const CONFIG_SHEET_NAME = '設定';

// データシートの列(1始まり)
// 今回の課題は「実投稿はせず、記事本文+メタディスクリプションをスプレッドシートに
// 書き出すところまで」で完了とする方針のため、本文(CONTENT)列をタイトルの次に追加している。
const COL = {
  KEYWORD: 1, // A: キーワード
  SUB_KEYWORD: 2, // B: サブキーワード
  STATUS: 3, // C: ステータス
  TITLE: 4, // D: タイトル
  CONTENT: 5, // E: 本文(HTML)
  WP_URL: 6, // F: WordPress URL(実投稿した場合のみ。今回はログ確認のみなので基本空)
  META_DESCRIPTION: 7, // G: メタディスクリプション
  UPDATED_AT: 8, // H: 最終更新日時
  ERROR: 9, // I: エラー内容(失敗時のみ)
};

// true: 実際にはWordPressへ投稿せず、送信するはずのペイロードをLogger.logで確認するだけにする
// (課題の「推奨・誰でも完了可」ルート)。実際に投稿まで試す場合は false にする。
const SKIP_REAL_WORDPRESS_POST = true;

const STATUS = {
  NOT_STARTED: '未生成',
  GENERATING: '生成中',
  GENERATED: '生成済み',
  POSTING: '投稿中',
  POSTED: '投稿済み',
  ERROR: 'エラー',
};

// 設定シートのセル位置(ラベルはB列、値はC列に入力してもらう想定)
const CONFIG_CELL = {
  MODEL: 'C2', // 省略可。空ならデフォルトモデルを使う
  WP_BASE_URL: 'C3', // 例: https://example-saas.com
  WP_USERNAME: 'C4', // WordPressのユーザー名
  TARGET_READER: 'C5', // ターゲット読者
  LENGTH_GUIDE: 'C6', // 文字数の目安(例: 2000〜3000字)
  SAMPLE_ARTICLE: 'C7', // クライアントの「理想の記事」をそのまま貼り付ける
};

const DEFAULT_MODEL = 'claude-sonnet-5';
const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';

// Script Propertiesに保存するキー名(機密情報はシートに書かない)
const PROP_KEYS = {
  ANTHROPIC_API_KEY: 'ANTHROPIC_API_KEY',
  WP_APP_PASSWORD: 'WP_APP_PASSWORD',
};

// ============================================================
// メニュー
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('記事生成')
    .addItem('選択した行で記事を生成', 'generateForSelectedRows')
    .addSeparator()
    .addItem('初期設定(APIキーなど)', 'showSetupDialog')
    .addItem('設定シートを初期化', 'initConfigSheet')
    .addToUi();
}

/**
 * 初回セットアップ: Anthropic APIキーとWordPressのアプリケーションパスワードを
 * Script Propertiesに保存する。機密情報なのでシートのセルには書かない。
 */
function showSetupDialog() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const apiKeyResponse = ui.prompt(
    '初期設定 (1/2)',
    'Anthropic APIキーを入力してください(console.anthropic.comで発行したもの)。\n' +
      '未入力のまま次に進むと、保存済みの値を変更しません。',
    ui.ButtonSet.OK_CANCEL
  );
  if (apiKeyResponse.getSelectedButton() !== ui.Button.OK) return;
  const apiKey = apiKeyResponse.getResponseText().trim();
  if (apiKey) props.setProperty(PROP_KEYS.ANTHROPIC_API_KEY, apiKey);

  const wpPasswordResponse = ui.prompt(
    '初期設定 (2/2)',
    'WordPressのアプリケーションパスワードを入力してください。\n' +
      '(WordPress管理画面 → ユーザー → プロフィール → アプリケーションパスワード で発行)\n' +
      '未入力のまま次に進むと、保存済みの値を変更しません。',
    ui.ButtonSet.OK_CANCEL
  );
  if (wpPasswordResponse.getSelectedButton() !== ui.Button.OK) return;
  const wpPassword = wpPasswordResponse.getResponseText().trim();
  if (wpPassword) props.setProperty(PROP_KEYS.WP_APP_PASSWORD, wpPassword);

  ui.alert(
    '保存しました。WordPressのURL・ユーザー名・サンプル記事などは「設定」シートに直接入力してください。'
  );
}

/**
 * 「設定」シートと「キーワード管理」シートの雛形を作る。
 * まだシートがない場合に実行すると、列見出しや入力欄のラベルを用意してくれる。
 */
function initConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) dataSheet = ss.insertSheet(DATA_SHEET_NAME);
  const headers = [
    'キーワード',
    'サブキーワード',
    'ステータス',
    'タイトル',
    '本文',
    'WordPress URL',
    'メタディスクリプション',
    '最終更新日時',
    'エラー内容',
  ];
  dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  dataSheet.setFrozenRows(1);

  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
  const configRows = [
    ['項目', '値(ここに入力)'],
    ['Anthropicモデル(空欄可)', ''],
    ['WordPress URL(例: https://example.com)', ''],
    ['WordPressユーザー名', ''],
    ['ターゲット読者', ''],
    ['文字数の目安(例: 2000〜3000字)', '2000〜3000字'],
    ['サンプル記事(クライアントの理想の記事をそのまま貼り付け)', ''],
  ];
  configSheet.getRange(1, 1, configRows.length, 2).setValues(configRows);
  configSheet.setColumnWidth(2, 500);

  SpreadsheetApp.getUi().alert(
    '雛形を作成しました。「設定」シートに値を入力し、「記事生成」メニューの「初期設定」でAPIキーを登録してください。'
  );
}

// ============================================================
// メイン処理: 選択した行で記事生成〜WordPress投稿
// ============================================================

function generateForSelectedRows() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  if (!sheet) {
    ui.alert('「' + DATA_SHEET_NAME + '」シートが見つかりません。先に「設定シートを初期化」を実行してください。');
    return;
  }

  const selection = sheet.getActiveRange();
  if (!selection) {
    ui.alert('記事を生成したい行を選択してから実行してください。');
    return;
  }

  const startRow = Math.max(selection.getRow(), 2); // 1行目はヘッダー
  const endRow = startRow + selection.getNumRows() - 1;

  // GASの実行時間は6分までなので、一度に大量の行を処理しようとしない。
  // 複数行まとめて選んだ場合でも、安全のため上限を設ける。
  const MAX_ROWS_PER_RUN = 5;
  if (endRow - startRow + 1 > MAX_ROWS_PER_RUN) {
    ui.alert(
      '一度に処理できるのは最大' + MAX_ROWS_PER_RUN + '行までです(GASの実行時間制限のため)。\n' +
        '選択範囲を狭くして再実行してください。'
    );
    return;
  }

  for (let row = startRow; row <= endRow; row++) {
    const keyword = sheet.getRange(row, COL.KEYWORD).getValue();
    if (!keyword) continue; // キーワードが空の行はスキップ
    try {
      generateArticleForRow(sheet, row);
    } catch (err) {
      updateStatus(sheet, row, STATUS.ERROR, String(err));
    }
    // APIのレート制限に配慮して、行の間に少し間隔を空ける
    Utilities.sleep(1000);
  }

  ui.alert('処理が完了しました。ステータス列を確認してください。');
}

function generateArticleForRow(sheet, row) {
  const keyword = sheet.getRange(row, COL.KEYWORD).getValue();
  const subKeyword = sheet.getRange(row, COL.SUB_KEYWORD).getValue();

  updateStatus(sheet, row, STATUS.GENERATING);

  const config = readConfig();
  const article = callClaudeForArticle(keyword, subKeyword, config);

  // 本文・タイトル・メタディスクリプションをスプレッドシートに書き出す
  sheet.getRange(row, COL.TITLE).setValue(article.title);
  sheet.getRange(row, COL.CONTENT).setValue(article.contentHtml);
  sheet.getRange(row, COL.META_DESCRIPTION).setValue(article.metaDescription);

  // WordPressに送るはずのペイロードを組み立てる(実際に送るかどうかはSKIP_REAL_WORDPRESS_POSTで切り替え)
  const payload = buildWordPressPayload(article);

  if (SKIP_REAL_WORDPRESS_POST) {
    // 実投稿はせず、送信予定の内容をログに出すだけ(課題の「推奨」ルート)
    Logger.log('[WordPress送信ペイロード(確認用・実際には送信していません)]\n' + JSON.stringify(payload, null, 2));
    sheet.getRange(row, COL.WP_URL).setValue('(未投稿・ログで内容確認のみ)');
    updateStatus(sheet, row, STATUS.GENERATED);
    return;
  }

  updateStatus(sheet, row, STATUS.POSTING);
  const postUrl = postToWordPressDraft(payload, config);
  sheet.getRange(row, COL.WP_URL).setValue(postUrl);
  updateStatus(sheet, row, STATUS.POSTED);
}

/**
 * WordPressに送るペイロードを組み立てる。
 * status は絶対に 'draft' 固定(自動公開はしない約束)。
 */
function buildWordPressPayload(article) {
  return {
    title: article.title,
    content: article.contentHtml,
    status: 'draft', // ← 絶対に 'publish' にしない
    excerpt: article.metaDescription,
  };
}

function updateStatus(sheet, row, status, errorMessage) {
  sheet.getRange(row, COL.STATUS).setValue(status);
  sheet.getRange(row, COL.UPDATED_AT).setValue(new Date());
  if (errorMessage) {
    sheet.getRange(row, COL.ERROR).setValue(errorMessage);
  }
  SpreadsheetApp.flush(); // 途中経過をすぐシートに反映する(失敗時に進捗が分かるように)
}

// ============================================================
// 設定シートの読み取り
// ============================================================

function readConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    throw new Error('「設定」シートが見つかりません。先に「設定シートを初期化」を実行してください。');
  }

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty(PROP_KEYS.ANTHROPIC_API_KEY);
  const wpAppPassword = props.getProperty(PROP_KEYS.WP_APP_PASSWORD);

  if (!apiKey) throw new Error('Anthropic APIキーが未設定です。「初期設定」メニューから登録してください。');
  // 実投稿する場合のみ、WordPressのアプリケーションパスワードが必須
  if (!SKIP_REAL_WORDPRESS_POST && !wpAppPassword) {
    throw new Error('WordPressのアプリケーションパスワードが未設定です。「初期設定」メニューから登録してください。');
  }

  const model = configSheet.getRange(CONFIG_CELL.MODEL).getValue() || DEFAULT_MODEL;
  const wpBaseUrl = String(configSheet.getRange(CONFIG_CELL.WP_BASE_URL).getValue() || '').replace(/\/$/, '');
  const wpUsername = configSheet.getRange(CONFIG_CELL.WP_USERNAME).getValue();
  const targetReader = configSheet.getRange(CONFIG_CELL.TARGET_READER).getValue();
  const lengthGuide = configSheet.getRange(CONFIG_CELL.LENGTH_GUIDE).getValue() || '2000〜3000字';
  const sampleArticle = configSheet.getRange(CONFIG_CELL.SAMPLE_ARTICLE).getValue();

  if (!sampleArticle) throw new Error('「設定」シートにサンプル記事(理想の記事)を貼り付けてください。');
  // 実投稿する場合のみ、WordPressのURL・ユーザー名が必須
  if (!SKIP_REAL_WORDPRESS_POST && !wpBaseUrl) {
    throw new Error('「設定」シートにWordPress URLを入力してください。');
  }
  if (!SKIP_REAL_WORDPRESS_POST && !wpUsername) {
    throw new Error('「設定」シートにWordPressユーザー名を入力してください。');
  }

  return {
    model: model,
    apiKey: apiKey,
    wpBaseUrl: wpBaseUrl,
    wpUsername: wpUsername,
    wpAppPassword: wpAppPassword,
    targetReader: targetReader,
    lengthGuide: lengthGuide,
    sampleArticle: sampleArticle,
  };
}

// ============================================================
// Claude APIでの記事生成(プロンプト設計の4要素を組み込む)
// ============================================================

const RECORD_ARTICLE_TOOL = {
  name: 'record_blog_article',
  description: '生成したブログ記事を構造化して返す',
  input_schema: {
    type: 'object',
    properties: {
      title: {
        type: 'string',
        description: '記事タイトル。SEOキーワードを自然に含め、32文字前後を目安にする',
      },
      content_html: {
        type: 'string',
        description:
          '記事本文。WordPressにそのまま投稿できるHTML(見出しは<h2>、本文は<p>タグ)。' +
          '導入→H2見出し3つ→まとめの構成を守ること',
      },
      meta_description: {
        type: 'string',
        description: '検索結果に表示されるメタディスクリプション。120字以内',
      },
    },
    required: ['title', 'content_html', 'meta_description'],
  },
};

function buildPrompt(keyword, subKeyword, config) {
  // プロンプト設計の4要素: ①役割定義 ②サンプル記事の引用 ③構成テンプレート ④制約条件
  const parts = [];

  parts.push('あなたはBtoB SaaSのマーケティングライターです。'); // ① 役割定義

  parts.push(
    '以下は、クライアントが「これが理想の品質」と考えているサンプル記事です。' +
      'このトーン・マナー・文体を再現して書いてください。\n\n' +
      '----- サンプル記事 ここから -----\n' +
      config.sampleArticle +
      '\n----- サンプル記事 ここまで -----'
  ); // ② サンプル記事の引用

  parts.push(
    '次の記事構成テンプレートに従ってください(各セクションの文字数の目安は必ず守ること。目安を大きく下回らないよう、' +
      '各見出しの下に具体例・背景説明・実務上の注意点などを交えて本文をしっかり書き込むこと):\n' +
      '- 導入: 300字程度\n' +
      '- H2見出し1つ目: 500〜700字程度(<p>段落を2〜3個使う)\n' +
      '- H2見出し2つ目: 500〜700字程度(<p>段落を2〜3個使う)\n' +
      '- H2見出し3つ目: 500〜700字程度(<p>段落を2〜3個使う)\n' +
      '- まとめ: 200〜300字程度'
  ); // ③ 構成テンプレート

  const constraints = [
    '文字数は必ず' + config.lengthGuide + 'の範囲に収める(これは上限ではなく必須の範囲。下回ってはいけない)',
    'SEOキーワード「' + keyword + '」' + (subKeyword ? '、サブキーワード「' + subKeyword + '」' : '') + 'を自然な形で含める',
    '既存記事の丸写しではなく、オリジナルの表現を心がける',
  ];
  if (config.targetReader) {
    constraints.unshift('ターゲット読者は「' + config.targetReader + '」');
  }
  parts.push('制約条件: ' + constraints.join(' / ')); // ④ 制約条件

  parts.push(
    'record_blog_articleツールを使って、タイトル・本文(HTML)・メタディスクリプションを返してください。'
  );

  return parts.join('\n\n');
}

function callClaudeForArticle(keyword, subKeyword, config) {
  const prompt = buildPrompt(keyword, subKeyword, config);

  const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': config.apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify({
      model: config.model,
      max_tokens: 4096,
      tools: [RECORD_ARTICLE_TOOL],
      tool_choice: { type: 'tool', name: 'record_blog_article' },
      messages: [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();

  if (statusCode !== 200) {
    throw new Error('Claude APIエラー (status ' + statusCode + '): ' + body);
  }

  const json = JSON.parse(body);
  const toolUseBlock = (json.content || []).find(function (block) {
    return block.type === 'tool_use';
  });

  if (!toolUseBlock) {
    throw new Error('Claude APIから構造化データが返されませんでした: ' + body);
  }

  const input = toolUseBlock.input;
  return {
    title: input.title,
    contentHtml: input.content_html,
    metaDescription: input.meta_description,
  };
}

// ============================================================
// WordPress REST APIへの下書き投稿
// ============================================================

function postToWordPressDraft(payload, config) {
  const endpoint = config.wpBaseUrl + '/wp-json/wp/v2/posts';
  const auth = 'Basic ' + Utilities.base64Encode(config.wpUsername + ':' + config.wpAppPassword);

  const response = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: auth },
    payload: JSON.stringify(payload), // status: 'draft' 固定。絶対に 'publish' にしない。
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();

  if (statusCode !== 201) {
    throw new Error('WordPress投稿エラー (status ' + statusCode + '): ' + body);
  }

  const json = JSON.parse(body);
  // 編集画面へのリンク(本人が下書きを確認しやすいように)
  return config.wpBaseUrl + '/wp-admin/post.php?post=' + json.id + '&action=edit';
}
