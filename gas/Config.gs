// Config.gs - 設定・APIキー管理

const SHEET_ID = '1zqXH-cssTIf976gLhVs5QO2p3W0FfB_74x6nKn1K_uU';

const MODES = {
  CHINA    : '中国輸入',
  DOMESTIC : '国内メーカー',
};

const SHEET_NAMES = {
  SETTINGS      : '設定',
  DATA_CHINA    : 'データ_中国輸入',
  DATA_DOMESTIC : 'データ_国内メーカー',
  EXCLUDES      : '除外ワード',
  CANDIDATES    : '除外候補',
  SYNONYMS      : '同義語',
  COOC_CHINA    : '共起ワード_中国輸入',
  COOC_DOMESTIC : '共起ワード_国内メーカー',
  SUMMARY       : 'サマリー',
};

const SCORE_RULES = {
  HIDDEN_GEM : { min: 1,  max: 3,    label: '🎯隠れた狙い目' },
  TRENDING   : { min: 4,  max: 9,    label: '📈注目ワード'   },
  SATURATED  : { min: 10, max: 9999, label: '⚠️飽和状態'    },
};

const RANKING_TOP_N = 300;
const PRODUCTS_PER_KEYWORD = 5;  // データシートに横展開する上位商品数

// === フェーズ2: 共起分析 定数 ===
const COOC_SEED_MIN_COUNT = 5;    // 種ワードの最小出現回数
const COOC_MAX_SEEDS      = 30;   // 1モード当たり処理する種ワード最大数（GAS 6分制限対策）
const SUGGEST_DELAY_MS    = 500;  // サジェスト間隔
const SEARCH_API_DELAY_MS = 1100; // 商品検索API間隔（1 req/sec）
const SUGGEST_URL         = 'https://search.rakuten.co.jp/suggest?q=';
const SEARCH_API_URL      = 'https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706';

function getDataSheetName(mode) {
  return mode === MODES.DOMESTIC ? SHEET_NAMES.DATA_DOMESTIC : SHEET_NAMES.DATA_CHINA;
}

function getCooccurrenceSheetName(mode) {
  return mode === MODES.DOMESTIC ? SHEET_NAMES.COOC_DOMESTIC : SHEET_NAMES.COOC_CHINA;
}

// 除外ワード/除外候補シートの「有効」列 (1-based)
function getExcludeColumnForMode(mode) {
  return mode === MODES.DOMESTIC ? 2 : 1;  // A=中国輸入, B=国内会社
}

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    rakutenAppId   : props.getProperty('RAKUTEN_APP_ID'),
    discordWebhook : props.getProperty('DISCORD_WEBHOOK'),
    claudeApiKey   : props.getProperty('CLAUDE_API_KEY'),
    keepaApiKey    : props.getProperty('KEEPA_API_KEY'),
  };
}

/**
 * 初回セットアップ: ここに値を入力して実行 → 実行後は値を消してOK
 */
function setupApiKeys() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('RAKUTEN_APP_ID',  'ここに楽天アプリIDを入力');
  props.setProperty('DISCORD_WEBHOOK', 'ここにDiscord Webhook URLを入力');
  // props.setProperty('CLAUDE_API_KEY', 'ここにClaude APIキーを入力');
  // props.setProperty('KEEPA_API_KEY',  'ここにKeepa APIキーを入力');
  Logger.log('APIキーを保存しました');
}
