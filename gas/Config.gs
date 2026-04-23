// Config.gs - 設定・APIキー管理・定数

const SHEET_ID = '1zqXH-cssTIf976gLhVs5QO2p3W0FfB_74x6nKn1K_uU';

const MODES = {
  CHINA    : '中国輸入',
  DOMESTIC : '国内メーカー',
};

const SHEET_NAMES = {
  SETTINGS         : '設定',
  RANKING_CHINA    : 'ランキング_中国輸入',
  RANKING_DOMESTIC : 'ランキング_国内メーカー',
  EXCLUDES         : '除外ワード',
  CANDIDATES       : '除外候補',
  SYNONYMS         : '同義語',
  WORD_POOL        : '語彙プール',
  TAG_DICT         : 'タグ辞書',
  SURVEYED         : '調査済みワード',
  SUGGEST_CHINA    : '推奨ワード_中国輸入',
  SUGGEST_DOMESTIC : '推奨ワード_国内メーカー',
  SUMMARY          : 'サマリー',
};

const SCORE_RULES = {
  HIDDEN_GEM : { min: 1,  max: 3,    label: '🎯隠れた狙い目' },
  TRENDING   : { min: 4,  max: 9,    label: '📈注目ワード'   },
  SATURATED  : { min: 10, max: 9999, label: '⚠️飽和状態'    },
};

const RANKING_TOP_N = 300;
const PRODUCTS_PER_KEYWORD = 5;

// === 楽天API エンドポイント ===
const RAKUTEN_API_BASE     = 'https://app.rakuten.co.jp/services/api';
const RAKUTEN_ITEM_SEARCH  = RAKUTEN_API_BASE + '/IchibaItem/Search/20220601';
const RAKUTEN_ITEM_RANKING = RAKUTEN_API_BASE + '/IchibaItem/Ranking/20220601';
const RAKUTEN_GENRE_SEARCH = RAKUTEN_API_BASE + '/IchibaGenre/Search/20140222';
const RAKUTEN_TAG_SEARCH   = RAKUTEN_API_BASE + '/IchibaTag/Search/20140222';
const RAKUTEN_API_DELAY_MS = 1100;  // 1 req/sec 安全マージン

// === 語彙プール構築 ===
const WORDPOOL_STEP_TIME_LIMIT_MS  = 5 * 60 * 1000;  // 5分で中断→次回継続
const WORDPOOL_ITEM_PAGES          = 2;              // ItemSearch取得ページ数 (30件×2=60商品)
const WORDPOOL_RANKING_PAGES       = 2;              // ItemRanking取得ページ数 (30件×2=60位まで)
const WORDPOOL_TAG_FETCH_LIMIT     = 20;             // 1ジャンル1runあたりのTagSearch呼び出し上限

// === お宝分析 ===
const HIDDEN_GEM_MAX_PER_GENRE = 50;        // ジャンル毎の出力上限
const HIDDEN_GEM_NEW_DAYS      = 14;        // 初出から何日間 NEW! 表示継続
const HIDDEN_GEM_URL_COUNT     = 6;         // URL列数 (URL1〜URL6)
const HIDDEN_GEM_COLOR_SURVEYED = '#F0F0F0'; // 調査済み・非該当月
const HIDDEN_GEM_COLOR_RECALL   = '#E8F0FE'; // 調査済み・再表示月

// === Gemini API ===
const GEMINI_MODEL        = 'gemini-2.5-flash-lite';  // シンプルなJSON整形用に高速版
const GEMINI_API_URL_BASE = 'https://generativelanguage.googleapis.com/v1beta/models/';
const GEMINI_MAX_TOKENS   = 8192;
const GEMINI_BATCH_SIZE   = 50;  // 1回のプロンプトに含めるワード数上限

// === チェックポイント用ScriptPropertiesキー ===
const WP_CHECKPOINT_DATE_KEY  = 'WP_CHECKPOINT_DATE';
const WP_CHECKPOINT_INDEX_KEY = 'WP_CHECKPOINT_INDEX';

function getRankingSheetName(mode) {
  return mode === MODES.DOMESTIC ? SHEET_NAMES.RANKING_DOMESTIC : SHEET_NAMES.RANKING_CHINA;
}

function getSuggestSheetName(mode) {
  return mode === MODES.DOMESTIC ? SHEET_NAMES.SUGGEST_DOMESTIC : SHEET_NAMES.SUGGEST_CHINA;
}

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    rakutenAppId   : props.getProperty('RAKUTEN_APP_ID'),
    discordWebhook : props.getProperty('DISCORD_WEBHOOK'),
    geminiApiKey   : props.getProperty('GEMINI_API_KEY'),
    claudeApiKey   : props.getProperty('CLAUDE_API_KEY'),  // 後方互換
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
