export interface TweetItem {
  id: number;
  text: string;
  post_time?: string; // e.g. "7～10", "22～2", "7～10, 18～21", or ""
  image_path: string; // e.g. "images/photo1.jpg" or absolute path or folder
  status: 'pending' | 'posted' | 'failed' | 'skipped';
  posted_at?: string;
  error_message?: string;
}

export interface ThemeSheet {
  id: string;
  name: string; // e.g. "テーマ1", "お知らせ"
  activeTimeRange: string; // e.g. "0:00-0:00"
  weekdays: string; // e.g. "毎日", "平日", "月水金"
  intervalRange: string; // e.g. "120~360分"
  tweetOrderMode: number; // 0: 上から順, 1: ランダム
  isActive: boolean; // true: 有効, false: 無効
  memo?: string;
  tweets: TweetItem[];
}

export interface AccountConfig {
  id: string;
  name: string;
  handle: string; // e.g. "@my_shop_account"
  postMethod?: 'browser' | 'api'; // 'browser' (Selenium / Playwright 完全無料) or 'api'
  // ブラウザ自動操縦用 (API完全不要)
  authToken?: string; // Colab / Cookie用 auth_token
  loginId?: string; // Xのログインメールアドレス or ユーザー名
  loginPassword?: string; // Xのパスワード
  authSecret2FA?: string; // 2段階認証TOTPキー（任意）
  profileDirName?: string; // Chromeプロファイル保存フォルダ（ログインセッション維持用）
  // API用（従来）
  apiKey: string;
  apiSecret: string;
  accessToken: string;
  accessSecret: string;
  bearerToken?: string;
  // Excel・間隔共通
  excelFileName: string;
  minIntervalSec: number; // e.g. 60
  maxIntervalSec: number; // e.g. 600
  intervalMinMinutes?: number; // e.g. 120
  intervalMaxMinutes?: number; // e.g. 360
  activeTimeStart?: string; // e.g. "0:00"
  activeTimeEnd?: string; // e.g. "0:00"
  tweetOrderMode?: number; // 0: 上から順, 1: ランダム
  repeatMode: boolean; // if true, reset pending after all posted
  isActive: boolean;
  tweets: TweetItem[];
  // テーマ別シート一覧（各シートが独立したテーマ＆設定を持つ）
  themes?: ThemeSheet[];
  // 提案1: 1窓・順番交代ローテーション個別設定（省略時は全体設定に準拠）
  actionsPerBatch?: number; // 1アカウントの1回の巡回で処理する件数 (例: 3)
  dmTargets?: DmTargetItem[];
}

export interface DmTargetItem {
  id: number;
  targetUsername: string; // 送信先 @ユーザー名
  messageText: string;    // 送信DM本文
  autoFollow: boolean;    // 送信前にフォローするか
  status: 'pending' | 'sent' | 'failed' | 'skipped';
  sent_at?: string;
  error_message?: string;
}

export interface BotOperationSettings {
  // 提案1: 1窓・順番交代ローテーション
  enableRotation: boolean; // 1窓交代ローテーションを行うか（Trueで常時1窓のみ）
  actionsPerRotation: number; // 1アカウントあたり何件ずつ実行して交代するか（例: 3〜5件）
  cycleCooldownMinutes: number; // 全アカウント1巡後の待機時間（分、例: 15〜30分）
  // 提案2: 画像・動画（リール等）の読み込みブロック
  blockMedia: boolean; // .jpg, .png, .mp4 等のメディア通信を遮断しメモリ50〜70%削減
  blockFonts: boolean; // Webフォント等の重いリソースも追加ブロック
}

export interface ExecutionLog {
  id: string;
  timestamp: string;
  accountId: string;
  accountHandle: string;
  level: 'info' | 'success' | 'warn' | 'error';
  message: string;
  tweetId?: number;
}
