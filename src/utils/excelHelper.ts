import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { AccountConfig, TweetItem, ThemeSheet, DmTargetItem } from '../types';

export const DEFAULT_THEME_MEMO =
  '投稿モード: 0=上から順 / 1=ランダム (数字だけでもOK) ｜ 下のシートを追加・複製するとテーマ別に管理できます';

const TEMPLATE_MAX_ROW = 200;

function triggerBlobDownload(bytes: Uint8Array, fileName: string) {
  const safeFileName = fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`;
  const blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = safeFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * AccountConfig からテーマ配列を正規化取得（レガシー互換）
 */
export function getAccountThemes(account: AccountConfig): ThemeSheet[] {
  if (account.themes && account.themes.length > 0) {
    return account.themes;
  }
  const minMin = account.intervalMinMinutes ?? Math.max(1, Math.round((account.minIntervalSec || 7200) / 60));
  const maxMin = account.intervalMaxMinutes ?? Math.max(minMin, Math.round((account.maxIntervalSec || 21600) / 60));
  const startTime = account.activeTimeStart || '0:00';
  const endTime = account.activeTimeEnd || '0:00';
  const mode = account.tweetOrderMode !== undefined ? account.tweetOrderMode : 1;

  return [
    {
      id: 'theme-1',
      name: 'テーマ1',
      activeTimeRange: `${startTime}-${endTime}`,
      weekdays: '毎日',
      intervalRange: `${minMin}~${maxMin}分`,
      tweetOrderMode: mode,
      isActive: account.isActive !== false,
      memo: DEFAULT_THEME_MEMO,
      tweets: account.tweets || [],
    },
  ];
}

/**
 * 1. アカウント情報.xlsx の生成（色付け＆罫線付き）
 */
export async function accountsInfoToExcelBinary(accounts: AccountConfig[]): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('アカウント情報');

  const headers = ['アカウントID', 'アカウントpw', 'アカウント名', 'auth_token'];
  const headerRow = ws.addRow(headers);
  headerRow.height = 26;
  headerRow.eachCell((cell) => {
    cell.font = { name: 'Yu Gothic UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF1E3A8A' } },
      left: { style: 'thin', color: { argb: 'FF3B82F6' } },
      bottom: { style: 'medium', color: { argb: 'FF1E3A8A' } },
      right: { style: 'thin', color: { argb: 'FF3B82F6' } },
    };
  });

  accounts.forEach((acc, idx) => {
    const r = ws.addRow([
      acc.loginId || acc.handle || '',
      acc.loginPassword || '',
      acc.name || '',
      acc.authToken || '',
    ]);
    r.height = 22;
    const bgArgb = idx % 2 === 1 ? 'FFF1F5F9' : 'FFFFFFFF';
    r.eachCell((cell) => {
      cell.font = { name: 'Yu Gothic UI', size: 10, color: { argb: 'FF0F172A' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      };
    });
  });

  ws.columns = [
    { width: 24 },
    { width: 24 },
    { width: 22 },
    { width: 36 },
  ];
  ws.views = [{ state: 'frozen', ySplit: 1, showGridLines: true }];

  const buffer = await wb.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

export async function downloadAccountsInfoExcel(
  accounts: AccountConfig[],
  fileName = 'アカウント情報.xlsx'
): Promise<void> {
  const bytes = await accountsInfoToExcelBinary(accounts);
  triggerBlobDownload(bytes, fileName);
}

/**
 * アカウント情報.xlsx の読み込み
 */
export async function parseAccountsInfoExcel(
  file: File
): Promise<Array<{ id: string; loginId: string; loginPassword?: string; name?: string; authToken?: string }>> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rows: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        const parsed = rows
          .map((row, idx) => {
            const loginId = String(
              row['アカウントID'] || row['ID'] || row['ユーザー名'] || row['username'] || row['loginId'] || ''
            ).trim();
            const loginPassword = String(
              row['アカウントpw'] || row['pw'] || row['パスワード'] || row['password'] || ''
            ).trim();
            const name = String(row['アカウント名'] || row['名前'] || row['name'] || `アカウント #${idx + 1}`).trim();
            const authToken = String(row['auth_token'] || row['authToken'] || row['トークン'] || '').trim();

            return {
              id: `imported-${Date.now()}-${idx}`,
              loginId,
              loginPassword,
              name,
              authToken,
            };
          })
          .filter((item) => item.loginId || item.name);

        resolve(parsed);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 2. テーマ別シートを色付き＆削除耐性（200行書式＋条件付き書式）で書き込み
 */
function populateThemeWorksheet(ws: ExcelJS.Worksheet, theme: ThemeSheet) {
  const borderThinSlate: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FFCBD5E1' } };
  const gridBorder: Partial<ExcelJS.Borders> = {
    top: borderThinSlate,
    left: borderThinSlate,
    bottom: borderThinSlate,
    right: borderThinSlate,
  };

  // Row 1: テーマ設定ヘッダー (紺色)
  const settingsHeaders = ['実行時間', '曜日', 'ポスト間隔', '投稿モード', '有効/無効', 'メモ'];
  const r1 = ws.getRow(1);
  r1.height = 28;
  settingsHeaders.forEach((h, i) => {
    const c = r1.getCell(i + 1);
    c.value = h;
    c.font = { name: 'Yu Gothic UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = {
      top: { style: 'medium', color: { argb: 'FF1E3A8A' } },
      left: { style: 'thin', color: { argb: 'FF3B82F6' } },
      bottom: { style: 'thin', color: { argb: 'FF93C5FD' } },
      right: { style: 'thin', color: { argb: 'FF3B82F6' } },
    };
  });

  // Row 2: テーマ設定値
  const modeStr = theme.tweetOrderMode === 0 ? '0: 上から順' : '1: ランダム';
  const activeStr = theme.isActive ? '有効' : '無効';
  const r2Vals = [
    theme.activeTimeRange || '0:00-0:00',
    theme.weekdays || '毎日',
    theme.intervalRange || '120~360分',
    modeStr,
    activeStr,
    theme.memo || DEFAULT_THEME_MEMO,
  ];
  const r2 = ws.getRow(2);
  r2.height = 42;
  r2Vals.forEach((val, i) => {
    const c = r2.getCell(i + 1);
    c.value = val;
    c.border = {
      top: { style: 'thin', color: { argb: 'FF93C5FD' } },
      left: { style: 'thin', color: { argb: 'FF93C5FD' } },
      bottom: { style: 'medium', color: { argb: 'FF1E3A8A' } },
      right: { style: 'thin', color: { argb: 'FF93C5FD' } },
    };
    if (i < 4) {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F1FF' } };
      c.font = { name: 'Yu Gothic UI', size: 11, bold: true, color: { argb: 'FF1E3A8A' } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    } else if (i === 4) {
      c.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: theme.isActive ? 'FFDCFCE7' : 'FFFEE2E2' },
      };
      c.font = {
        name: 'Yu Gothic UI',
        size: 11,
        bold: true,
        color: { argb: theme.isActive ? 'FF15803D' : 'FFB91C1C' },
      };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    } else {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      c.font = { name: 'Yu Gothic UI', size: 9, color: { argb: 'FF475569' } };
      c.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    }
  });

  // Row 3: 空行スペーサー
  ws.getRow(3).height = 10;

  // Row 4: 投稿テーブルヘッダー
  const tweetHeaders = ['No', 'ツイート文章', '投稿時間', '画像パス', 'ステータス', '投稿日時'];
  const r4 = ws.getRow(4);
  r4.height = 26;
  tweetHeaders.forEach((h, i) => {
    const c = r4.getCell(i + 1);
    c.value = h;
    c.font = { name: 'Yu Gothic UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: i === 2 ? 'FFB45309' : 'FF1E40AF' },
    };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = {
      top: { style: 'medium', color: { argb: 'FF1E3A8A' } },
      left: { style: 'thin', color: { argb: 'FF475569' } },
      bottom: { style: 'medium', color: { argb: 'FF1E3A8A' } },
      right: { style: 'thin', color: { argb: 'FF475569' } },
    };
  });

  // Row 5〜max(200, tweets.length+30): データ＆事前書式行
  const tweets = theme.tweets || [];
  const lastRow = Math.max(TEMPLATE_MAX_ROW, tweets.length + 30);

  for (let idx = 0; idx <= lastRow - 5; idx++) {
    const rNum = 5 + idx;
    const isEven = rNum % 2 === 0;
    const baseBg = isEven ? 'FFF1F5F9' : 'FFFFFFFF';
    const timeEmptyBg = isEven ? 'FFFEF9E7' : 'FFFFFDF5';

    const t = idx < tweets.length ? tweets[idx] : null;
    const row = ws.getRow(rNum);

    const noCell = row.getCell(1);
    noCell.value = t ? t.id || idx + 1 : null;
    noCell.font = { name: 'Yu Gothic UI', size: 10, bold: true, color: { argb: 'FF64748B' } };
    noCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseBg } };
    noCell.alignment = { horizontal: 'center', vertical: 'middle' };
    noCell.border = gridBorder;

    const textCell = row.getCell(2);
    textCell.value = t ? t.text || '' : null;
    textCell.font = { name: 'Yu Gothic UI', size: 10, color: { argb: 'FF0F172A' } };
    textCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseBg } };
    textCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    textCell.border = gridBorder;

    const ptVal = t?.post_time?.trim() || '';
    const ptCell = row.getCell(3);
    ptCell.value = ptVal || null;
    ptCell.font = { name: 'Yu Gothic UI', size: 10, bold: true, color: { argb: 'FF92400E' } };
    ptCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ptVal ? 'FFFEF3C7' : timeEmptyBg } };
    ptCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    ptCell.border = gridBorder;

    const imgCell = row.getCell(4);
    imgCell.value = t && t.image_path ? t.image_path : null;
    imgCell.font = { name: 'Yu Gothic UI', size: 10, color: { argb: 'FF334155' } };
    imgCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseBg } };
    imgCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    imgCell.border = gridBorder;

    const stVal = t ? t.status || 'pending' : null;
    const stCell = row.getCell(5);
    stCell.value = stVal;
    stCell.alignment = { horizontal: 'center', vertical: 'middle' };
    stCell.border = gridBorder;
    if (stVal === 'posted') {
      stCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
      stCell.font = { name: 'Yu Gothic UI', size: 10, bold: true, color: { argb: 'FF15803D' } };
    } else if (stVal === 'pending') {
      stCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
      stCell.font = { name: 'Yu Gothic UI', size: 10, bold: true, color: { argb: 'FFB45309' } };
    } else if (stVal === 'failed') {
      stCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
      stCell.font = { name: 'Yu Gothic UI', size: 10, bold: true, color: { argb: 'FFB91C1C' } };
    } else if (stVal === 'skipped') {
      stCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      stCell.font = { name: 'Yu Gothic UI', size: 10, color: { argb: 'FF475569' } };
    } else {
      stCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseBg } };
      stCell.font = { name: 'Yu Gothic UI', size: 10, bold: true, color: { argb: 'FFB45309' } };
    }

    const dtCell = row.getCell(6);
    dtCell.value = t && t.posted_at ? t.posted_at : null;
    dtCell.font = { name: 'Yu Gothic UI', size: 10, color: { argb: 'FF475569' } };
    dtCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: baseBg } };
    dtCell.alignment = { horizontal: 'center', vertical: 'middle' };
    dtCell.border = gridBorder;

    if (t && t.text && (t.text.includes('\n') || t.text.length > 48)) {
      const lines = Math.max(t.text.split('\n').length, Math.floor(t.text.length / 36) + 1);
      row.height = Math.min(140, Math.max(24, lines * 16 + 6));
    } else {
      row.height = 24;
    }
  }

  ws.columns = [
    { width: 14 }, // A: 実行時間 / No
    { width: 54 }, // B: 曜日 / ツイート文章
    { width: 18 }, // C: ポスト間隔 / 投稿時間
    { width: 26 }, // D: 投稿モード / 画像パス
    { width: 14 }, // E: 有効/無効 / ステータス
    { width: 38 }, // F: メモ / 投稿日時
  ];
  ws.views = [{ state: 'frozen', ySplit: 4, showGridLines: true }];

  // 条件付き書式（セルをDelete・ペーストしても色・ゼブラ保護）
  ws.addConditionalFormatting({
    ref: 'E2',
    rules: [
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['"無効"'],
        priority: 1,
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFEE2E2' } },
          font: { bold: true, color: { argb: 'FFB91C1C' } },
        },
      },
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['"有効"'],
        priority: 2,
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFDCFCE7' } },
          font: { bold: true, color: { argb: 'FF15803D' } },
        },
      },
    ],
  });

  ws.addConditionalFormatting({
    ref: `E5:E${lastRow}`,
    rules: [
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['"posted"'],
        priority: 1,
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFDCFCE7' } },
          font: { bold: true, color: { argb: 'FF15803D' } },
        },
      },
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['"投稿済"'],
        priority: 2,
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFDCFCE7' } },
          font: { bold: true, color: { argb: 'FF15803D' } },
        },
      },
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['"pending"'],
        priority: 3,
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFEF3C7' } },
          font: { bold: true, color: { argb: 'FFB45309' } },
        },
      },
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['"failed"'],
        priority: 4,
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFEE2E2' } },
          font: { bold: true, color: { argb: 'FFB91C1C' } },
        },
      },
      {
        type: 'cellIs',
        operator: 'equal',
        formulae: ['"skipped"'],
        priority: 5,
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFE2E8F0' } },
          font: { color: { argb: 'FF475569' } },
        },
      },
      {
        type: 'expression',
        formulae: ['AND(ISBLANK(E5),MOD(ROW(),2)=0)'],
        priority: 6,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFF1F5F9' } } },
      },
      {
        type: 'expression',
        formulae: ['AND(ISBLANK(E5),MOD(ROW(),2)=1)'],
        priority: 7,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFFFF' } } },
      },
    ],
  });

  ws.addConditionalFormatting({
    ref: `C5:C${lastRow}`,
    rules: [
      {
        type: 'expression',
        formulae: ['LEN(TRIM(C5))>0'],
        priority: 1,
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFEF3C7' } },
          font: { bold: true, color: { argb: 'FF92400E' } },
        },
      },
      {
        type: 'expression',
        formulae: ['AND(LEN(TRIM(C5))=0,MOD(ROW(),2)=0)'],
        priority: 2,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFEF9E7' } } },
      },
      {
        type: 'expression',
        formulae: ['AND(LEN(TRIM(C5))=0,MOD(ROW(),2)=1)'],
        priority: 3,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFDF5' } } },
      },
    ],
  });

  for (const rng of [`A5:B${lastRow}`, `D5:D${lastRow}`, `F5:F${lastRow}`]) {
    ws.addConditionalFormatting({
      ref: rng,
      rules: [
        {
          type: 'expression',
          formulae: ['MOD(ROW(),2)=0'],
          priority: 1,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFF1F5F9' } } },
        },
        {
          type: 'expression',
          formulae: ['MOD(ROW(),2)=1'],
          priority: 2,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFFFF' } } },
        },
      ],
    });
  }
}

/**
 * アカウント名.xlsx（全テーマシート入り・カラー＆削除耐性書式）のバイナリ出力
 */
export async function accountMultiTabToBinary(account: AccountConfig): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const themes = getAccountThemes(account);
  const usedNames = new Set<string>();

  themes.forEach((theme, idx) => {
    let rawName = (theme.name || `テーマ${idx + 1}`).replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || `テーマ${idx + 1}`;
    let sheetName = rawName;
    let suffix = 2;
    while (usedNames.has(sheetName)) {
      sheetName = `${rawName.slice(0, 27)}_${suffix++}`;
    }
    usedNames.add(sheetName);

    const ws = wb.addWorksheet(sheetName);
    populateThemeWorksheet(ws, theme);
  });

  const buffer = await wb.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

/**
 * アカウント名.xlsx のダウンロード
 */
export async function downloadAccountMultiTabExcel(
  account: AccountConfig,
  customFileName?: string
): Promise<void> {
  const bytes = await accountMultiTabToBinary(account);
  const baseName = customFileName || account.excelFileName || `${account.name}.xlsx`;
  triggerBlobDownload(bytes, baseName);
}

export async function downloadAccountExcel(
  tweets: TweetItem[],
  fileName: string,
  account?: AccountConfig
): Promise<void> {
  if (account) {
    await downloadAccountMultiTabExcel({ ...account, tweets }, fileName);
  } else {
    const fallbackAcc: AccountConfig = {
      id: 'temp',
      name: fileName.replace('.xlsx', ''),
      handle: '',
      apiKey: '',
      apiSecret: '',
      accessToken: '',
      accessSecret: '',
      excelFileName: fileName,
      minIntervalSec: 7200,
      maxIntervalSec: 21600,
      intervalMinMinutes: 120,
      intervalMaxMinutes: 360,
      activeTimeStart: '0:00',
      activeTimeEnd: '0:00',
      tweetOrderMode: 1,
      repeatMode: true,
      isActive: true,
      tweets,
    };
    await downloadAccountMultiTabExcel(fallbackAcc, fileName);
  }
}

export async function tweetsToExcelBinary(tweets: TweetItem[]): Promise<Uint8Array> {
  const fallbackAcc: AccountConfig = {
    id: 'temp',
    name: 'tweets',
    handle: '',
    apiKey: '',
    apiSecret: '',
    accessToken: '',
    accessSecret: '',
    excelFileName: 'tweets.xlsx',
    minIntervalSec: 7200,
    maxIntervalSec: 21600,
    intervalMinMinutes: 120,
    intervalMaxMinutes: 360,
    activeTimeStart: '0:00',
    activeTimeEnd: '0:00',
    tweetOrderMode: 1,
    repeatMode: true,
    isActive: true,
    tweets,
  };
  return accountMultiTabToBinary(fallbackAcc);
}

/**
 * Excel 読み込みパーサー
 * 新「テーマ別マルチシート形式（1〜2行目設定＋4行目ヘッダー＋投稿時間列）」と
 * 旧「設定＋ツイート文章形式」の両方を自動判別して読み込む
 */
export interface ParsedExcelResult {
  themes: ThemeSheet[];
  tweets: TweetItem[];
  settings?: {
    intervalMinMinutes?: number;
    intervalMaxMinutes?: number;
    activeTimeStart?: string;
    activeTimeEnd?: string;
    tweetOrderMode?: number;
  };
}

function parseLegacyIntervalMinMax(intervalText: string): { minM: number; maxM: number } {
  const nums = (intervalText.normalize('NFKC').match(/\d+/g) || []).map((n) => parseInt(n, 10));
  if (nums.length >= 2) {
    const minM = Math.max(1, Math.min(nums[0], nums[1]));
    const maxM = Math.max(minM, Math.max(nums[0], nums[1]));
    return { minM, maxM };
  }
  if (nums.length === 1) {
    const v = Math.max(1, nums[0]);
    return { minM: v, maxM: v };
  }
  return { minM: 120, maxM: 360 };
}

export async function parseExcelFile(file: File): Promise<ParsedExcelResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });

        // 旧「設定」シートがあれば互換救出
        let legacySettings: ParsedExcelResult['settings'] = undefined;
        const configSheetName = workbook.SheetNames.find(
          (name) => name === '設定' || name === 'Config' || name === 'Setting'
        );
        if (configSheetName) {
          const configSheet = workbook.Sheets[configSheetName];
          const rawAoa: any[][] = XLSX.utils.sheet_to_json(configSheet, { header: 1, defval: '' });
          let intervalMin = 120;
          let intervalMax = 360;
          let startTime = '0:00';
          let endTime = '0:00';
          let mode = 1;

          for (let i = 0; i < rawAoa.length; i++) {
            const rowStr = rawAoa[i].join(' ');
            if (rowStr.includes('ツイート間隔') || rowStr.includes('範囲内')) {
              const valRow = rawAoa[i + 1] || rawAoa[i + 2] || [];
              const nums = valRow.filter((v: any) => typeof v === 'number' || (!isNaN(Number(v)) && v !== ''));
              if (nums.length >= 2) {
                intervalMin = Math.max(1, Number(nums[0]));
                intervalMax = Math.max(intervalMin, Number(nums[1]));
              }
            }
            if (rowStr.includes('ツイート実行時間') || rowStr.includes('時:分')) {
              const valRow = rawAoa[i + 1] || rawAoa[i + 2] || [];
              const times = valRow.filter((v: any) => String(v).includes(':'));
              if (times.length >= 2) {
                startTime = String(times[0]).trim();
                endTime = String(times[1]).trim();
              }
            }
            if (rowStr.includes('モード') || rowStr.includes('0:上から順')) {
              const valRow = rawAoa[i + 1] || rawAoa[i + 2] || [];
              const foundMode = valRow.find((v: any) => v === 0 || v === 1 || v === '0' || v === '1');
              if (foundMode !== undefined) mode = Number(foundMode);
            }
          }
          legacySettings = {
            intervalMinMinutes: intervalMin,
            intervalMaxMinutes: intervalMax,
            activeTimeStart: startTime,
            activeTimeEnd: endTime,
            tweetOrderMode: mode,
          };
        }

        const themes: ThemeSheet[] = [];
        for (let sIdx = 0; sIdx < workbook.SheetNames.length; sIdx++) {
          const sName = workbook.SheetNames[sIdx];
          if (sName === '設定' || sName === '日常ツイート') continue;

          const ws = workbook.Sheets[sName];
          const aoa: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (aoa.length === 0) continue;

          const row0 = (aoa[0] || []).map((v) => String(v || '').trim());
          const hasThemeBanner = row0[0]?.includes('実行時間') && (row0[1]?.includes('曜日') || row0[2]?.includes('間隔'));

          let activeTimeRange = legacySettings
            ? `${legacySettings.activeTimeStart || '0:00'}-${legacySettings.activeTimeEnd || '0:00'}`
            : '0:00-0:00';
          let weekdays = '毎日';
          let intervalRange = legacySettings
            ? `${legacySettings.intervalMinMinutes || 120}~${legacySettings.intervalMaxMinutes || 360}分`
            : '120~360分';
          let tweetOrderMode = legacySettings?.tweetOrderMode ?? 1;
          let isActive = true;
          let memo = DEFAULT_THEME_MEMO;

          if (hasThemeBanner && aoa.length > 1) {
            const row1 = aoa[1] || [];
            if (String(row1[0] || '').trim()) activeTimeRange = String(row1[0]).trim();
            if (String(row1[1] || '').trim()) weekdays = String(row1[1]).trim();
            if (String(row1[2] || '').trim()) intervalRange = String(row1[2]).trim();
            const rawMode = String(row1[3] || '').trim();
            if (rawMode) {
              tweetOrderMode = rawMode.startsWith('0') || (rawMode.includes('上') && !rawMode.includes('ランダム')) ? 0 : 1;
            }
            const rawActive = String(row1[4] || '').trim();
            if (rawActive) {
              isActive = !['無効', '停止', 'OFF', 'off', '0', 'false'].includes(rawActive);
            }
            if (String(row1[5] || '').trim()) memo = String(row1[5]).trim();
          }

          // 投稿テーブルのヘッダー行を特定
          let headerRowIdx = hasThemeBanner ? 3 : 0;
          for (let r = hasThemeBanner ? 2 : 0; r < Math.min(aoa.length, 15); r++) {
            const rowVals = (aoa[r] || []).map((v) => String(v || '').trim().toLowerCase());
            if (rowVals.some((v) => ['ツイート文章', 'ポスト文章', '本文', 'text'].includes(v))) {
              headerRowIdx = r;
              break;
            }
          }

          const headers = (aoa[headerRowIdx] || []).map((v) => String(v || '').trim().toLowerCase());
          let noCol = headers.findIndex((h) => ['no', 'no.', 'id', '番号'].includes(h));
          let textCol = headers.findIndex((h) =>
            ['ツイート文章', 'ポスト文章', 'ツイート本文', '本文', 'text'].includes(h)
          );
          const ptCol = headers.findIndex((h) => ['投稿時間', '投稿時間帯', '時間指定', 'post_time', 'time'].includes(h));
          let imgCol = headers.findIndex((h) => ['画像パス', '画像', '画像フォルダ', 'image_path', 'image'].includes(h));
          let statusCol = headers.findIndex((h) => ['ステータス', 'status'].includes(h));
          let dateCol = headers.findIndex((h) => ['投稿日時', 'posted_at'].includes(h));

          if (textCol < 0) textCol = 1;
          if (noCol < 0) noCol = 0;
          if (imgCol < 0) imgCol = ptCol >= 0 ? 3 : 2;
          if (statusCol < 0) statusCol = ptCol >= 0 ? 4 : 3;
          if (dateCol < 0) dateCol = ptCol >= 0 ? 5 : 4;

          const parsedTweets: TweetItem[] = [];
          for (let r = headerRowIdx + 1; r < aoa.length; r++) {
            const row = aoa[r] || [];
            const text = String(row[textCol] ?? '').trim();
            const postTime = ptCol >= 0 ? String(row[ptCol] ?? '').trim() : '';
            const imagePath = String(row[imgCol] ?? '').trim();
            if (!text && !imagePath) continue;

            const rawStatus = String(row[statusCol] ?? '').trim().toLowerCase();
            let status: TweetItem['status'] = 'pending';
            if (['posted', '投稿済', '投稿済み', '完了'].includes(rawStatus)) status = 'posted';
            else if (['failed', '失敗', 'エラー'].includes(rawStatus)) status = 'failed';
            else if (['skipped', 'スキップ'].includes(rawStatus)) status = 'skipped';

            parsedTweets.push({
              id: Number(row[noCol]) || parsedTweets.length + 1,
              text,
              post_time: postTime,
              image_path: imagePath,
              status,
              posted_at: String(row[dateCol] ?? '').trim(),
            });
          }

          themes.push({
            id: `theme-${Date.now()}-${sIdx}`,
            name: sName,
            activeTimeRange,
            weekdays,
            intervalRange,
            tweetOrderMode,
            isActive,
            memo,
            tweets: parsedTweets,
          });
        }

        const primaryTheme = themes[0];
        const parsedInterval = primaryTheme ? parseLegacyIntervalMinMax(primaryTheme.intervalRange) : { minM: 120, maxM: 360 };
        const [startT, endT] = primaryTheme
          ? primaryTheme.activeTimeRange.split(/[-〜～]/).map((s) => s.trim())
          : ['0:00', '0:00'];

        resolve({
          themes,
          tweets: primaryTheme?.tweets || [],
          settings: legacySettings || {
            intervalMinMinutes: parsedInterval.minM,
            intervalMaxMinutes: parsedInterval.maxM,
            activeTimeStart: startT || '0:00',
            activeTimeEnd: endT || startT || '0:00',
            tweetOrderMode: primaryTheme?.tweetOrderMode ?? 1,
          },
        });
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 7. dm_targets.xlsx 用のワークシート生成（自動DM＆フォロー用）
 */
export function dmTargetsToWorksheet(targets: DmTargetItem[]): XLSX.WorkSheet {
  const headers = ['id', 'target_username', 'message', 'status', 'auto_follow', 'sent_at', 'error_message'];
  const rows = targets.map((t) => [
    t.id,
    t.targetUsername.startsWith('@') ? t.targetUsername : `@${t.targetUsername}`,
    t.messageText,
    t.status,
    t.autoFollow ? 1 : 0,
    t.sent_at || '',
    t.error_message || '',
  ]);

  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = [
    { wch: 8 },
    { wch: 22 },
    { wch: 45 },
    { wch: 12 },
    { wch: 12 },
    { wch: 20 },
    { wch: 25 },
  ];
  return ws;
}

export async function dmTargetsToExcelBinary(targets: DmTargetItem[]): Promise<Uint8Array> {
  const ws = dmTargetsToWorksheet(targets);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'DM送信先');
  const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  return new Uint8Array(buffer);
}

export async function downloadDmTargetsExcel(
  targets: DmTargetItem[],
  fileName = 'dm_targets.xlsx'
): Promise<void> {
  const bytes = await dmTargetsToExcelBinary(targets);
  triggerBlobDownload(bytes, fileName);
}
