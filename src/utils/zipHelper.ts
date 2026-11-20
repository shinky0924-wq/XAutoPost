import JSZip from 'jszip';
import { AccountConfig, BotOperationSettings, DmTargetItem } from '../types';
import { accountMultiTabToBinary, accountsInfoToExcelBinary } from './excelHelper';
import autoTweetPlaywrightRaw from '../../auto_tweet_playwright.py?raw';
import updatePs1Raw from '../../update.ps1?raw';
import updateThemesExcelRaw from '../../update_themes_excel.py?raw';
import xBotGuiPyRaw from '../../x_bot_gui.py?raw';
import xBotGuiVbsRaw from '../../X_BOT_GUI.vbs?raw';
import stopBotVbsRaw from '../../STOP_BOT.vbs?raw';
import { BROWSER_REQUIREMENTS, DEFAULT_SETTINGS } from '../data/defaultData';

export async function generatePythonProjectZip(
  accounts: AccountConfig[],
  settings?: BotOperationSettings,
  _dmTargets?: DmTargetItem[]
): Promise<void> {
  const currentSettings = settings || DEFAULT_SETTINGS;
  const zip = new JSZip();

  // 0. 最小限の起動・停止・アップデート用ファイルのみ生成（重複ランチャー類は生成しない）
  zip.file('X_BOT_GUI.vbs', xBotGuiVbsRaw.replace(/\r?\n/g, '\r\n'));
  zip.file('STOP_BOT.vbs', stopBotVbsRaw.replace(/\r?\n/g, '\r\n'));

  const ps1Normalized = updatePs1Raw.replace(/\r?\n/g, '\r\n');
  const ps1WithBom = ps1Normalized.startsWith('\uFEFF') ? ps1Normalized : '\uFEFF' + ps1Normalized;
  zip.file('update.ps1', ps1WithBom);
  zip.file(
    'UPDATE.bat',
    '@echo off\r\nchcp 65001 >nul\r\ntitle X_AutoTweetBot Update\r\ncd /d "%~dp0"\r\npowershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1"\r\nif %errorlevel% neq 0 (\r\n    echo.\r\n    echo [ERROR] Update failed with code %errorlevel%.\r\n    pause\r\n)\r\n'
  );

  // 1. 実行・GUI・Excel更新本体（Python）
  zip.file('auto_tweet_playwright.py', autoTweetPlaywrightRaw);
  zip.file('x_bot_gui.py', xBotGuiPyRaw);
  zip.file('update_themes_excel.py', updateThemesExcelRaw);
  zip.file('requirements.txt', BROWSER_REQUIREMENTS);

  // 2. 設定・アカウント情報
  zip.file('settings.json', JSON.stringify(currentSettings, null, 2));

  const accountsJsonData = accounts.map((acc) => ({
    id: acc.id,
    name: acc.name,
    handle: acc.handle,
    postMethod: acc.postMethod || 'browser',
    loginId: acc.loginId || '',
    loginPassword: acc.loginPassword || '',
    excelFileName: acc.excelFileName.endsWith('.xlsx') ? acc.excelFileName : `${acc.excelFileName}.xlsx`,
    minIntervalSec: acc.intervalMinMinutes ? acc.intervalMinMinutes * 60 : acc.minIntervalSec,
    maxIntervalSec: acc.intervalMaxMinutes ? acc.intervalMaxMinutes * 60 : acc.maxIntervalSec,
    intervalMinMinutes: acc.intervalMinMinutes || 120,
    intervalMaxMinutes: acc.intervalMaxMinutes || 360,
    activeTimeStart: acc.activeTimeStart || '0:00',
    activeTimeEnd: acc.activeTimeEnd || '0:00',
    tweetOrderMode: acc.tweetOrderMode !== undefined ? acc.tweetOrderMode : 1,
    repeatMode: acc.repeatMode,
    isActive: acc.isActive,
    actionsPerBatch: acc.actionsPerBatch || currentSettings.actionsPerRotation,
  }));
  zip.file('accounts.json', JSON.stringify(accountsJsonData, null, 2));

  // 3. アカウント情報.xlsx & 各アカウント別 テーマExcel
  zip.file('アカウント情報.xlsx', await accountsInfoToExcelBinary(accounts));
  for (const acc of accounts) {
    const fileName = acc.excelFileName.endsWith('.xlsx') ? acc.excelFileName : `${acc.excelFileName}.xlsx`;
    const excelBinary = await accountMultiTabToBinary(acc);
    zip.file(fileName, excelBinary);
  }

  // 4. 必要な空フォルダのみ作成（不要なREADME.txt等は置かない）
  zip.folder('images');
  zip.folder('browser_sessions');
  zip.folder('log');
  zip.folder('screenshots');

  // ZIP生成
  const content = await zip.generateAsync({ type: 'blob' });

  // ダウンロードトリガー
  const url = URL.createObjectURL(content);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'x_auto_tweet_browser_bot.zip';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
