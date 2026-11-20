import React, { useState } from 'react';
import {
  Code2,
  Download,
  Copy,
  Check,
  FileSpreadsheet,
  Globe,
  Cpu,
  RefreshCw,
  Zap,
  Sliders,
  MonitorPlay,
  ShieldOff,
} from 'lucide-react';
import { AccountConfig, BotOperationSettings, DmTargetItem } from '../types';
import autoTweetPlaywrightRaw from '../../auto_tweet_playwright.py?raw';
import xBotGuiPyRaw from '../../x_bot_gui.py?raw';
import xBotGuiVbsRaw from '../../X_BOT_GUI.vbs?raw';
import stopBotVbsRaw from '../../STOP_BOT.vbs?raw';
import updatePs1Raw from '../../update.ps1?raw';
import updateThemesExcelRaw from '../../update_themes_excel.py?raw';
import { BROWSER_REQUIREMENTS, DEFAULT_SETTINGS } from '../data/defaultData';
import { generatePythonProjectZip } from '../utils/zipHelper';
import { downloadAccountExcel } from '../utils/excelHelper';

interface PythonExporterProps {
  accounts: AccountConfig[];
  settings?: BotOperationSettings;
  dmTargets?: DmTargetItem[];
}

export const PythonExporter: React.FC<PythonExporterProps> = ({
  accounts,
  settings = DEFAULT_SETTINGS,
  dmTargets = [],
}) => {
  const [activeFile, setActiveFile] = useState<
    'guivbs' | 'stopvbs' | 'guipy' | 'playwright' | 'updatethemes' | 'updateps1' | 'settingsjson' | 'json' | 'req'
  >('guivbs');
  const [copied, setCopied] = useState(false);
  const [isZipping, setIsZipping] = useState(false);

  const accountsJsonString = JSON.stringify(
    accounts.map((acc) => ({
      id: acc.id,
      name: acc.name,
      handle: acc.handle,
      postMethod: acc.postMethod || 'browser',
      loginId: acc.loginId || '',
      loginPassword: acc.loginPassword || '',
      profileDirName: acc.profileDirName || acc.handle.replace(/[@]/g, ''),
      excelFileName: acc.excelFileName.endsWith('.xlsx') ? acc.excelFileName : `${acc.excelFileName}.xlsx`,
      minIntervalSec: acc.minIntervalSec,
      maxIntervalSec: acc.maxIntervalSec,
      actionsPerBatch: acc.actionsPerBatch || settings.actionsPerRotation,
      repeatMode: acc.repeatMode,
      isActive: acc.isActive,
    })),
    null,
    2
  );

  const settingsJsonString = JSON.stringify(settings, null, 2);

  const getActiveCode = () => {
    switch (activeFile) {
      case 'guivbs':
        return xBotGuiVbsRaw;
      case 'stopvbs':
        return stopBotVbsRaw;
      case 'guipy':
        return xBotGuiPyRaw;
      case 'playwright':
        return autoTweetPlaywrightRaw;
      case 'updatethemes':
        return updateThemesExcelRaw;
      case 'updateps1':
        return updatePs1Raw;
      case 'settingsjson':
        return settingsJsonString;
      case 'json':
        return accountsJsonString;
      case 'req':
        return BROWSER_REQUIREMENTS;
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(getActiveCode());
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleDownloadZip = async () => {
    setIsZipping(true);
    try {
      await generatePythonProjectZip(accounts, settings, dmTargets);
    } catch (err: any) {
      alert(`ZIP生成エラー: ${err.message}`);
    } finally {
      setIsZipping(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* Header Banner */}
      <div className="bg-gradient-to-r from-neutral-900 via-neutral-900 to-sky-950/40 border border-neutral-800 rounded-2xl p-5 sm:p-6 shadow-sm">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-sky-500/20 text-sky-400 border border-sky-500/30">
                必要ファイルのみ厳選構成
              </span>
              <span className="text-xs text-neutral-400">
                Playwright 完全無料・API不要・GUI統合
              </span>
            </div>
            <h2 className="text-xl font-bold text-white tracking-tight flex items-center gap-2">
              <span>全自動実行コード ＆ 起動パック</span>
            </h2>
            <p className="text-neutral-400 text-xs max-w-2xl leading-relaxed">
              不要なランチャー類を省き、<strong>X_BOT_GUI.vbs</strong>（起動・停止・Chromeプロフィール・Excel管理）と
              <strong>アカウント別ログ（log/&lt;アカウント名&gt;.log）＆スクショ（screenshots/&lt;アカウント名&gt;/）</strong>に対応した最新版パックです。
            </p>
          </div>

          <div className="flex items-center gap-3">
            <button
              onClick={handleDownloadZip}
              disabled={isZipping}
              className="flex items-center gap-2 px-5 py-3 rounded-xl bg-gradient-to-r from-sky-400 to-teal-400 hover:from-sky-300 hover:to-teal-300 text-neutral-950 font-bold text-sm shadow-lg transition cursor-pointer disabled:opacity-50"
            >
              <Download className="w-5 h-5" />
              <span>{isZipping ? 'ZIP作成中...' : '全自動起動パック (ZIP) をDL'}</span>
            </button>
          </div>
        </div>

        {/* 3-way features card */}
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mt-6 pt-6 border-t border-neutral-800/80 text-xs">
          <div className="bg-neutral-950/80 p-3.5 rounded-xl border border-sky-500/30 space-y-1">
            <span className="font-bold text-sky-400 flex items-center gap-1.5 text-xs">
              <MonitorPlay className="w-3.5 h-3.5" />
              ① X_BOT_GUI.vbs へ完全統合
            </span>
            <p className="text-neutral-400 text-[11px] leading-relaxed">
              ボット開始・安全停止・Chrome手動起動・Excel更新がGUI1つで完結。重複VBS/BATを廃止しました。
            </p>
          </div>
          <div className="bg-neutral-950/80 p-3.5 rounded-xl border border-emerald-500/30 space-y-1">
            <span className="font-bold text-emerald-400 flex items-center gap-1.5 text-xs">
              <Cpu className="w-3.5 h-3.5" />
              ② アカウント別ログ＆スクショ保存
            </span>
            <p className="text-neutral-400 text-[11px] leading-relaxed">
              <code className="text-emerald-300">log/アカウント名.log</code> と{' '}
              <code className="text-emerald-300">screenshots/アカウント名/</code> へ自動振り分け保存します。
            </p>
          </div>
          <div className="bg-neutral-950/80 p-3.5 rounded-xl border border-purple-500/30 space-y-1">
            <span className="font-bold text-purple-400 flex items-center gap-1.5 text-xs">
              <RefreshCw className="w-3.5 h-3.5" />
              ③ 1窓交代ローテーション＆メディア遮断
            </span>
            <p className="text-neutral-400 text-[11px] leading-relaxed">
              テーマ別シートと投稿時間フィルタに対応し、画像・動画読み込みをブロックして省メモリ稼働します。
            </p>
          </div>
        </div>
      </div>

      {/* Code Viewer */}
      <div className="bg-neutral-900 border border-neutral-800 rounded-2xl overflow-hidden shadow-sm">
        <div className="bg-neutral-950 border-b border-neutral-800 px-4 py-2.5 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-1.5 overflow-x-auto max-w-full pb-1 sm:pb-0">
            <button
              onClick={() => setActiveFile('guivbs')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-mono transition cursor-pointer ${
                activeFile === 'guivbs'
                  ? 'bg-neutral-800 text-sky-400 border border-neutral-700 shadow-sm font-semibold'
                  : 'text-neutral-400 hover:text-white'
              }`}
            >
              <MonitorPlay className="w-3.5 h-3.5 text-sky-400" />
              X_BOT_GUI.vbs
            </button>
            <button
              onClick={() => setActiveFile('stopvbs')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-mono transition cursor-pointer ${
                activeFile === 'stopvbs'
                  ? 'bg-neutral-800 text-rose-400 border border-neutral-700 shadow-sm font-semibold'
                  : 'text-neutral-400 hover:text-white'
              }`}
            >
              <ShieldOff className="w-3.5 h-3.5 text-rose-400" />
              STOP_BOT.vbs
            </button>
            <button
              onClick={() => setActiveFile('guipy')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-mono transition cursor-pointer ${
                activeFile === 'guipy'
                  ? 'bg-neutral-800 text-emerald-400 border border-neutral-700 shadow-sm font-semibold'
                  : 'text-neutral-400 hover:text-white'
              }`}
            >
              <Globe className="w-3.5 h-3.5 text-emerald-400" />
              x_bot_gui.py
            </button>
            <button
              onClick={() => setActiveFile('playwright')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-mono transition cursor-pointer ${
                activeFile === 'playwright'
                  ? 'bg-neutral-800 text-cyan-400 border border-neutral-700 shadow-sm font-semibold'
                  : 'text-neutral-400 hover:text-white'
              }`}
            >
              <Globe className="w-3.5 h-3.5 text-cyan-400" />
              auto_tweet_playwright.py
            </button>
            <button
              onClick={() => setActiveFile('updatethemes')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-mono transition cursor-pointer ${
                activeFile === 'updatethemes'
                  ? 'bg-neutral-800 text-sky-400 border border-neutral-700 shadow-sm font-semibold'
                  : 'text-neutral-400 hover:text-white'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-sky-400" />
              update_themes_excel.py
            </button>
            <button
              onClick={() => setActiveFile('updateps1')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-mono transition cursor-pointer ${
                activeFile === 'updateps1'
                  ? 'bg-neutral-800 text-emerald-400 border border-neutral-700 shadow-sm font-semibold'
                  : 'text-neutral-400 hover:text-white'
              }`}
            >
              <Zap className="w-3.5 h-3.5 text-emerald-400" />
              update.ps1
            </button>
            <button
              onClick={() => setActiveFile('settingsjson')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-mono transition cursor-pointer ${
                activeFile === 'settingsjson'
                  ? 'bg-neutral-800 text-teal-400 border border-neutral-700 shadow-sm font-semibold'
                  : 'text-neutral-400 hover:text-white'
              }`}
            >
              <Sliders className="w-3.5 h-3.5 text-teal-400" />
              settings.json
            </button>
            <button
              onClick={() => setActiveFile('json')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-mono transition cursor-pointer ${
                activeFile === 'json'
                  ? 'bg-neutral-800 text-amber-400 border border-neutral-700 shadow-sm font-semibold'
                  : 'text-neutral-400 hover:text-white'
              }`}
            >
              <Code2 className="w-3.5 h-3.5 text-amber-400" />
              accounts.json
            </button>
          </div>

          <button
            onClick={handleCopy}
            className="flex items-center gap-1.5 px-3 py-1 bg-neutral-800 hover:bg-neutral-700 text-neutral-300 rounded-lg text-xs font-mono transition cursor-pointer flex-shrink-0"
          >
            {copied ? (
              <>
                <Check className="w-3.5 h-3.5 text-emerald-400" />
                <span className="text-emerald-400">コピー完了！</span>
              </>
            ) : (
              <>
                <Copy className="w-3.5 h-3.5" />
                <span>コードをコピー</span>
              </>
            )}
          </button>
        </div>

        <div className="p-4 sm:p-5 overflow-x-auto bg-neutral-950 font-mono text-xs leading-relaxed text-neutral-300 max-h-[550px] overflow-y-auto">
          <pre className="whitespace-pre">{getActiveCode()}</pre>
        </div>
      </div>

      {/* Excel Download shortcuts for each account */}
      <div className="bg-neutral-900 border border-neutral-800 rounded-xl p-4">
        <h4 className="text-xs font-semibold text-neutral-300 mb-3 flex items-center gap-2">
          <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
          Excelファイル個別ダウンロード
        </h4>
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
          {accounts.map((acc) => (
            <button
              key={acc.id}
              onClick={() => downloadAccountExcel(acc.tweets, acc.excelFileName)}
              className="flex items-center justify-between p-3 rounded-lg bg-neutral-950/60 border border-neutral-800 hover:border-neutral-700 text-left transition cursor-pointer group"
            >
              <div className="min-w-0">
                <span className="text-xs font-semibold text-white block truncate">{acc.name}</span>
                <span className="text-[11px] font-mono text-emerald-400 block truncate">{acc.excelFileName}</span>
              </div>
              <Download className="w-4 h-4 text-neutral-500 group-hover:text-emerald-400 transition ml-2 flex-shrink-0" />
            </button>
          ))}
        </div>
      </div>
    </div>
  );
};
