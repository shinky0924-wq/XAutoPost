import React, { useState, useRef, useMemo } from 'react';
import {
  Plus,
  Copy,
  Trash2,
  Download,
  Upload,
  Image as ImageIcon,
  RotateCcw,
  Clock,
  ExternalLink,
  FileSpreadsheet,
  Play,
  Info,
  Edit2,
  Check,
} from 'lucide-react';
import { AccountConfig, ThemeSheet, TweetItem } from '../types';
import {
  DEFAULT_THEME_MEMO,
  downloadAccountMultiTabExcel,
  getAccountThemes,
  parseExcelFile,
} from '../utils/excelHelper';
import { isTimeInRanges } from '../utils/timeFilter';

interface ExcelSheetEditorProps {
  account: AccountConfig;
  onUpdateAccount?: (updated: AccountConfig) => void;
  onUpdateTweets: (tweets: TweetItem[]) => void;
  onUpdateFileName: (fileName: string) => void;
  onNavigateToSimulator?: () => void;
}

export const ExcelSheetEditor: React.FC<ExcelSheetEditorProps> = ({
  account,
  onUpdateAccount,
  onUpdateTweets,
  onUpdateFileName,
  onNavigateToSimulator,
}) => {
  const themes = useMemo(() => getAccountThemes(account), [account]);
  const [activeThemeId, setActiveThemeId] = useState<string>(themes[0]?.id || 'theme-1');
  const [renamingThemeId, setRenamingThemeId] = useState<string | null>(null);
  const [renameDraft, setRenameDraft] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);

  const activeTheme = themes.find((t) => t.id === activeThemeId) || themes[0];

  const calculateTwitterLength = (text: string): number => {
    let len = 0;
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i);
      if ((code >= 0x0020 && code <= 0x007e) || (code >= 0xff61 && code <= 0xff9f)) {
        len += 0.5;
      } else {
        len += 1;
      }
    }
    return Math.ceil(len);
  };

  const commitThemes = (nextThemes: ThemeSheet[]) => {
    if (onUpdateAccount) {
      onUpdateAccount({
        ...account,
        themes: nextThemes,
        tweets: nextThemes[0]?.tweets || account.tweets,
      });
    } else if (nextThemes[0]) {
      onUpdateTweets(nextThemes[0].tweets);
    }
  };

  const handleUpdateActiveTheme = (patch: Partial<ThemeSheet>) => {
    const nextThemes = themes.map((t) => (t.id === activeTheme.id ? { ...t, ...patch } : t));
    commitThemes(nextThemes);
  };

  const handleAddTheme = () => {
    const nextNum = themes.length + 1;
    const newTheme: ThemeSheet = {
      id: `theme-${Date.now()}`,
      name: `テーマ${nextNum}`,
      activeTimeRange: '0:00-0:00',
      weekdays: '毎日',
      intervalRange: '120~360分',
      tweetOrderMode: 1,
      isActive: true,
      memo: DEFAULT_THEME_MEMO,
      tweets: [
        {
          id: 1,
          text: '',
          post_time: '',
          image_path: '',
          status: 'pending',
        },
      ],
    };
    commitThemes([...themes, newTheme]);
    setActiveThemeId(newTheme.id);
  };

  const handleDuplicateTheme = () => {
    const clonedTweets = activeTheme.tweets.map((t, idx) => ({
      ...t,
      id: idx + 1,
      status: 'pending' as const,
      posted_at: '',
    }));
    const newTheme: ThemeSheet = {
      ...activeTheme,
      id: `theme-${Date.now()}`,
      name: `${activeTheme.name}_コピー`,
      tweets: clonedTweets,
    };
    commitThemes([...themes, newTheme]);
    setActiveThemeId(newTheme.id);
  };

  const handleDeleteTheme = (id: string) => {
    if (themes.length <= 1) {
      alert('最低1つのテーマシートが必要です。');
      return;
    }
    const remaining = themes.filter((t) => t.id !== id);
    commitThemes(remaining);
    if (activeTheme.id === id) {
      setActiveThemeId(remaining[0].id);
    }
  };

  const handleStartRename = (theme: ThemeSheet) => {
    setRenamingThemeId(theme.id);
    setRenameDraft(theme.name);
  };

  const handleFinishRename = () => {
    if (renamingThemeId && renameDraft.trim()) {
      const nextThemes = themes.map((t) =>
        t.id === renamingThemeId ? { ...t, name: renameDraft.trim() } : t
      );
      commitThemes(nextThemes);
    }
    setRenamingThemeId(null);
  };

  // 投稿行操作
  const handleAddRow = () => {
    const curTweets = activeTheme.tweets || [];
    const nextId = curTweets.length > 0 ? Math.max(...curTweets.map((t) => t.id)) + 1 : 1;
    const newTweet: TweetItem = {
      id: nextId,
      text: '',
      post_time: '',
      image_path: '',
      status: 'pending',
    };
    handleUpdateActiveTheme({ tweets: [...curTweets, newTweet] });
  };

  const handleDeleteRow = (id: number) => {
    handleUpdateActiveTheme({
      tweets: activeTheme.tweets.filter((t) => t.id !== id),
    });
  };

  const handleCellChange = (id: number, field: keyof TweetItem, value: any) => {
    handleUpdateActiveTheme({
      tweets: activeTheme.tweets.map((t) => (t.id === id ? { ...t, [field]: value } : t)),
    });
  };

  const handleResetAllToPending = () => {
    if (window.confirm(`テーマ「${activeTheme.name}」の全ステータスを未投稿 (pending) にリセットしますか？`)) {
      handleUpdateActiveTheme({
        tweets: activeTheme.tweets.map((t) => ({
          ...t,
          status: 'pending' as const,
          posted_at: '',
        })),
      });
    }
  };

  const handleDownloadExcel = () => {
    downloadAccountMultiTabExcel(account, account.excelFileName);
  };

  const applyImportedFile = async (file: File) => {
    const res = await parseExcelFile(file);
    if (res.themes.length === 0 && res.tweets.length === 0) {
      alert('有効なExcelデータが見つかりませんでした。');
      return;
    }
    if (onUpdateAccount && res.themes.length > 0) {
      onUpdateAccount({
        ...account,
        themes: res.themes,
        tweets: res.themes[0].tweets,
        intervalMinMinutes: res.settings?.intervalMinMinutes ?? account.intervalMinMinutes,
        intervalMaxMinutes: res.settings?.intervalMaxMinutes ?? account.intervalMaxMinutes,
        activeTimeStart: res.settings?.activeTimeStart ?? account.activeTimeStart,
        activeTimeEnd: res.settings?.activeTimeEnd ?? account.activeTimeEnd,
        tweetOrderMode: res.settings?.tweetOrderMode ?? account.tweetOrderMode,
        excelFileName: file.name,
      });
      setActiveThemeId(res.themes[0].id);
    } else if (res.tweets.length > 0) {
      onUpdateTweets(res.tweets);
      onUpdateFileName(file.name);
    }
    const totalLoaded = res.themes.reduce((sum, th) => sum + th.tweets.length, 0);
    alert(`「${file.name}」から ${res.themes.length} テーマ（合計 ${totalLoaded} 件）を読み込みました！`);
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      await applyImportedFile(file);
    } catch (err: any) {
      alert(`Excelファイルの読み込みに失敗しました: ${err.message}`);
    }
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      try {
        await applyImportedFile(file);
      } catch (err: any) {
        alert(`ファイル読み込みエラー: ${err.message}`);
      }
    }
  };

  const totalTweetsCount = themes.reduce((sum, th) => sum + th.tweets.length, 0);
  const now = new Date();

  return (
    <div className="space-y-4">
      {/* Top action header */}
      <div className="bg-neutral-900 border border-neutral-800 rounded-xl p-4 sm:p-5">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <FileSpreadsheet className="w-5 h-5 text-emerald-400" />
              <h3 className="text-base font-semibold text-white">
                テーマ別 Excel 管理（色付き・削除耐性書式対応）
              </h3>
              <span className="text-xs bg-sky-500/10 text-sky-400 border border-sky-500/20 px-2 py-0.5 rounded-full font-mono">
                {themes.length} テーマ / 計 {totalTweetsCount} 件
              </span>
            </div>
            <div className="flex items-center gap-2 mt-1.5">
              <span className="text-xs text-neutral-400">対象ファイル名:</span>
              <input
                id="input-excel-filename"
                type="text"
                value={account.excelFileName}
                onChange={(e) => onUpdateFileName(e.target.value)}
                className="text-xs font-mono bg-neutral-950 border border-neutral-800 rounded px-2 py-1 text-sky-400 w-48 focus:border-sky-500 focus:outline-none"
              />
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {onNavigateToSimulator && (
              <button
                id="btn-excel-test-now"
                type="button"
                onClick={onNavigateToSimulator}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-sky-500 hover:bg-sky-400 text-neutral-950 text-xs font-bold transition cursor-pointer shadow-sm"
              >
                <Play className="w-3.5 h-3.5 fill-current" />
                テスト実行
              </button>
            )}

            <button
              id="btn-add-tweet-row"
              onClick={handleAddRow}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-medium transition cursor-pointer"
            >
              <Plus className="w-4 h-4" />
              ツイート追加
            </button>

            <button
              id="btn-download-excel"
              onClick={handleDownloadExcel}
              className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-emerald-700 hover:bg-emerald-600 text-white font-semibold text-xs transition cursor-pointer shadow-sm"
              title="色分け・200行事前書式・条件付き書式付きでExcel出力"
            >
              <Download className="w-4 h-4 text-emerald-200" />
              {account.excelFileName || `${account.name}.xlsx`} をDL
            </button>

            <label className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-neutral-800 hover:bg-neutral-700 text-neutral-200 border border-neutral-700 text-xs font-medium transition cursor-pointer">
              <Upload className="w-4 h-4 text-sky-400" />
              Excel取込
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx, .xls"
                className="hidden"
                onChange={handleFileUpload}
              />
            </label>

            <button
              id="btn-reset-pending"
              onClick={handleResetAllToPending}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-neutral-800/80 hover:bg-neutral-700 text-neutral-400 hover:text-neutral-200 text-xs transition cursor-pointer"
              title="このテーマの全ツイートを未投稿にリセット"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              リセット
            </button>
          </div>
        </div>

        <div
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={handleDrop}
          className={`mt-3 p-2 rounded-lg border border-dashed text-center text-xs transition ${
            dragOver
              ? 'border-emerald-500 bg-emerald-500/10 text-emerald-300'
              : 'border-neutral-800 text-neutral-500 hover:border-neutral-700'
          }`}
        >
          既存の「{account.excelFileName}」をここにドラッグ＆ドロップすると、全テーマシートを自動読み込みします
        </div>
      </div>

      {/* Theme Sheet Tabs Bar (Excelの下シートを再現) */}
      <div className="flex items-center justify-between gap-2 flex-wrap bg-neutral-900 border border-neutral-800 rounded-xl p-2">
        <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
          {themes.map((th) => {
            const isSelected = th.id === activeTheme.id;
            const isRenaming = renamingThemeId === th.id;

            return (
              <div
                key={th.id}
                onClick={() => !isRenaming && setActiveThemeId(th.id)}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer border ${
                  isSelected
                    ? 'bg-[#1E3A8A] text-white border-blue-500 shadow-sm'
                    : 'bg-neutral-950 text-neutral-400 hover:text-neutral-200 border-neutral-800'
                }`}
              >
                <span
                  className={`w-2 h-2 rounded-full flex-shrink-0 ${
                    th.isActive ? 'bg-emerald-400' : 'bg-rose-400'
                  }`}
                  title={th.isActive ? '有効' : '無効'}
                />

                {isRenaming ? (
                  <div className="flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
                    <input
                      type="text"
                      value={renameDraft}
                      onChange={(e) => setRenameDraft(e.target.value)}
                      onKeyDown={(e) => e.key === 'Enter' && handleFinishRename()}
                      autoFocus
                      className="w-24 bg-neutral-950 text-white px-1.5 py-0.5 rounded border border-sky-400 text-xs focus:outline-none"
                    />
                    <button
                      type="button"
                      onClick={handleFinishRename}
                      className="text-emerald-300 hover:text-white cursor-pointer"
                    >
                      <Check className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ) : (
                  <>
                    <span>{th.name}</span>
                    <span className="text-[10px] opacity-75 font-mono">({th.tweets.length})</span>
                    {isSelected && (
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleStartRename(th);
                        }}
                        className="ml-0.5 text-blue-200 hover:text-white cursor-pointer"
                        title="テーマ名を変更"
                      >
                        <Edit2 className="w-3 h-3" />
                      </button>
                    )}
                    {themes.length > 1 && (
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleDeleteTheme(th.id);
                        }}
                        className="ml-0.5 text-neutral-400 hover:text-rose-300 cursor-pointer"
                        title="このテーマを削除"
                      >
                        <Trash2 className="w-3 h-3" />
                      </button>
                    )}
                  </>
                )}
              </div>
            );
          })}
        </div>

        <div className="flex items-center gap-1.5">
          <button
            type="button"
            onClick={handleAddTheme}
            className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-sky-500/15 hover:bg-sky-500/25 text-sky-300 border border-sky-500/30 text-xs font-semibold transition cursor-pointer"
          >
            <Plus className="w-3.5 h-3.5" />
            テーマ追加
          </button>
          <button
            type="button"
            onClick={handleDuplicateTheme}
            className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-neutral-800 hover:bg-neutral-700 text-neutral-300 border border-neutral-700 text-xs font-semibold transition cursor-pointer"
          >
            <Copy className="w-3.5 h-3.5" />
            複製
          </button>
        </div>
      </div>

      {/* Excel Rows 1-2 Visual Banner (Matches user's attached screenshot 1:1) */}
      <div className="bg-neutral-900 border border-neutral-800 rounded-xl overflow-hidden shadow-md">
        <div className="overflow-x-auto">
          <table className="w-full text-xs border-collapse min-w-[780px]">
            <thead>
              <tr className="bg-[#1E3A8A] text-white font-bold text-center">
                <th className="py-2 px-3 border border-blue-500/50 w-32">実行時間</th>
                <th className="py-2 px-3 border border-blue-500/50 w-32">曜日</th>
                <th className="py-2 px-3 border border-blue-500/50 w-36">ポスト間隔</th>
                <th className="py-2 px-3 border border-blue-500/50 w-36">投稿モード</th>
                <th className="py-2 px-3 border border-blue-500/50 w-28">有効/無効</th>
                <th className="py-2 px-3 border border-blue-500/50">メモ</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="p-1.5 bg-[#E8F1FF] border border-blue-300 text-center">
                  <input
                    type="text"
                    value={activeTheme.activeTimeRange}
                    onChange={(e) => handleUpdateActiveTheme({ activeTimeRange: e.target.value })}
                    placeholder="0:00-0:00"
                    className="w-full bg-transparent text-center font-bold text-[#1E3A8A] focus:bg-white focus:outline-none rounded py-1 px-1"
                  />
                </td>
                <td className="p-1.5 bg-[#E8F1FF] border border-blue-300 text-center">
                  <input
                    type="text"
                    value={activeTheme.weekdays}
                    onChange={(e) => handleUpdateActiveTheme({ weekdays: e.target.value })}
                    placeholder="毎日"
                    className="w-full bg-transparent text-center font-bold text-[#1E3A8A] focus:bg-white focus:outline-none rounded py-1 px-1"
                  />
                </td>
                <td className="p-1.5 bg-[#E8F1FF] border border-blue-300 text-center">
                  <input
                    type="text"
                    value={activeTheme.intervalRange}
                    onChange={(e) => handleUpdateActiveTheme({ intervalRange: e.target.value })}
                    placeholder="120~360分"
                    className="w-full bg-transparent text-center font-bold text-[#1E3A8A] focus:bg-white focus:outline-none rounded py-1 px-1"
                  />
                </td>
                <td className="p-1.5 bg-[#E8F1FF] border border-blue-300 text-center">
                  <select
                    value={activeTheme.tweetOrderMode}
                    onChange={(e) => handleUpdateActiveTheme({ tweetOrderMode: Number(e.target.value) })}
                    className="w-full bg-transparent text-center font-bold text-[#1E3A8A] focus:bg-white focus:outline-none rounded py-1 cursor-pointer"
                  >
                    <option value={1}>1: ランダム</option>
                    <option value={0}>0: 上から順</option>
                  </select>
                </td>
                <td
                  className={`p-1.5 border border-blue-300 text-center ${
                    activeTheme.isActive ? 'bg-[#DCFCE7]' : 'bg-[#FEE2E2]'
                  }`}
                >
                  <select
                    value={activeTheme.isActive ? '有効' : '無効'}
                    onChange={(e) => handleUpdateActiveTheme({ isActive: e.target.value === '有効' })}
                    className={`w-full bg-transparent text-center font-bold focus:outline-none rounded py-1 cursor-pointer ${
                      activeTheme.isActive ? 'text-[#15803D]' : 'text-[#B91C1C]'
                    }`}
                  >
                    <option value="有効">有効</option>
                    <option value="無効">無効</option>
                  </select>
                </td>
                <td className="p-2 bg-[#F8FAFC] border border-blue-300 text-[11px] text-slate-600 leading-snug">
                  <input
                    type="text"
                    value={activeTheme.memo || DEFAULT_THEME_MEMO}
                    onChange={(e) => handleUpdateActiveTheme({ memo: e.target.value })}
                    className="w-full bg-transparent text-slate-600 focus:bg-white focus:outline-none rounded px-1 py-0.5"
                  />
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      {/* Excel Rows 4+: Post Table with 投稿時間 immediately right of ツイート文章 */}
      <div className="bg-neutral-900 border border-neutral-800 rounded-xl overflow-hidden shadow-sm">
        <div className="p-3 bg-neutral-950/90 border-b border-neutral-800 flex flex-wrap items-center justify-between gap-2">
          <span className="text-xs font-bold text-neutral-200 flex items-center gap-2">
            <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
            テーマ「{activeTheme.name}」の投稿リスト ({activeTheme.tweets.length} 件)
          </span>
          <div className="flex items-center gap-2 text-[11px] text-amber-300/90">
            <Info className="w-3.5 h-3.5 text-amber-400 flex-shrink-0" />
            <span>
              投稿時間: <code className="bg-neutral-800 px-1 rounded">7～10</code> /{' '}
              <code className="bg-neutral-800 px-1 rounded">22～2</code> /{' '}
              <code className="bg-neutral-800 px-1 rounded">7～10, 18～21</code> / 空白=通常稼働（時間外は即スキップ）
            </span>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs text-neutral-300 border-collapse">
            <thead className="bg-[#1E40AF] text-white border-b border-neutral-800 text-[11px] tracking-wider">
              <tr>
                <th className="py-2.5 px-3 w-12 text-center">No</th>
                <th className="py-2.5 px-4 min-w-[290px]">ツイート文章</th>
                <th className="py-2.5 px-3 min-w-[160px] bg-[#B45309] text-white text-center">
                  投稿時間
                </th>
                <th className="py-2.5 px-4 min-w-[190px]">画像パス</th>
                <th className="py-2.5 px-3 w-28 text-center">ステータス</th>
                <th className="py-2.5 px-3 w-36 text-center">投稿日時</th>
                <th className="py-2.5 px-3 w-14 text-center">操作</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-neutral-800/60 font-sans">
              {activeTheme.tweets.length === 0 ? (
                <tr>
                  <td colSpan={7} className="py-8 text-center text-neutral-500">
                    登録されたツイートがありません。「ツイート追加」ボタンから登録してください。
                  </td>
                </tr>
              ) : (
                activeTheme.tweets.map((tweet, idx) => {
                  const charCount = calculateTwitterLength(tweet.text);
                  const isOverLimit = charCount > 140;
                  const hasPostTime = Boolean(tweet.post_time && tweet.post_time.trim());
                  const matchesNow = isTimeInRanges(tweet.post_time, now);

                  return (
                    <tr
                      key={tweet.id}
                      className={`hover:bg-neutral-800/40 transition group ${
                        tweet.status === 'posted' ? 'opacity-65' : ''
                      }`}
                    >
                      <td className="py-3 px-3 text-center font-mono text-neutral-400">
                        {idx + 1}
                      </td>

                      {/* ツイート文章 */}
                      <td className="py-2.5 px-4 align-top">
                        <textarea
                          rows={2}
                          value={tweet.text}
                          onChange={(e) => handleCellChange(tweet.id, 'text', e.target.value)}
                          placeholder="ツイート本文を入力（{A|B} や {月日}{曜日} タグも利用可）"
                          className="w-full bg-neutral-950/80 border border-neutral-800 rounded-lg p-2 text-xs text-neutral-100 placeholder-neutral-600 focus:border-sky-500 focus:outline-none resize-y leading-relaxed"
                        />
                        <div className="flex justify-end mt-1 px-1">
                          <span
                            className={`text-[10px] font-mono px-1.5 py-0.2 rounded ${
                              isOverLimit ? 'bg-rose-500/20 text-rose-400 font-bold' : 'text-neutral-400'
                            }`}
                          >
                            {charCount}/140字
                          </span>
                        </div>
                      </td>

                      {/* 投稿時間 (ツイート文章のすぐ右) */}
                      <td className="py-2.5 px-3 align-top bg-amber-950/10">
                        <div className="space-y-1">
                          <input
                            type="text"
                            value={tweet.post_time || ''}
                            onChange={(e) => handleCellChange(tweet.id, 'post_time', e.target.value)}
                            placeholder="例: 7～10 (空白=常時)"
                            className="w-full bg-neutral-950/90 border border-amber-500/30 rounded-lg py-1.5 px-2.5 text-xs text-center font-mono text-amber-200 placeholder-neutral-600 focus:border-amber-400 focus:outline-none"
                          />
                          <div className="text-center text-[10px]">
                            {!hasPostTime ? (
                              <span className="text-neutral-500">空白: 通常稼働</span>
                            ) : matchesNow ? (
                              <span className="text-emerald-400 font-medium">現在: 投稿対象内</span>
                            ) : (
                              <span className="text-amber-400 font-medium">現在: 時間外(即スキップ)</span>
                            )}
                          </div>
                        </div>
                      </td>

                      {/* 画像パス */}
                      <td className="py-2.5 px-4 align-top">
                        <div className="relative flex items-center">
                          <ImageIcon className="w-3.5 h-3.5 text-neutral-500 absolute left-2.5" />
                          <input
                            type="text"
                            value={tweet.image_path}
                            onChange={(e) => handleCellChange(tweet.id, 'image_path', e.target.value)}
                            placeholder="例: images/sample.jpg"
                            className="w-full bg-neutral-950/80 border border-neutral-800 rounded-lg py-1.5 pl-8 pr-2.5 text-xs font-mono text-neutral-200 placeholder-neutral-600 focus:border-sky-500 focus:outline-none"
                          />
                        </div>
                      </td>

                      {/* ステータス */}
                      <td className="py-3 px-3 text-center align-top">
                        <select
                          value={tweet.status}
                          onChange={(e) => handleCellChange(tweet.id, 'status', e.target.value)}
                          className={`text-xs font-medium px-2 py-1 rounded-full border cursor-pointer focus:outline-none ${
                            tweet.status === 'pending'
                              ? 'bg-amber-500/10 text-amber-400 border-amber-500/30'
                              : tweet.status === 'posted'
                              ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30'
                              : tweet.status === 'failed'
                              ? 'bg-rose-500/10 text-rose-400 border-rose-500/30'
                              : 'bg-neutral-800 text-neutral-400 border-neutral-700'
                          }`}
                        >
                          <option value="pending">未投稿</option>
                          <option value="posted">投稿済</option>
                          <option value="failed">失敗</option>
                          <option value="skipped">スキップ</option>
                        </select>
                      </td>

                      {/* 投稿日時 */}
                      <td className="py-3 px-3 font-mono text-[11px] text-neutral-400 align-top text-center">
                        {tweet.posted_at ? (
                          <div className="flex items-center justify-center gap-1 text-neutral-400">
                            <Clock className="w-3 h-3 text-neutral-500 flex-shrink-0" />
                            <span className="truncate">{tweet.posted_at}</span>
                          </div>
                        ) : (
                          <span className="text-neutral-600">-</span>
                        )}
                      </td>

                      {/* 操作 */}
                      <td className="py-3 px-3 text-center align-top">
                        <div className="flex items-center justify-center gap-1">
                          <a
                            href={`https://twitter.com/intent/tweet?text=${encodeURIComponent(tweet.text)}`}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="px-2 py-1 bg-sky-500/10 hover:bg-sky-500 text-sky-400 hover:text-neutral-950 border border-sky-500/30 rounded text-[11px] font-medium transition cursor-pointer flex items-center gap-1 whitespace-nowrap"
                            title="ブラウザでX投稿テスト"
                          >
                            <ExternalLink className="w-3 h-3" />
                            実ポスト
                          </a>
                          <button
                            onClick={() => handleDeleteRow(tweet.id)}
                            className="p-1 rounded text-neutral-500 hover:text-rose-400 hover:bg-neutral-800 transition cursor-pointer"
                            title="行を削除"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {/* Footer guide */}
        <div className="bg-neutral-950/60 p-3 px-4 border-t border-neutral-800 flex flex-col sm:flex-row items-center justify-between text-xs text-neutral-400 gap-2">
          <div className="flex items-center gap-3">
            <span>
              未投稿:{' '}
              <strong className="text-amber-400 font-mono">
                {activeTheme.tweets.filter((t) => t.status === 'pending').length}
              </strong>{' '}
              件
            </span>
            <span>
              投稿済:{' '}
              <strong className="text-emerald-400 font-mono">
                {activeTheme.tweets.filter((t) => t.status === 'posted').length}
              </strong>{' '}
              件
            </span>
          </div>
          <div className="text-[11px] text-neutral-500">
            ※ 書き出しExcelは200行目まで色・罫線・条件付き書式を適用済みのため、Delete削除しても書式が消えません。
          </div>
        </div>
      </div>
    </div>
  );
};
