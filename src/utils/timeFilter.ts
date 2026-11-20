/**
 * ポスト別「投稿時間」およびテーマ別「実行時間」「曜日」の柔軟な判定ユーティリティ
 * auto_tweet_playwright.py の判定ロジックと完全一致:
 *   - 7～10, ７～１０, 7〜10, 7-10, 7時から10時, 7時～10時, 7:00-10:00 -> 7:00〜10:00 (6:59や10:01は対象外)
 *   - 日またぎ: 22～2 (22:00〜翌2:00)
 *   - 複数時間帯: 7～10, 18～21 (カンマ・読点区切り)
 *   - 空白 / 0:00-0:00 -> true
 */

function normalizeNFKC(input: string): string {
  return (input || '').normalize('NFKC').trim();
}

export function parseSingleTimeToMinutes(token: string): number | null {
  let s = normalizeNFKC(token).replace(/\s+/g, '');
  if (!s) return null;

  const isPm = /午後|pm/i.test(s);
  s = s.replace(/(翌|午前|午後|am|pm)/gi, '');

  const match = s.match(/^(\d{1,2})\s*(?::(\d{1,2})|時(?:(半)|(\d{1,2})分?)?)?$/);
  if (!match) return null;

  let hour = parseInt(match[1], 10);
  let minute = 0;
  if (match[2] !== undefined) {
    minute = parseInt(match[2], 10);
  } else if (match[3] === '半') {
    minute = 30;
  } else if (match[4] !== undefined) {
    minute = parseInt(match[4], 10);
  }

  if (isPm && hour >= 1 && hour <= 11) {
    hour += 12;
  }

  if (hour === 24 && minute === 0) return 24 * 60;
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return null;
  return hour * 60 + minute;
}

export function isTimeInRanges(timeSpec?: string, nowDate: Date = new Date()): boolean {
  const raw = normalizeNFKC(timeSpec || '');
  if (!raw || ['none', 'nan', '終日', '24時間', 'いつでも', '常時', 'all'].includes(raw.toLowerCase())) {
    return true;
  }

  const curMin = nowDate.getHours() * 60 + nowDate.getMinutes();
  const segments = raw
    .split(/[,、;\n/]+/)
    .map((s) => s.trim())
    .filter(Boolean);

  if (segments.length === 0) return true;

  let validRangeFound = false;
  for (const seg of segments) {
    const unified = seg.replace(/(?:から|〜|～|~|−|–|—|ー|-|\bto\b)/gi, '~');
    const parts = unified
      .split('~')
      .map((p) => p.trim())
      .filter(Boolean);

    if (parts.length < 2) {
      const singleMin = parts[0] ? parseSingleTimeToMinutes(parts[0]) : null;
      if (singleMin !== null) {
        validRangeFound = true;
        if (curMin === singleMin) return true;
      }
      continue;
    }

    const sMin = parseSingleTimeToMinutes(parts[0]);
    const eMin = parseSingleTimeToMinutes(parts[1]);
    if (sMin === null || eMin === null) continue;

    validRangeFound = true;
    if (sMin === 0 && eMin === 0) return true;

    if (sMin <= eMin) {
      if (curMin >= sMin && curMin <= eMin) return true;
    } else {
      if (curMin >= sMin || curMin <= eMin) return true;
    }
  }

  return validRangeFound ? false : true;
}

export function isTodayAllowed(weekdaySpec?: string, nowDate: Date = new Date()): boolean {
  const raw = normalizeNFKC(weekdaySpec || '');
  if (!raw || ['毎日', '全日', 'すべて', 'all', '*', 'none', 'nan'].includes(raw.toLowerCase())) {
    return true;
  }

  // JS getDay(): 0=Sun..6=Sat -> Python weekday: 0=Mon..6=Sun
  const jsDay = nowDate.getDay();
  const wd = jsDay === 0 ? 6 : jsDay - 1;

  if (/平日|月~金|月-金|月〜金/.test(raw)) return wd < 5;
  if (/土日|週末/.test(raw)) return wd >= 5;

  const jpDays = ['月', '火', '水', '木', '金', '土', '日'];
  const enDays = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];
  const low = raw.toLowerCase();

  const hasToken = jpDays.some((d) => raw.includes(d)) || enDays.some((d) => low.includes(d));
  if (!hasToken) return true;

  return raw.includes(jpDays[wd]) || low.includes(enDays[wd]);
}
