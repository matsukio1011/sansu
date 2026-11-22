/**
 * 学習アプリ共通 Google Apps Script テンプレート（複数アプリ・マルチシート対応）
 *
 * 【デプロイ手順（新規作成時）】
 * 1. スプレッドシート（アプリ履歴保存）を開く
 *    https://docs.google.com/spreadsheets/d/1MEanFW4V1ldOnmuh-_9IbQmGv5TORwGiMJBaN6_7Kes/edit
 * 2. メニューの「拡張機能」→「Apps Script」を開く
 * 3. コードを貼り付けて保存（Ctrl+S）
 * 4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員
 *
 * 【既存デプロイの更新手順（URL を変えずに反映）】
 * 1. Apps Script でコードを貼り替えて保存
 * 2. 「デプロイ」→「デプロイを管理」→ 鉛筆アイコン → バージョン「新バージョン」→「デプロイ」
 *
 * 【API 仕様】
 * - GET (パラメータなし / action=roster): 「名簿」シート A列2行目以降の氏名配列 { ok: true, names: [...] }
 * - GET ?action=history&sheet=<シート名>: テスト履歴（10列、逆順・最大200件）
 *   ※ sheet 省略時は後方互換のため '単位変換履歴' を参照
 * - GET ?action=practice_history&sheet=<シート名>: 練習履歴（12列、逆順・最大200件）
 *   ※ sheet 省略時は後方互換のため '単位変換練習履歴' を参照
 * - POST (Content-Type: text/plain;charset=utf-8):
 *   JSON の data.sheet で保存先シートを指定可能（'名簿' への書き込みは拒否）。
 *   data.sheet 省略時は data.mode === 'practice' なら '単位変換練習履歴'、それ以外は '単位変換履歴'。
 */

const ROSTER_SHEET_NAME = '名簿';
const DEFAULT_TEST_HISTORY_SHEET = '単位変換履歴';
const DEFAULT_PRACTICE_HISTORY_SHEET = '単位変換練習履歴';
const MAX_CELL_CHARS = 45000;

const TEST_HISTORY_HEADERS = [
  '日時',
  '氏名',
  'コース',
  'ジャンル',
  '電卓',
  '結果',
  '合計回答数',
  '正解数',
  '不正解数',
  '所要時間(秒)'
];

const PRACTICE_HISTORY_HEADERS = [
  '日時',
  '氏名',
  'コース',
  'ジャンル',
  '電卓',
  '合計回答数',
  '正解数',
  '不正解数',
  '正答率(%)',
  '所要時間(秒)',
  '問題と解答',
  '詳細JSON'
];

function createJsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 日時セル（Date または文字列）を JST の yyyy/MM/dd HH:mm:ss 文字列へ整形する
 */
function formatTimestampCell(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  }
  return String(value).trim();
}

/**
 * 読み書き対象のシート名を解決する（名簿シートへの上書きを防止）
 */
function resolveTargetSheetName(requestedName, fallbackName) {
  const name = requestedName != null ? String(requestedName).trim() : '';
  return name || fallbackName;
}

/**
 * シートから最新 limit 件を逆順で読み出してマッピングする
 */
function readRecentRecords(ss, sheetName, colCount, mapper) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const actualCols = Math.min(colCount, Math.max(1, sheet.getLastColumn()));
  const values = sheet.getRange(2, 1, lastRow - 1, actualCols).getValues();
  const records = [];
  const limit = 200;
  for (let i = values.length - 1; i >= 0 && records.length < limit; i--) {
    const row = values[i];
    const ts = formatTimestampCell(row[0]);
    const name = row[1] == null ? '' : String(row[1]).trim();
    if (!ts && !name) continue;
    records.push(mapper(row, ts, name));
  }
  return records;
}

/**
 * 練習の回答一覧配列から人間が読める改行テキストを生成する（最大45,000文字）
 */
function buildPracticeQaText(items) {
  if (!Array.isArray(items) || items.length === 0) return '';
  const lines = [];
  for (let i = 0; i < items.length; i++) {
    const item = items[i] || {};
    const q = item.q != null ? String(item.q).trim() : '';
    const chosen = item.chosen != null ? String(item.chosen).trim() : '';
    const ans = item.ans != null ? String(item.ans).trim() : '';
    const ok = Boolean(item.isCorrect);
    if (ok) {
      lines.push((i + 1) + '. ⭕ ' + q + ' → ' + chosen);
    } else {
      lines.push((i + 1) + '. ❌ ' + q + ' → あなた: ' + chosen + '（正解: ' + ans + '）');
    }
  }
  const joined = lines.join('\n');
  if (joined.length <= MAX_CELL_CHARS) return joined;
  return joined.slice(0, MAX_CELL_CHARS - 15) + '\n…（省略）';
}

/**
 * 練習の回答一覧配列を有効な JSON 文字列へシリアライズする（最大45,000文字）
 */
function serializePracticeItems(items) {
  if (!Array.isArray(items) || items.length === 0) return '[]';
  const normalized = [];
  for (let i = 0; i < items.length; i++) {
    const item = items[i] || {};
    normalized.push({
      q: item.q != null ? String(item.q).trim() : '',
      chosen: item.chosen != null ? String(item.chosen).trim() : '',
      ans: item.ans != null ? String(item.ans).trim() : '',
      isCorrect: Boolean(item.isCorrect)
    });
  }
  const full = JSON.stringify(normalized);
  if (full.length <= MAX_CELL_CHARS) return full;

  const trimmed = [];
  for (let i = 0; i < normalized.length; i++) {
    trimmed.push(normalized[i]);
    if (JSON.stringify(trimmed).length > MAX_CELL_CHARS) {
      trimmed.pop();
      break;
    }
  }
  return JSON.stringify(trimmed);
}

/**
 * 詳細JSON文字列を安全に配列へパースする
 */
function parsePracticeItemsJson(raw) {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(String(raw));
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

/**
 * 指定シートを取得し、存在しない・空の場合はヘッダー行付きで初期化する
 */
function getOrCreateHistorySheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 名簿一覧（action=roster）、テスト履歴（action=history）、練習履歴（action=practice_history）を返す
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action ? String(params.action) : 'roster';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'history') {
      const sheetName = resolveTargetSheetName(params.sheet, DEFAULT_TEST_HISTORY_SHEET);
      if (sheetName === ROSTER_SHEET_NAME) {
        return createJsonResponse({ ok: false, error: '名簿シートは履歴として読み込めません。' });
      }
      const records = readRecentRecords(ss, sheetName, 10, function (row, ts, name) {
        return {
          timestamp: ts,
          name: name,
          course: row[2] == null ? '' : String(row[2]),
          category: row[3] == null ? '' : String(row[3]),
          calc: row[4] == null ? '' : String(row[4]),
          result: row[5] == null ? '' : String(row[5]),
          totalAnswered: Number(row[6] || 0),
          correctCount: Number(row[7] || 0),
          wrongCount: Number(row[8] || 0),
          durationSec: Number(row[9] || 0)
        };
      });
      return createJsonResponse({ ok: true, records: records });
    }

    if (action === 'practice_history') {
      const sheetName = resolveTargetSheetName(params.sheet, DEFAULT_PRACTICE_HISTORY_SHEET);
      if (sheetName === ROSTER_SHEET_NAME) {
        return createJsonResponse({ ok: false, error: '名簿シートは履歴として読み込めません。' });
      }
      const records = readRecentRecords(ss, sheetName, 12, function (row, ts, name) {
        return {
          timestamp: ts,
          name: name,
          course: row[2] == null ? '' : String(row[2]),
          category: row[3] == null ? '' : String(row[3]),
          calc: row[4] == null ? '' : String(row[4]),
          totalAnswered: Number(row[5] || 0),
          correctCount: Number(row[6] || 0),
          wrongCount: Number(row[7] || 0),
          accuracy: Number(row[8] || 0),
          durationSec: Number(row[9] || 0),
          qaText: row[10] == null ? '' : String(row[10]),
          items: parsePracticeItemsJson(row[11])
        };
      });
      return createJsonResponse({ ok: true, records: records });
    }

    const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
    if (!sheet) {
      return createJsonResponse({
        ok: false,
        error: 'シート「' + ROSTER_SHEET_NAME + '」が見つかりません。'
      });
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return createJsonResponse({ ok: true, names: [] });
    }

    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const seen = {};
    const names = [];
    for (let i = 0; i < values.length; i++) {
      const raw = values[i][0];
      const name = raw == null ? '' : String(raw).trim();
      if (!name || seen[name]) continue;
      seen[name] = true;
      names.push(name);
    }

    return createJsonResponse({ ok: true, names: names });
  } catch (err) {
    return createJsonResponse({
      ok: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

/**
 * テストモードまたは練習モードの結果を該当シートに1行追記する
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ ok: false, error: 'リクエストボディが空です。' });
    }

    const data = JSON.parse(e.postData.contents);
    const isPractice = data.mode === 'practice';
    const targetSheetName = resolveTargetSheetName(
      data.sheet,
      isPractice ? DEFAULT_PRACTICE_HISTORY_SHEET : DEFAULT_TEST_HISTORY_SHEET
    );

    if (targetSheetName === ROSTER_SHEET_NAME) {
      return createJsonResponse({ ok: false, error: '「名簿」シートへは書き込みできません。' });
    }

    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = data.timestamp && String(data.timestamp).trim()
      ? String(data.timestamp).trim()
      : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    const name = data.name != null ? String(data.name).trim() : '';
    const course = data.course != null ? String(data.course) : '';
    const category = data.category != null ? String(data.category) : '';
    const calc = data.calc != null ? String(data.calc) : '';
    const totalAnswered = Number(data.totalAnswered || 0);
    const correctCount = Number(data.correctCount || 0);
    const wrongCount = Number(data.wrongCount || 0);
    const durationSec = Number(data.durationSec || 0);

    if (isPractice) {
      const sheet = getOrCreateHistorySheet(ss, targetSheetName, PRACTICE_HISTORY_HEADERS);
      if (sheet.getLastColumn() < 12) {
        sheet.getRange(1, 11, 1, 2).setValues([['問題と解答', '詳細JSON']]);
      }
      const accuracy = data.accuracy != null
        ? Number(data.accuracy)
        : (totalAnswered > 0 ? Math.round((correctCount / totalAnswered) * 100) : 0);
      const items = Array.isArray(data.items) ? data.items : [];
      const qaText = (data.qaText && String(data.qaText).trim())
        ? String(data.qaText).trim().slice(0, MAX_CELL_CHARS)
        : buildPracticeQaText(items);
      const detailsJson = serializePracticeItems(items);
      sheet.appendRow([
        timestamp,
        name,
        course,
        category,
        calc,
        totalAnswered,
        correctCount,
        wrongCount,
        accuracy,
        durationSec,
        qaText,
        detailsJson
      ]);
      return createJsonResponse({ ok: true });
    }

    const sheet = getOrCreateHistorySheet(ss, targetSheetName, TEST_HISTORY_HEADERS);
    sheet.appendRow([
      timestamp,
      name,
      course,
      category,
      calc,
      data.result != null ? String(data.result) : '',
      totalAnswered,
      correctCount,
      wrongCount,
      durationSec
    ]);
    return createJsonResponse({ ok: true });
  } catch (err) {
    return createJsonResponse({
      ok: false,
      error: err && err.message ? err.message : String(err)
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}
