const LOG_SHEET = "ログ";
const SUMMARY_SHEET = "集計";
const DAYS = "newer_than:30d";

// カードごとの「メールの探し方」と「金額の読み取り方」
const SOURCES = [
  {
    card: "楽天カード",
    query: "subject:速報版 subject:カード利用のお知らせ",
    parse: parseRakuten_,
  },
];

/* ---------- 読み取り処理 ---------- */

// HTMLメールを先に空白を1つにまとめてから読み取る
function parseRakuten_(body) {
  const text = body.replace(/[\s\u00a0\u3000]+/g, " ");
  const out = [];
  const re = /(\d{4})\/(\d{2})\/(\d{2})[^\d]{0,40}?(\d[\d,]*)\s*円/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    out.push([
      new Date(+m[1], +m[2] - 1, +m[3]),
      Number(m[4].replace(/,/g, "")),
    ]);
  }
  return out;
}

// HTMLメール本文を、セルの区切りを残したテキストに変換する
function htmlToText_(html) {
  return html
    .replace(/<(style|script)[\s\S]*?<\/\1>/gi, "")
    .replace(/<\/(td|th)>/gi, " ")
    .replace(/<(br|\/p|\/div|\/tr|\/li)[^>]*>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&yen;/g, "¥")
    .replace(/&amp;/g, "&")
    .replace(/&#x([0-9a-f]+);/gi, (_, n) =>
      String.fromCharCode(parseInt(n, 16)),
    )
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(n));
}

// まず通常のテキスト本文、読み取れなければHTML本文から読み取る
function parseMessage_(parse, msg) {
  const fromPlain = parse(msg.getPlainBody());
  return fromPlain.length ? fromPlain : parse(htmlToText_(msg.getBody()));
}

// 読み取れなかったメールの本文を、実行ログに出す
function logUnparsed_(msg) {
  const plain = msg.getPlainBody();
  const html = htmlToText_(msg.getBody());
  const at = (t) => {
    const s = t.replace(/[\s\u00a0\u3000]+/g, " ");
    const i = Math.max(s.indexOf("ご利用金額"), 0);
    return s.slice(i, i + 300);
  };
  Logger.log(
    "読み取れなかったメール: " +
      JSON.stringify({
        件名: msg.getSubject(),
        受信: String(msg.getDate()),
        テキスト長: plain.length,
        HTML長: html.length,
        テキスト抜粋: at(plain),
        HTML抜粋: at(html),
      }),
  );
}

/* ---------- 初期設定 ---------- */

// 「ログ」の列を 利用日 / 金額 / 受信日時 / メッセージID / カード にそろえる
function migrateLog_(log) {
  const head = log.getRange(1, 1, 1, 2).getValues()[0];
  if (head[0] === "利用日" && head[1] === "利用者/利用先") {
    // 旧6列版: 利用者/利用先の列だけ削除
    log.deleteColumn(2);
  } else if (head[0] === "金額") {
    // 4列版: 利用日の列を先頭に戻す
    log.insertColumnBefore(1);
    const n = log.getLastRow() - 1;
    if (n > 0) {
      const recv = log.getRange(2, 3, n, 1).getValues();
      log
        .getRange(2, 1, n, 1)
        .setValues(
          recv.map((r) => [
            r[0]
              ? new Date(r[0].getFullYear(), r[0].getMonth(), r[0].getDate())
              : "",
          ]),
        );
    }
  }
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ログシート
  const log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  migrateLog_(log);
  log
    .getRange("A1:E1")
    .setValues([["利用日", "金額", "受信日時", "メッセージID", "カード"]]);
  log.getRange("A:A").setNumberFormat("yyyy/mm/dd");
  log.getRange("B:B").setNumberFormat("#,##0");
  log.getRange("C:C").setNumberFormat("yyyy/mm/dd hh:mm");
  const n = log.getLastRow() - 1;
  if (n > 0) {
    const col = log.getRange(2, 5, n, 1);
    col.setValues(col.getValues().map((r) => [r[0] || "楽天カード"]));
  }

  // 集計シート
  const sum = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  const start = "DATE(YEAR(TODAY()),MONTH(TODAY()),1)";
  const r = `'${LOG_SHEET}'!`;
  const base = `${r}B:B,${r}A:A,">="&${start},${r}A:A,"<"&EDATE(${start},1)`;

  sum.getRange("A1").setValue("今月の利用額");
  sum
    .getRange("B1")
    .setFormula(`=SUMIFS(${base})`)
    .setNumberFormat('#,##0"円"');

  sum.getRange("D1:E10").clear();
  sum.getRange("D1").setValue("カード別(今月)");
  SOURCES.forEach((s, i) => {
    sum.getRange(2 + i, 4).setValue(s.card);
    sum
      .getRange(2 + i, 5)
      .setFormula(`=SUMIFS(${base},${r}E:E,D${2 + i})`)
      .setNumberFormat('#,##0"円"');
  });

  // トリガーを作り直す
  ScriptApp.getProjectTriggers()
    .filter((t) =>
      ["logCards", "logRakutenMails"].includes(t.getHandlerFunction()),
    )
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("logCards").timeBased().everyMinutes(1).create();

  logCards();
}

/* ---------- メイン処理(1分ごとに実行) ---------- */

function logCards() {
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);

  // 記録済みかどうかは「メッセージID」(D列)で判定する
  const lastRow = log.getLastRow();
  const done = new Set(
    lastRow > 1
      ? log
          .getRange(2, 4, lastRow - 1, 1)
          .getValues()
          .flat()
      : [],
  );
  const rows = [];

  SOURCES.forEach((src) => {
    GmailApp.search(`${src.query} ${DAYS}`).forEach((thread) => {
      thread.getMessages().forEach((msg) => {
        const id = msg.getId();
        if (done.has(id)) return; // 記録済み
        const items = parseMessage_(src.parse, msg);
        if (items.length === 0) {
          logUnparsed_(msg);
          return;
        }
        items.forEach((it) =>
          rows.push([it[0], it[1], msg.getDate(), id, src.card]),
        );
        done.add(id);
      });
    });
  });

  if (rows.length) {
    log.getRange(log.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  }
}

// 「ログ」の記録をいったん全部消して、Gmailから取り込み直す
function rebuildLog() {
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  const n = log.getLastRow() - 1;
  if (n > 0) log.getRange(2, 1, n, 5).clearContent();
  logCards();
}
