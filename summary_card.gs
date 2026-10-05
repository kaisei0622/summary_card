const LOG_SHEET = "ログ";
const SUMMARY_SHEET = "集計";
const LABEL = "card_logged";
const DAYS = "newer_than:30d";

// カードごとの「メールの探し方」と「金額の読み取り方」
const SOURCES = [
  {
    card: "三井住友(Olive)",
    query: "subject:三井住友カード subject:ご利用のお知らせ",
    parse: parseSmbc_,
  },
  {
    card: "りそなデビット",
    query:
      "subject:りそなデビットカード subject:ご利用およびご利用代金引落しのご連絡",
    parse: parseResona_,
  },
];

/* ---------- 読み取り ---------- */

function parseRakuten_(body) {
  const out = [];
  const re = /(\d{4})\/(\d{2})\/(\d{2})\s+(\S+)\s+([\d,]+)\s*円/g;
  let m;
  while ((m = re.exec(body)) !== null)
    out.push(toRow_(m[1], m[2], m[3], m[4], m[5]));
  return out;
}

function parseSmbc_(body) {
  const out = [];
  const re =
    /利用日\s*[:：]\s*(\d{4})\/(\d{2})\/(\d{2})[\s\S]*?利用先\s*[:：]\s*([^\n]+)[\s\S]*?利用金額\s*[:：]\s*([\d,]+)\s*円/g;
  let m;
  while ((m = re.exec(body)) !== null)
    out.push(toRow_(m[1], m[2], m[3], m[4], m[5]));
  return out;
}

function parseResona_(body) {
  const out = [];
  const re =
    /ご利用日時\s*[:：]\s*(\d{4})\/(\d{2})\/(\d{2})[\s\S]*?ご利用金額\s*[:：]\s*([\d,]+)\s*円[\s\S]*?ご利用加盟店名[^:：\n]*[:：]\s*([^\n]+)/g;
  let m;
  while ((m = re.exec(body)) !== null)
    out.push(toRow_(m[1], m[2], m[3], m[5], m[4]));
  return out;
}

function toRow_(y, mo, d, name, amount) {
  return [
    new Date(+y, +mo - 1, +d),
    name.trim(),
    Number(amount.replace(/,/g, "")),
  ];
}

/* ---------- 初期設定 ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ログシート(既存データは残す。F列「カード」を追加し、空欄は楽天カードで埋める)
  const log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  log
    .getRange("A1:F1")
    .setValues([
      ["利用日", "利用者/利用先", "金額", "受信日時", "メッセージID", "カード"],
    ]);
  log.getRange("A:A").setNumberFormat("yyyy/mm/dd");
  log.getRange("C:C").setNumberFormat("#,##0");
  log.getRange("D:D").setNumberFormat("yyyy/mm/dd hh:mm");
  const n = log.getLastRow() - 1;
  if (n > 0) {
    const col = log.getRange(2, 6, n, 1);
    col.setValues(col.getValues().map((r) => [r[0] || "楽天カード"]));
  }

  // 集計シート
  const sum = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  const start = "DATE(YEAR(TODAY()),MONTH(TODAY()),1)";
  const r = `'${LOG_SHEET}'!`;
  const base = `${r}C:C,${r}A:A,">="&${start},${r}A:A,"<"&EDATE(${start},1)`;

  sum.getRange("A1").setValue("今月の利用額");
  sum
    .getRange("B1")
    .setFormula(`=SUMIFS(${base})`)
    .setNumberFormat('#,##0"円"');

  sum.getRange("D1").setValue("カード別(今月)");
  SOURCES.forEach((s, i) => {
    sum.getRange(2 + i, 4).setValue(s.card);
    sum
      .getRange(2 + i, 5)
      .setFormula(`=SUMIFS(${base},${r}F:F,D${2 + i})`)
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

function logCards() {
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  const label =
    GmailApp.getUserLabelByName(LABEL) || GmailApp.createLabel(LABEL);

  const lastRow = log.getLastRow();
  const done = new Set(
    lastRow > 1
      ? log
          .getRange(2, 5, lastRow - 1, 1)
          .getValues()
          .flat()
      : [],
  );
  const rows = [];

  SOURCES.forEach((src) => {
    GmailApp.search(`${src.query} ${DAYS} -label:${LABEL}`).forEach(
      (thread) => {
        let unparsed = false;
        thread.getMessages().forEach((msg) => {
          const id = msg.getId();
          if (done.has(id)) return; // 記録済み
          const items = src.parse(msg.getPlainBody());
          if (items.length === 0) {
            unparsed = true;
            return;
          }
          items.forEach((it) =>
            rows.push([it[0], it[1], it[2], msg.getDate(), id, src.card]),
          );
          done.add(id);
        });
        // 読み取れなかったメールがあるスレッドは、次回再試行できるようラベルを付けない
        if (!unparsed) thread.addLabel(label);
      },
    );
  });

  if (rows.length) {
    log.getRange(log.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }
}
