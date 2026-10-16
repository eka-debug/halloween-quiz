/**
 * Барабани долі: бекенд квізу.
 * Записує результати в Google Sheet, надсилає персонажа в Slack і віддає статистику сторінці.
 *
 * Властивості скрипта (Project Settings → Script properties):
 *   SLACK_WEBHOOK_URL  вебхук Slack (Incoming Webhook) для каналу, куди летять результати
 *   QUIZ_URL           необовʼязково: посилання на квіз уже вписане в код, властивість потрібна лише для іншої адреси
 */

var SHEET_RESULTS = 'Результати';
var SHEET_STATS = 'Статистика';
var HEADERS = ['Час', 'Імʼя', 'Персонаж', 'Ключ', 'Легендарний', 'Під маскою', 'Відповіді', 'Надіслано в Slack'];

var CHARACTERS = {
  vam: { name: 'Деймон Сальваторе', emoji: '🧛' },
  wit: { name: 'Малефісента', emoji: '🖤' },
  wer: { name: 'Бітлджус', emoji: '🪲' },
  rea: { name: 'Блейд', emoji: '🗡️' },
  jig: { name: 'Джигсо з «Пили»', emoji: '🧩' },
  mor: { name: 'Мортиша Адамс', emoji: '🥀' },
  fre: { name: 'Фредді Крюгер', emoji: '🌙' },
  lok: { name: 'Локі', emoji: '🃏' },
  // запасні герої: у квізі зʼявляться, коли їх додати на сторінці
  buf: { name: 'Баффі', emoji: '🪵' },
  con: { name: 'Константин', emoji: '🔥' },
  sal: { name: 'Кіт Салем', emoji: '🐈‍⬛' },
  chr: { name: 'Зачаровані', emoji: '🔮' },
  scw: { name: 'Багряна відьма', emoji: '🌀' },
  gom: { name: 'Гомес Адамс', emoji: '🚂' },
  ivm: { name: 'Інтелігентний вампір', emoji: '📰' },
  vla: { name: 'Владислав Тепеш', emoji: '🏰' },
  ste: { name: 'Стефан', emoji: '📓' },
  bel: { name: 'Белла', emoji: '🍎' },
  jac: { name: 'Джейкоб', emoji: '🐺' },
  chu: { name: 'Чупакабра', emoji: '🐐' },
  hag: { name: 'Хаґрід', emoji: '🐉' },
  vol: { name: 'Воландеморт', emoji: '🐍' },
  ron: { name: 'Рон Уізлі', emoji: '🕷️' },
  her: { name: 'Герміона', emoji: '📚' },
  gan: { name: 'Ґендальф', emoji: '🧙' },
  rad: { name: 'Радаґаст', emoji: '🦔' },
  ger: { name: 'Ґеральт із Рівії', emoji: '⚔️' },
  oom: { name: 'Умпа Лумпа', emoji: '🍫' },
  ora: { name: 'Джек Скеллінгтон', emoji: '🎃' }
};

/** Запусти один раз вручну: створює аркуші, заголовки і формули статистики. */
function setup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_RESULTS) || ss.insertSheet(SHEET_RESULTS);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#3B1830').setFontColor('#F3E6D1');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, HEADERS.length, 150);
  sh.setColumnWidth(3, 220);

  var st = ss.getSheetByName(SHEET_STATS) || ss.insertSheet(SHEET_STATS);
  st.clear();
  st.getRange('A1').setValue('Персонажі (перший результат кожної людини + анонімні)').setFontWeight('bold');
  st.getRange('A2').setFormula("=QUERY('" + SHEET_RESULTS + "'!A:H, \"select C, count(C) where C is not null group by C order by count(C) desc label C 'Персонаж', count(C) 'Скільки'\", 1)");
  Logger.log('Готово: аркуші створено.');
}

/** Перевірка Slack без квізу. */
function testSlack() {
  var res = postToSlack_({ name: 'Тест', key: 'rea', rare: false, base: 'rea' });
  Logger.log(res ? 'Повідомлення надіслано' : 'Не надіслано: перевір SLACK_WEBHOOK_URL');
}

/** Сторінка надсилає сюди результат. */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var key = String(data.character || '');
    var base = String(data.base || key);
    if (!CHARACTERS[key] || !CHARACTERS[base]) return json_({ ok: false, error: 'unknown character' });

    var name = clean_(data.name, 40);
    var rare = data.rare === true && key === 'ora';
    var answers = String(data.answers || '').replace(/[^0-9]/g, '').slice(0, 20);

    var sh = sheet_();
    var firstTime = !name || !alreadyPlayed_(sh, name);
    var shouldPost = !!name && (firstTime || rare);
    var posted = shouldPost ? postToSlack_({ name: name, key: key, rare: rare, base: base }) : false;

    sh.appendRow([
      new Date(), name,
      firstTime ? CHARACTERS[key].name : '',   // повторні проходження не йдуть у статистику
      key, rare ? 'так' : '', rare ? CHARACTERS[base].name : '', answers, posted ? 'так' : ''
    ]);

    return json_({ ok: true, posted: posted, firstTime: firstTime, stats: stats_(sh) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Сторінка бере звідси лічильник «Квіз уже пройшли». */
function doGet() {
  return json_({ ok: true, stats: stats_(sheet_()) });
}

/* ---------- helpers ---------- */

function sheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_RESULTS);
  if (!sh) { setup(); sh = ss.getSheetByName(SHEET_RESULTS); }
  return sh;
}

function clean_(v, max) {
  return String(v || '').replace(/[\u0000-\u001F]/g, ' ').replace(/^[=+\-@]+/, '').trim().slice(0, max);
}

function alreadyPlayed_(sh, name) {
  var last = sh.getLastRow();
  if (last < 2) return false;
  var rows = sh.getRange(2, 2, last - 1, 1).getValues();
  var n = name.toLowerCase();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === n) return true;
  }
  return false;
}

/** Рахує персонажів: перший результат кожної людини з імʼям + усі анонімні. */
function stats_(sh) {
  var counts = {}, total = 0;
  var last = sh.getLastRow();
  if (last >= 2) {
    var rows = sh.getRange(2, 3, last - 1, 2).getValues(); // C: Персонаж (порожньо для повторів), D: Ключ
    rows.forEach(function (r) {
      if (!r[0]) return;
      var k = String(r[1]);
      counts[k] = (counts[k] || 0) + 1;
      total++;
    });
  }
  return { total: total, counts: counts };
}

function slackEscape_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function postToSlack_(r) {
  // У SLACK_WEBHOOK_URL можна вписати кілька вебхуків через кому (наприклад, з різних воркспейсів).
  var urls = String(PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || '')
    .split(/[\s,;]+/).filter(function (u) { return /^https:\/\//.test(u); });
  if (!urls.length) return false;
  var quiz = PropertiesService.getScriptProperties().getProperty('QUIZ_URL') || 'https://hr-mgnt.github.io/halloween-quiz/';
  var ch = CHARACTERS[r.key];
  var who = '*' + slackEscape_(r.name) + '*';
  var text = r.rare
    ? ch.emoji + ' Три сімки на барабанах! ' + who + ' вибиває легендарний дроп: *' + ch.name + '*. Під маскою: ' + CHARACTERS[r.base].name + '.'
    : ch.emoji + ' ' + who + ' отримує персонажа: *' + ch.name + '*.';
  if (quiz) text += '\nА хто ти? <' + quiz + '|Крутити барабани долі>';
  var ok = false;
  urls.forEach(function (url) {
    try {
      var resp = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ text: text, unfurl_links: false }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) ok = true;
    } catch (e) { /* один вебхук не заважає іншим */ }
  });
  return ok;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
