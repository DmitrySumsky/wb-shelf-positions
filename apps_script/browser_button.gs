/* ---------------------------------------------------------------------------
 * v2.4 (22.09.2026) — «🌐 Обновить все книги через браузер».
 *
 * С 22.09 WB не пускает облако к витрине, и полки с ценами снимает расширение
 * «Полки WB» в Chrome того, кто нажал. Одна кнопка собирает ВЕСЬ контур книг
 * (все бренды разом): следующий менеджер открывает уже обновлённую книгу.
 * Кнопка только показывает, когда был последний сбор, и открывает вкладку WB
 * с меткой сбора — дальше работает расширение.
 *
 * Дописывается к Code.gs / CodePrices.gs при заливке (deploy_button.py);
 * __HUB_URL__ и __CONTOUR__ подставляются там же.
 * ------------------------------------------------------------------------- */

var HUB_URL = '__HUB_URL__';
var CONTOUR = '__CONTOUR__';
var WB_START = 'https://www.wildberries.ru/#wbshelf=' + CONTOUR;

function hubStatus_() {
  try {
    var r = UrlFetchApp.fetch(HUB_URL + '?action=status&contour=' + CONTOUR,
                              {muteHttpExceptions: true, followRedirects: true});
    return JSON.parse(r.getContentText());
  } catch (e) {
    return {ok: false, error: String(e)};
  }
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
    return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c];
  });
}

function browserUpdate() {
  var st = hubStatus_();
  var today = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-MM-dd');
  var lines = [];
  if (!st.ok) {
    lines.push('<p style="color:#b00">Хаб сбора не ответил: ' + esc_(st.error) +
               '. Сбор всё равно можно запустить.</p>');
  } else {
    if (st.running) {
      lines.push('<p><b>Сбор уже идёт</b> — начал(а) ' + esc_(st.running.who || 'кто-то') +
                 ' в ' + esc_(String(st.running.at).slice(11, 16)) +
                 '. Второй запускать не нужно: книги обновятся сами.</p>');
    }
    if (st.last_at && String(st.last_at).slice(0, 10) === today) {
      lines.push('<p>Сегодня уже собрано в ' + esc_(String(st.last_at).slice(11, 16)) +
                 (st.last_who ? ' (' + esc_(st.last_who) + ')' : '') +
                 '. Повторный сбор обновит сегодняшнюю колонку.</p>');
    } else if (st.last_at) {
      lines.push('<p>Последний сбор: ' + esc_(st.last_at) + '. Сегодня ещё не собирали.</p>');
    } else {
      lines.push('<p>Сегодня ещё не собирали.</p>');
    }
  }
  var html =
    '<div style="font:14px/1.5 Arial,sans-serif">' +
    '<p>Сбор идёт в <b>вашем Chrome</b> расширением «Полки WB» и обновляет ' +
    '<b>все книги брендов</b> — полки и цены.</p>' + lines.join('') +
    '<ol style="padding-left:18px;margin:6px 0">' +
    '<li>Нажмите кнопку — откроется вкладка WB.</li>' +
    '<li>Если WB покажет проверку — пройдите её.</li>' +
    '<li>Не закрывайте вкладку, пока в её углу не появится «Готово» (обычно 10–30 минут).</li>' +
    '</ol>' +
    '<p><button style="font-size:14px;padding:6px 14px;cursor:pointer" ' +
    'onclick="window.open(\'' + WB_START + '\', \'_blank\'); google.script.host.close();">' +
    'Открыть WB и начать сбор</button></p>' +
    '<p style="color:#666;font-size:12px">Вкладка открылась, а панели сбора в углу нет — ' +
    'в этом Chrome не стоит расширение «Полки WB». Попросите Дмитрия.</p></div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(460).setHeight(390),
    'Обновить все книги через браузер');
}
