/**
 * Кнопка «Обновить цены» в КНИГЕ ЦЕН бренда (лист «Цены»).
 *
 * Зачем отдельный файл от Code.gs: с 10.08.2026 цены живут в своей книге
 * (`prices_book` в brands.json), а кнопка стояла только в книге полок — чтобы
 * обновить цены, человек открывал соседнюю таблицу. Здесь меню без полок:
 * в книге цен обновлять больше нечего.
 *
 * Считает не таблица: работу делает Python в GitHub Actions
 * (`brand-shelves.yml`, вход task=prices). Кнопка только шлёт workflow_dispatch.
 *
 * Шаблон: __BRAND__ и __GH_TOKEN__ подставляет deploy_button.py при заливке,
 * поэтому в репозитории лежит файл без токена.
 */

var REPO = 'DmitrySumsky/wb-shelf-positions';
var WORKFLOW = 'brand-shelves.yml';
var BRAND = '__BRAND__';
var TOKEN_KEY = 'GH_TOKEN';
// Отпечаток токена, зашитого при ПОСЛЕДНЕЙ заливке. Нужен, чтобы перезаливка с
// другим токеном чинила книгу сама: свойство GH_TOKEN переживает `clasp push`,
// и книга, куда однажды приехал негодный токен, отвечала бы 403 вечно.
var SEED_KEY = 'GH_TOKEN_SEED';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Цены WB')
    .addItem('Обновить цены', 'updatePrices')
    .addItem('Открыть журнал прогонов', 'openRuns')
    .addSeparator()
    .addItem('Заменить токен GitHub', 'setToken')
    .addToUi();
}

function token_() {
  var props = PropertiesService.getScriptProperties();
  var seeded = '__GH_TOKEN__';
  var hasSeed = seeded.indexOf('__') !== 0;   // плейсхолдер не заменён при заливке
  if (hasSeed && props.getProperty(SEED_KEY) !== seeded) {
    // Приехал НОВЫЙ зашитый токен — он и главный: 13.08.2026 в книги цен залили
    // токен без права «Actions: write», и кнопка отвечала 403.
    props.setProperty(TOKEN_KEY, seeded);
    props.setProperty(SEED_KEY, seeded);
    return seeded;
  }
  // Дальше токен живёт в свойствах скрипта: там его меняет пункт меню, и эта
  // ручная замена перезаливкой того же кода не затирается.
  return props.getProperty(TOKEN_KEY) || (hasSeed ? seeded : '');
}

function setToken() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Токен GitHub', 'Вставьте токен с правом Actions: write:',
                      ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var value = res.getResponseText().trim();
  if (!value) return;
  PropertiesService.getScriptProperties().setProperty(TOKEN_KEY, value);
  ui.alert('Токен сохранён.');
}

function updatePrices() { dispatch_('prices'); }

function dispatch_(task) {
  var ui = SpreadsheetApp.getUi();
  var tok = token_();
  if (!tok) {
    ui.alert('Нет токена GitHub. Меню «Цены WB» → «Заменить токен GitHub».');
    return;
  }
  var url = 'https://api.github.com/repos/' + REPO + '/actions/workflows/' +
            WORKFLOW + '/dispatches';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + tok,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    },
    payload: JSON.stringify({ref: 'main', inputs: {brand: BRAND, task: task}}),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code === 204) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Обновление запущено. Цены снимаются ~минуту, потом колонка за сегодня ' +
      'появится сама — перезагрузите страницу.', 'Цены: ' + BRAND, 15);
  } else {
    // Тело ответа GitHub — единственное, что объясняет отказ (истёк токен,
    // нет прав на репозиторий, переименован workflow).
    ui.alert('GitHub ответил ' + code + ':\n' + res.getContentText().slice(0, 500));
  }
}

function openRuns() {
  var url = 'https://github.com/' + REPO + '/actions/workflows/' + WORKFLOW;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      '<p><a href="' + url + '" target="_blank">Открыть журнал прогонов</a></p>')
      .setWidth(320).setHeight(80),
    'Прогоны');
}
