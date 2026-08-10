/**
 * Кнопка «Обновить полки» в книге бренда.
 *
 * Считает не эта таблица: работу делает Python в GitHub Actions (обход полок WB
 * занимает минуты и не влезает в лимиты Apps Script). Кнопка только дёргает
 * workflow_dispatch и показывает ссылку на прогон.
 *
 * Шаблон: __BRAND__ и __GH_TOKEN__ подставляет deploy_button.py при заливке,
 * поэтому в репозитории лежит файл без токена.
 */

var REPO = 'DmitrySumsky/wb-shelf-positions';
var WORKFLOW = 'brand-shelves.yml';
var BRAND = '__BRAND__';
var TOKEN_KEY = 'GH_TOKEN';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Полки WB')
    .addItem('Обновить полки', 'updateShelves')
    .addItem('Обновить цены', 'updatePrices')
    .addItem('Обновить всё', 'updateAll')
    .addItem('Открыть журнал прогонов', 'openRuns')
    .addSeparator()
    .addItem('Заменить токен GitHub', 'setToken')
    .addToUi();
}

function token_() {
  var props = PropertiesService.getScriptProperties();
  var saved = props.getProperty(TOKEN_KEY);
  if (saved) return saved;
  // Первый запуск: переносим зашитый при заливке токен в свойства скрипта,
  // дальше он живёт только там и обновляется пунктом меню.
  var seeded = '__GH_TOKEN__';
  if (seeded.indexOf('__') === 0) return '';
  props.setProperty(TOKEN_KEY, seeded);
  return seeded;
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

function updateShelves() { dispatch_('shelves'); }
function updatePrices()  { dispatch_('prices'); }
function updateAll()     { dispatch_('both'); }

/**
 * Запуск прогона. task: shelves — только лист «Полки», prices — только лист
 * «Цены», both — оба листа одним прогоном (цены идут после полок).
 */
function dispatch_(task) {
  var ui = SpreadsheetApp.getUi();
  var tok = token_();
  if (!tok) {
    ui.alert('Нет токена GitHub. Меню «Полки WB» → «Заменить токен GitHub».');
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
    var what = task === 'prices' ? 'Цены' : (task === 'both' ? 'Полки и цены' : 'Полки');
    var eta = task === 'prices' ? 'Цены снимаются ~минуту' : 'Полки обходятся ~5 минут';
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Обновление запущено. ' + eta + ', потом колонка за сегодня ' +
      'появится сама — перезагрузите страницу.', what + ': ' + BRAND, 15);
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
