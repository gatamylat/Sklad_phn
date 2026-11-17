import {compare, exportWorkbook, readPurchases, readRegistry, readRequest, relationTitle, showItem, showLocations, showQuantity, suggest} from './stock.js';

const $ = id => document.getElementById(id);
const state = {purchases: null, registry: null, request: null, busy: false, loading: new Set(), history: [], report: null, filter: 'all'};
const excelMime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const dateFormat = new Intl.DateTimeFormat('ru-RU', {dateStyle: 'short', timeStyle: 'short'});
const scope = new URL('./', location.href).pathname;
const store = openStore();

function node(tag, text = '', className = '') {
  const item = document.createElement(tag);
  item.textContent = text;
  if (className) item.className = className;
  return item;
}

function button(text, className, onClick) {
  const item = node('button', text, className);
  item.type = 'button';
  item.addEventListener('click', onClick);
  return item;
}

function openStore() {
  return new Promise(resolve => {
    if (!('indexedDB' in window)) return resolve(null);
    const request = indexedDB.open(`sklad-phone:${scope}`, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore('sources');
      request.result.createObjectStore('checks', {keyPath: 'id', autoIncrement: true});
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = request.onblocked = () => resolve(null);
  }).catch(() => null);
}

async function dbOperation(name, mode, act) {
  const db = await store;
  if (!db) throw new Error('Локальное хранилище недоступно.');
  return new Promise((resolve, reject) => {
    const tx = db.transaction(name, mode);
    const operation = act(tx.objectStore(name));
    tx.oncomplete = () => resolve(operation.result);
    tx.onerror = tx.onabort = () => reject(tx.error || new Error('Локальное хранилище недоступно.'));
  });
}

function status(message, error = false) {
  $('status').dataset.error = String(error);
  $('status').textContent = message;
}

function updateControls() {
  $('run').disabled = state.busy || state.loading.size > 0 || !state.request || !state.purchases || !state.registry;
  $('run').textContent = state.busy ? 'Идёт проверка…' : 'Проверить заявку';
  for (const role of ['request', 'purchases', 'registry']) $(role).disabled = state.busy || state.loading.has(role);
}

function dateText(value) {
  if (!value) return '';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : dateFormat.format(date);
}

function showFile(role) {
  const {info} = state[role];
  $(`${role}-file`).textContent = info.file;
  $(`${role}-action`).textContent = 'Заменить';
  const date = $(`${role}-date`);
  if (date) {
    date.textContent = `${info.exported ? 'Выгрузка' : 'Дата файла'}: ${dateText(info.exported || info.modified)}`;
    date.hidden = false;
  }
}

async function restore() {
  try {
    const sources = await Promise.all(['purchases', 'registry'].map(role => dbOperation('sources', 'readonly', data => data.get(role))));
    ['purchases', 'registry'].forEach((role, i) => {
      if (sources[i] && !state[role] && !state.loading.has(role)) {
        state[role] = sources[i];
        showFile(role);
      }
    });
  } catch {
    status('Локальное хранилище недоступно. Результат можно сохранить в Excel.', true);
  }
  updateControls();
}

async function pick(role, input) {
  const file = input.files?.[0];
  input.value = '';
  if (!file) return;
  state.loading.add(role);
  updateControls();
  status(`Открываю «${file.name}»…`);
  try {
    const info = {file: file.name, modified: file.lastModified ? new Date(file.lastModified).toISOString() : ''};
    let data;
    if (role === 'registry') {
      const {positions, exported} = readRegistry(await file.text());
      if (exported) info.exported = exported;
      data = positions;
    } else {
      const bytes = await file.arrayBuffer();
      if (role === 'request') {
        const {title, rows} = readRequest(bytes);
        info.sheet = title;
        data = rows;
      } else {
        data = readPurchases(bytes);
      }
    }
    state[role] = {info, data};
    showFile(role);
    if (state.report) {
      state.report = null;
      $('results').hidden = true;
    }
    if (role !== 'request') await dbOperation('sources', 'readwrite', source => source.put(state[role], role));
    status('');
  } catch (error) {
    status(`${file.name} · ${error.message}`, true);
  } finally {
    state.loading.delete(role);
    updateControls();
  }
}

async function run() {
  if ($('run').disabled) return;
  const files = {Заявка: state.request.info, Закупки: state.purchases.info, Реестр: state.registry.info};
  const report = {time: new Date().toISOString(), files};
  state.busy = true;
  updateControls();
  status('Сверяю остатки…');
  try {
    const {rows, pending} = compare(state.request.data, state.purchases.data, state.registry.data);
    const key = $('key').value.trim();
    if (key && pending.length) status(`Проверяю спорные совпадения в DeepSeek: ${pending.length}.`);
    report.ai = await suggest(pending, key);
    report.rows = rows;
  } catch (error) {
    report.error = error.message;
  }
  let saved;
  try {
    await dbOperation('checks', 'readwrite', data => data.add(report));
    saved = true;
  } catch {
    saved = false;
    state.history.unshift(report);
  }
  state.busy = false;
  updateControls();
  if (report.error) {
    status(report.error, true);
    return;
  }
  status(saved ? 'Проверка сохранена в журнале.' : 'Журнал недоступен. Сохраните результат в Excel.', !saved);
  renderReport(report);
  if (!$('check-view').hidden) $('results').focus();
}

function kind(row) {
  return {'Взять со склада': 'take', 'Частично со склада': 'partial', Докупить: 'buy'}[row.status] || 'manual';
}

function card(row) {
  const item = node('article', '', 'card');
  item.dataset.kind = kind(row);
  const main = node('div', '', 'row-main');
  const heading = node('div', '', 'row-heading');
  heading.append(node('h3', row.name || row.article || `Строка ${row.row}`), node('span', row.status, 'badge'));
  main.append(heading, node('p', [row.article && row.article !== row.name && row.article, `Строка ${row.row}`].filter(Boolean).join(' · '), 'muted'));
  const amounts = node('div', '', 'amounts');
  for (const [label, value] of [['Нужно', row.need], ['Взять', row.take], ['Докупить', row.buy]]) {
    const amount = node('div');
    amount.append(node('span', label), node('strong', `${showQuantity(value)}${value !== null && row.unit ? ` ${row.unit}` : ''}`));
    amounts.append(amount);
  }
  main.append(amounts);
  if (row.locations.length) main.append(node('p', `Полка: ${showLocations(row).join(' · ')}`, 'place'));
  const detail = node('details', '', 'row-detail');
  detail.append(node('summary', 'Основание и источники'));
  const body = node('div', '', 'detail');
  body.append(node('p', row.reason));
  for (const [key, label] of [['purchases', 'Закупки'], ['registry', 'Реестр'], ['suggestions', 'Подсказки DeepSeek'], ['candidates', 'Кандидаты']]) {
    if (!row[key]?.length) continue;
    body.append(node('strong', label));
    for (const entry of row[key]) body.append(node('p', `${entry.relation ? `${relationTitle(entry.relation)}: ` : ''}${showItem(entry)}`));
  }
  detail.append(body);
  item.append(main, detail);
  return item;
}

function renderReport(report) {
  state.report = report;
  state.filter = 'all';
  const area = $('results');
  const head = node('div', '', 'section-head');
  head.append(node('h2', 'Результат'), button('Сохранить Excel', 'secondary', () => {
    const name = `${report.files.Заявка.file.replace(/\.(xlsx|xlsm)$/iu, '')} — проверка.xlsx`;
    saveFile(exportWorkbook(report), name, excelMime).catch(error => status(error.message, true));
  }));
  const filters = node('div', '', 'range');
  filters.setAttribute('role', 'group');
  filters.setAttribute('aria-label', 'Фильтр результата');
  const list = node('div', '', 'result-list');
  const matches = (row, filter) => filter === 'all' || (filter === 'take' && row.take > 0) ||
    (filter === 'buy' && row.buy > 0) || (filter === 'manual' && kind(row) === 'manual');
  function fill() {
    const rows = report.rows.filter(row => matches(row, state.filter));
    list.replaceChildren(...(rows.length ? rows.map(card) : [node('p', 'Нет позиций по этому фильтру.', 'empty')]));
    for (const item of filters.children) item.setAttribute('aria-pressed', String(item.dataset.filter === state.filter));
  }
  for (const [filter, label] of [['all', 'Все'], ['take', 'Со склада'], ['buy', 'К закупке'], ['manual', 'Уточнить']]) {
    const item = button(`${label} · ${report.rows.filter(row => matches(row, filter)).length}`, 'filter', () => {
      state.filter = filter;
      fill();
    });
    item.dataset.filter = filter;
    filters.append(item);
  }
  const dates = Object.entries(report.files).map(([name, file]) => `${name}: ${dateText(file.exported || file.modified) || '—'}`).join(' · ');
  area.replaceChildren(head, node('p', report.files.Заявка.file, 'info'), node('p', dates, 'info'), node('p', report.ai, 'info'), filters, list);
  area.hidden = false;
  fill();
}

async function saveFile(bytes, name, mime) {
  const blob = new Blob([bytes], {type: mime});
  const file = new File([blob], name, {type: mime});
  if (navigator.canShare?.({files: [file]})) {
    try {
      await navigator.share({files: [file]});
      return;
    } catch (error) {
      if (error.name === 'AbortError') return;
    }
  }
  const link = node('a');
  const url = URL.createObjectURL(blob);
  link.href = url;
  link.download = name;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

function showView(history) {
  $('check-view').hidden = history;
  $('history-view').hidden = !history;
  $('mode').textContent = history ? 'К проверке' : 'Журнал';
  window.scrollTo(0, 0);
}

async function renderHistory() {
  const list = $('history-list');
  let reports;
  try {
    const saved = await dbOperation('checks', 'readonly', data => data.getAll());
    reports = [...state.history, ...saved.reverse()];
  } catch {
    reports = state.history;
  }
  state.visibleHistory = reports;
  $('history-export').disabled = !reports.length;
  if (!reports.length) {
    list.replaceChildren(node('p', 'В журнале пока нет проверок.', 'empty'));
    return;
  }
  list.replaceChildren(...reports.map(report => {
    const open = button('', 'history-row', () => {
      showView(false);
      if (report.error) {
        status(report.error, true);
      } else {
        status('Открыта проверка из журнала.');
        renderReport(report);
        $('results').focus();
      }
    });
    open.append(node('strong', report.files.Заявка.file), node('span', `${dateText(report.time)} · ${report.error || `${report.rows.length} поз.`}`, 'small'));
    return open;
  }));
}

for (const role of ['request', 'purchases', 'registry']) {
  $(role).addEventListener('change', event => pick(role, event.currentTarget));
}
$('run').addEventListener('click', run);
$('mode').addEventListener('click', () => {
  const open = $('history-view').hidden;
  showView(open);
  if (open) renderHistory();
});
$('history-export').addEventListener('click', () => {
  saveFile(JSON.stringify(state.visibleHistory, null, 2), 'Журнал проверок склада.json', 'application/json').catch(error => {
    $('history-list').prepend(node('p', error.message, 'empty'));
  });
});

restore();
if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(() => {});
