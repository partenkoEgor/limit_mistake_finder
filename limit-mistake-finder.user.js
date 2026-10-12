// ==UserScript==
// @name         TH Management — Limit Mistake Finder
// @namespace    th-management-limit-mistake-finder
// @version      1.0
// @description  Проходит все страницы текущего фильтра, берёт тикеты в статусе Credited (M) и сверяет Amount by receipt с лимитами субагента со страницы Confluence «ЛИМИТЫ РУЧНЫЕ». Тикеты ниже/выше лимита выводятся списком и выгружаются в Excel.
// @match        *://th-managment.com/*/admin/backoffice/paymentsupport*
// @match        *://my-managment.com/*/admin/backoffice/paymentsupport*
// @match        *://managment.io/*/admin/backoffice/paymentsupport*
// @match        *://th-managment.com/*/admin/backoffice/ExtendedPaymentRequestList*
// @match        *://my-managment.com/*/admin/backoffice/ExtendedPaymentRequestList*
// @match        *://managment.io/*/admin/backoffice/ExtendedPaymentRequestList*
// @require      https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @grant        unsafeWindow
// @connect      doc.office.lan
// @run-at       document-start
// ==/UserScript==

(function () {
  'use strict';

  // не запускаться во фреймах и не инжектиться дважды
  if (window.top !== window.self) return;
  const W = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
  if (W.__lmf_loaded) return;
  W.__lmf_loaded = true;

  const VERSION = '1.0';

  // ------------------------------------------------------------------
  // НАСТРОЙКИ
  // ------------------------------------------------------------------
  const CONFIG = {
    // Страница Confluence с лимитами (та же, что в Team Helper → limitsFinder)
    limitsPageUrl: 'https://doc.office.lan/spaces/MENA/pages/373073072/%D0%9B%D0%98%D0%9C%D0%98%D0%A2%D0%AB+%D0%A0%D0%A3%D0%A7%D0%9D%D0%AB%D0%95',
    limitsTimeout: 20000,

    // Какие статусы проверяем. Сравнение без регистра и пробелов, числовой
    // префикс вида "207 - " отбрасывается.
    statuses: ['Credited (M)'],

    // Ключи полей в JSON-ответе списка. null — определить автоматически:
    // сначала по таблице на экране (Ticket ID / Subagent / Amount by
    // receipt / Status), потом по названию ключа. Если «Диагностика»
    // показала, что поле определилось неверно, впишите ключ сюда,
    // например amount: 'amountByReceipt'.
    fields: {
      ticketId: null,
      status: null,
      amount: null,
      subagent: null,
      country: null,
      userId: null,
      currency: null,
    },

    // Заголовки колонок таблицы на экране — по ним сверяем ключи JSON.
    // Сравнение по подстроке, в нижнем регистре.
    domColumns: {
      ticketId: ['ticket id'],
      status: ['external status', 'status'],
      amount: ['amount by receipt'],
      subagent: ['subagent'],
      country: ['country'],
      userId: ['user id'],
    },

    // Ручные лимиты: перекрывают Confluence. Ключ — название субагента
    // (как в тикете), значение — { min, max } (любую границу можно не
    // указывать). Пример: 'Vodafone Cash': { min: 100, max: 30000 },
    limitOverrides: {
    },

    // Пауза между страницами списка (мс)
    pageDelay: 400,
    // Сколько раз повторять страницу при ошибке сети / 5xx
    pageRetries: 3,
    // Ожидание, пока сервер готовит выборку (is_polling)
    pollInterval: 800,
    pollMax: 8,
    // Защита от бесконечного цикла
    maxPages: 2000,
  };

  // @@core-start
  // ------------------------------------------------------------------
  // ЧИСТЫЕ ФУНКЦИИ (без DOM) — разбор сумм, лимитов, статусов, полей
  // ------------------------------------------------------------------
  const Core = (function () {
    function normSpace(s) {
      return String(s == null ? '' : s)
        .replace(/[    ]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    }

    function normStatus(s) {
      return normSpace(s).toLowerCase().replace(/^\d+\s*[-–—:]\s*/, '').replace(/\s+/g, '');
    }

    function normName(s) {
      return normSpace(s).toLowerCase()
        .replace(/ё/g, 'е')
        .replace(/[\s_\-–—.,;:/\\|()[\]"'«»]+/g, ' ')
        .trim();
    }

    // "1 234,56" / "1,234.56" / "1234.5" / 1234 → число
    function parseAmount(v) {
      if (typeof v === 'number') return isFinite(v) ? v : null;
      let s = normSpace(v);
      if (!s) return null;
      s = s.replace(/[^\d.,]/g, '');
      if (!/\d/.test(s)) return null;
      const hasComma = s.indexOf(',') !== -1;
      const hasDot = s.indexOf('.') !== -1;
      if (hasComma && hasDot) {
        // десятичный разделитель — тот, что встречается последним
        if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
        else s = s.replace(/,/g, '');
      } else if (hasComma) {
        if (/^\d{1,3}(,\d{3})+$/.test(s)) s = s.replace(/,/g, '');
        else s = s.replace(/,/g, '.');
      } else if (hasDot) {
        if (/^\d{1,3}(\.\d{3}){2,}$/.test(s)) s = s.replace(/\./g, '');
      }
      // после замен могло остаться несколько точек — берём до второй
      const m = s.match(/^\d*\.?\d*/);
      const n = parseFloat(m ? m[0] : s);
      return isFinite(n) ? n : null;
    }

    // Число в тексте лимита: "5 000", "5,000", "1 500.50", "5k", "5 тыс"
    const NUM = '(?:\\d{1,3}(?:[ \\u00a0\\u202f,.]\\d{3})+(?:[.,]\\d{1,2})?(?!\\d)|\\d+(?:[.,]\\d+)?)'
      + '(?:\\s?(?:k|к|тыс\\.?)(?![a-zа-яё]))?';

    function parseNumberToken(tok) {
      let t = normSpace(tok).toLowerCase();
      let mult = 1;
      const km = t.match(/\s?(k|к|тыс\.?)$/);
      if (km) { mult = 1000; t = t.slice(0, km.index); }
      const v = parseAmount(t);
      return v == null ? null : v * mult;
    }

    function extractNumbers(text) {
      const re = new RegExp(NUM, 'gi');
      const out = [];
      let m;
      const t = normSpace(text);
      while ((m = re.exec(t))) {
        const v = parseNumberToken(m[0]);
        if (v != null) out.push(v);
      }
      return out;
    }

    const B = '(?:^|[^a-zа-яё])'; // граница слова (\b не работает с кириллицей)
    // между словом и числом допускаем до двух слов: «минимальный депозит 300»
    const GAP = '(?:\\s+[a-zа-яё]{2,}\\.?){0,2}';
    const MIN_RE = new RegExp(B + '(?:min(?:imum)?|мин(?:имум|имальн[а-яё]*)?|from|от)\\.?' + GAP + '\\s*[:=\\-–]?\\s*(' + NUM + ')', 'i');
    const MAX_RE = new RegExp(B + '(?:max(?:imum)?|макс(?:имум|имальн[а-яё]*)?|up\\s*to|to|до)\\.?' + GAP + '\\s*[:=\\-–]?\\s*(' + NUM + ')', 'i');
    const RANGE_RE = new RegExp('(' + NUM + ')\\s*(?:[a-zа-яё$€£]{1,4}\\.?\\s*)?(?:-|–|—|~|\\.{2,3}|…|to|до)\\s*(' + NUM + ')', 'i');

    // Текст ячейки лимита → { min, max } (null — граница не найдена)
    function parseLimitText(text) {
      const t = normSpace(text);
      if (!t) return { min: null, max: null };
      const mn = t.match(MIN_RE);
      const mx = t.match(MAX_RE);
      const r = t.match(RANGE_RE);
      // «100 to 5000» / «100 до 5000» — не путать с одной границей «до 5000»
      if (r && !(mn && mx)) {
        const a = parseNumberToken(r[1]);
        const b = parseNumberToken(r[2]);
        if (a != null && b != null) return { min: Math.min(a, b), max: Math.max(a, b) };
      }
      if (mn || mx) {
        return {
          min: mn ? parseNumberToken(mn[1]) : null,
          max: mx ? parseNumberToken(mx[1]) : null,
        };
      }
      const nums = extractNumbers(t);
      if (nums.length === 2) return { min: Math.min(nums[0], nums[1]), max: Math.max(nums[0], nums[1]) };
      return { min: null, max: null };
    }

    const COL_MIN_RE = /(^| )(min|minimum|мин|минимум|минимальн\S*|от)( |$)/;
    const COL_MAX_RE = /(^| )(max|maximum|макс|максимум|максимальн\S*|до)( |$)/;
    const COL_SKIP_RE = /вывод|withdraw|payout|коммент|comment|примеч|note|контакт|contact|ссылк|link|комисс|fee/;
    const COL_LIMIT_RE = /лимит|limit|min|max|мин|макс|сумм|amount|range|диапазон|депозит|deposit/;

    // Строка таблицы Confluence → { min, max }
    function parseLimitRow(headers, cells) {
      const h = headers.map(x => normName(x));
      let min = null;
      let max = null;
      for (let i = 1; i < cells.length; i++) {
        const hi = h[i] || '';
        if (COL_SKIP_RE.test(hi)) continue;
        const isMin = COL_MIN_RE.test(hi);
        const isMax = COL_MAX_RE.test(hi);
        if (isMin === isMax) continue;
        const n = extractNumbers(cells[i])[0];
        if (n == null) continue;
        if (isMin && min == null) min = n;
        if (isMax && max == null) max = n;
      }
      if (min != null || max != null) return { min, max };

      const cols = [];
      for (let i = 1; i < cells.length; i++) {
        if (!COL_SKIP_RE.test(h[i] || '')) cols.push(i);
      }
      const limitish = cols.filter(i => COL_LIMIT_RE.test(h[i] || ''));
      for (const i of (limitish.length ? limitish : cols)) {
        const r = parseLimitText(cells[i]);
        if (r.min != null || r.max != null) return r;
      }
      return { min: null, max: null };
    }

    // tables: [{ heading, headers:[], rows:[{ cells:[text], striked }] }]
    function buildLimitIndex(tables) {
      const entries = [];
      tables.forEach(tbl => {
        tbl.rows.forEach(row => {
          const name = normSpace(row.cells[0]);
          if (!name) return;
          const norm = normName(name);
          if (!norm) return;
          const parts = String(row.cells[0]).split(/[\n,;/|]+/).map(normName).filter(Boolean);
          const lim = parseLimitRow(tbl.headers, row.cells);
          entries.push({
            name,
            norm,
            parts,
            min: lim.min,
            max: lim.max,
            heading: tbl.heading || '',
            striked: !!row.striked,
            raw: row.cells.slice(1).map(normSpace).filter(Boolean).join(' | '),
          });
        });
      });
      return entries;
    }

    function normOverrides(overrides) {
      const out = {};
      Object.keys(overrides || {}).forEach(k => {
        const v = overrides[k] || {};
        out[normName(k)] = {
          min: v.min == null ? null : Number(v.min),
          max: v.max == null ? null : Number(v.max),
        };
      });
      return out;
    }

    function hasLimit(e) { return e.min != null || e.max != null; }

    // → { status: 'ok'|'nolimit'|'unparsed'|'ambiguous'|'nosubagent', limit, entries }
    function matchLimit(entries, overrides, subagent) {
      const n = normName(subagent);
      if (!n) return { status: 'nosubagent' };
      if (overrides && overrides[n]) {
        const o = overrides[n];
        return { status: 'ok', limit: { name: subagent, min: o.min, max: o.max, heading: 'CONFIG.limitOverrides', raw: '' } };
      }
      const tiers = [
        entries.filter(e => e.norm === n),
        entries.filter(e => e.parts.indexOf(n) !== -1),
        entries.filter(e => e.norm.indexOf(n) !== -1 || (e.norm.length >= 4 && n.indexOf(e.norm) !== -1)),
      ];
      const tier = tiers.find(t => t.length);
      if (!tier) return { status: 'nolimit' };
      const active = tier.filter(e => !e.striked);
      const pool = active.length ? active : tier;
      const parsed = pool.filter(hasLimit);
      if (!parsed.length) return { status: 'unparsed', entries: pool };
      const distinct = {};
      parsed.forEach(e => { distinct[e.min + '|' + e.max] = e; });
      const keys = Object.keys(distinct);
      if (keys.length > 1) return { status: 'ambiguous', entries: parsed };
      return { status: 'ok', limit: distinct[keys[0]] };
    }

    // → { verdict: 'below'|'above'|'ok', diff }
    function evaluate(amount, limit) {
      if (limit.min != null && amount < limit.min) return { verdict: 'below', diff: amount - limit.min };
      if (limit.max != null && amount > limit.max) return { verdict: 'above', diff: amount - limit.max };
      return { verdict: 'ok', diff: 0 };
    }

    // ── Поля строки списка ────────────────────────────────────────────

    function isObj(x) { return x && typeof x === 'object' && !Array.isArray(x); }

    // Строка списка бывает объектом или массивом объектов — склеиваем
    // в один плоский объект (вложенные объекты — через точку)
    function flattenRow(item) {
      const out = {};
      const parts = Array.isArray(item) ? item : [item];
      parts.forEach(p => {
        if (!isObj(p)) return;
        Object.keys(p).forEach(k => {
          const v = p[k];
          if (isObj(v)) {
            Object.keys(v).forEach(k2 => {
              if (!isObj(v[k2]) && !Array.isArray(v[k2])) out[k + '.' + k2] = v[k2];
            });
          } else if (!Array.isArray(v) && !(k in out)) {
            out[k] = v;
          }
        });
      });
      return out;
    }

    function isText(v) { return typeof v === 'string' && /[a-zа-яё]/i.test(v); }

    const FIELD_DEFS = {
      ticketId: {
        keys: ['idTicket', 'idTkt', 'ticketId', 'id_ticket', 'ticket_id'],
        pats: [[/ticket/i, /id/i], [/tkt/i]],
        valid: v => /^\d{5,}$/.test(String(v == null ? '' : v).trim()),
      },
      status: {
        keys: ['nameExternalStatus', 'externalStatus', 'nameExtStatus', 'nameStatus', 'statusName', 'status'],
        pats: [[/ext/i, /status/i, /name/i], [/ext/i, /status/i], [/status/i, /name/i], [/status/i]],
        valid: isText,
      },
      amount: {
        keys: ['amountByReceipt', 'amount_by_receipt', 'amountReceipt', 'receiptAmount', 'sumByReceipt', 'sumReceipt'],
        pats: [[/receipt/i, /amount|sum/i], [/amount/i], [/^sum/i]],
        valid: v => v !== '' && v != null && typeof v !== 'boolean' && parseAmount(v) != null,
      },
      subagent: {
        keys: ['nameSubagent', 'subagentName', 'nameSubAgent', 'subAgentName', 'nameQuerySubagent', 'subagent', 'subAgent'],
        pats: [[/sub_?agent/i, /name/i], [/sub_?agent/i]],
        valid: isText,
      },
      country: {
        keys: ['nameCountry', 'countryName', 'country'],
        pats: [[/country/i, /name/i], [/country/i]],
        valid: isText,
      },
      userId: {
        keys: ['idUser', 'userId', 'id_user', 'user_id'],
        pats: [[/user/i, /id/i]],
        valid: v => /^\d+$/.test(String(v == null ? '' : v).trim()),
      },
      currency: {
        keys: ['nameCurrency', 'currencyCode', 'codeCurrency', 'currency'],
        pats: [[/currency/i, /name|code/i], [/currency/i]],
        valid: isText,
      },
    };

    function allKeys(rows) {
      const seen = {};
      const keys = [];
      rows.forEach(r => Object.keys(r).forEach(k => { if (!seen[k]) { seen[k] = 1; keys.push(k); } }));
      return keys;
    }

    // Ключи-кандидаты для поля в порядке приоритета (только те, где есть
    // хотя бы одно подходящее значение)
    function rankedKeys(rows, field) {
      const def = FIELD_DEFS[field];
      const keys = allKeys(rows);
      const out = [];
      const push = k => { if (out.indexOf(k) === -1 && rows.some(r => def.valid(r[k]))) out.push(k); };
      def.keys.forEach(k => { if (keys.indexOf(k) !== -1) push(k); });
      def.keys.forEach(k => keys.forEach(x => { if (x.toLowerCase() === k.toLowerCase()) push(x); }));
      def.pats.forEach(pats => keys.forEach(k => { if (pats.every(p => p.test(k))) push(k); }));
      return out;
    }

    function detectKey(rows, field) {
      return rankedKeys(rows, field)[0] || null;
    }

    function looksNumeric(s) { return /^[-+]?[\d\s.,]+(?:\s*[a-z]{3})?$/i.test(s); }

    function valuesEqual(a, b) {
      const na = normSpace(a).toLowerCase();
      const nb = normSpace(b).toLowerCase();
      if (!na || !nb) return false;
      if (na === nb) return true;
      if (looksNumeric(na) && looksNumeric(nb)) {
        const x = parseAmount(na);
        const y = parseAmount(nb);
        return x != null && y != null && Math.abs(x - y) < 0.005;
      }
      return false;
    }

    // pairs: [{ row, domValue }] — какой ключ JSON совпадает с колонкой
    // на экране. При равенстве очков — по приоритету rankedKeys.
    function pickKeyByValues(pairs, field) {
      const filled = pairs.filter(p => normSpace(p.domValue));
      if (!filled.length) return null;
      const scores = {};
      filled.forEach(p => {
        Object.keys(p.row).forEach(k => {
          if (valuesEqual(p.row[k], p.domValue)) scores[k] = (scores[k] || 0) + 1;
        });
      });
      const need = Math.max(1, Math.ceil(filled.length * 0.6));
      const cands = Object.keys(scores).filter(k => scores[k] >= need);
      if (!cands.length) return null;
      const best = Math.max.apply(null, cands.map(k => scores[k]));
      const top = cands.filter(k => scores[k] === best);
      if (top.length === 1) return top[0];
      const ranked = field ? rankedKeys(pairs.map(p => p.row), field) : [];
      return ranked.find(k => top.indexOf(k) !== -1) || top[0];
    }

    return {
      normSpace, normStatus, normName, parseAmount, parseNumberToken, extractNumbers,
      parseLimitText, parseLimitRow, buildLimitIndex, normOverrides, matchLimit, evaluate,
      flattenRow, rankedKeys, detectKey, valuesEqual, pickKeyByValues, FIELD_DEFS,
    };
  })();
  // @@core-end

  // ------------------------------------------------------------------
  // УТИЛИТЫ
  // ------------------------------------------------------------------
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const pad2 = v => ('0' + v).slice(-2);
  const log = (...a) => console.log('%c[Limit Finder ' + VERSION + ']', 'color:#b39dff;font-weight:bold', ...a);

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function fmtNum(v) {
    if (v == null || v === '') return '—';
    const n = Number(v);
    if (!isFinite(n)) return String(v);
    return n.toLocaleString('ru-RU', { maximumFractionDigits: 2 });
  }

  function getHeaders() {
    const h = { 'X-Requested-With': 'XMLHttpRequest', 'Accept': 'application/json', 'Content-Type': 'application/json' };
    const csrf = document.head && document.head.querySelector('meta[name="csrf-token"]');
    if (csrf) h['X-CSRF-TOKEN'] = csrf.getAttribute('content');
    return h;
  }

  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) return navigator.clipboard.writeText(text);
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.cssText = 'position:fixed;left:-9999px;top:0;';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); } finally { ta.remove(); }
    return Promise.resolve();
  }

  // ------------------------------------------------------------------
  // ПЕРЕХВАТ ЗАПРОСА СПИСКА (как в pay_mena Reporter)
  // ------------------------------------------------------------------
  let listUrl = null;
  let listPayload = null;
  let lastRows = null; // строки последнего ответа — то, что сейчас на экране

  function isListBody(b) {
    return b && typeof b === 'object' && typeof b.numberPage !== 'undefined' && typeof b.dateFrom !== 'undefined';
  }

  // resp.data → массив сырых строк
  function extractRawRows(resp) {
    if (!resp || !Array.isArray(resp.data) || !resp.data.length) return [];
    let raw = resp.data;
    // Как в Reporter: бывает data = [[t1, t2, …]] — разворачиваем. Если же
    // каждая строка — массив из разных объектов, оставляем как есть
    // (flattenRow склеит их).
    if (Array.isArray(raw[0]) && raw[0].length > 0) {
      const isTkt = x => !!x && !Array.isArray(x) && !!(x.idTicket || x.idTkt);
      const isRow = x => isTkt(x) || (Array.isArray(x) && isTkt(x[0]));
      if (raw[0].every(isRow)) raw = raw[0];
    }
    return raw.filter(Boolean);
  }

  function onListCaptured(url, payload, resp) {
    listUrl = url;
    listPayload = JSON.parse(JSON.stringify(payload));
    const rows = extractRawRows(resp).map(Core.flattenRow);
    if (rows.length) lastRows = rows;
    log('Запрос списка перехвачен', url);
    if (ui.ready) refreshCaptureState();
  }

  (function patchXHR() {
    const proto = W.XMLHttpRequest && W.XMLHttpRequest.prototype;
    if (!proto) return;
    const origOpen = proto.open;
    const origSend = proto.send;
    proto.open = function (m, u) {
      this.__lmf_m = m;
      this.__lmf_u = u;
      return origOpen.apply(this, arguments);
    };
    proto.send = function (body) {
      const self = this;
      if (String(this.__lmf_m).toUpperCase() === 'POST' && body) {
        try {
          const p = typeof body === 'string' ? JSON.parse(body) : body;
          if (isListBody(p)) {
            self.addEventListener('load', function () {
              try {
                const r = JSON.parse(self.responseText);
                if (r && typeof r.is_polling !== 'undefined') onListCaptured(self.responseURL || self.__lmf_u, p, r);
              } catch (e) { /* не JSON */ }
            });
          }
        } catch (e) { /* тело не JSON */ }
      }
      return origSend.apply(this, arguments);
    };
  })();

  const origFetch = W.fetch;
  (function patchFetch() {
    if (typeof origFetch !== 'function') return;
    W.fetch = function (url, opts) {
      const res = origFetch.apply(this, arguments);
      try {
        const method = (opts && opts.method) || (url && url.method) || 'GET';
        if (String(method).toUpperCase() === 'POST' && opts && opts.body) {
          const p = typeof opts.body === 'string' ? JSON.parse(opts.body) : opts.body;
          if (isListBody(p)) {
            res.then(r => r.clone().json().then(d => {
              if (d && typeof d.is_polling !== 'undefined') {
                onListCaptured(r.url || (typeof url === 'string' ? url : url.url), p, d);
              }
            })).catch(() => {});
          }
        }
      } catch (e) { /* тело не JSON */ }
      return res;
    };
  })();

  // ------------------------------------------------------------------
  // ЛИМИТЫ ИЗ CONFLUENCE (по образцу Team Helper → limitsFinder)
  // ------------------------------------------------------------------
  let limitIndex = null; // { entries, loadedAt, tables }

  function requestPage(url) {
    return new Promise((resolve, reject) => {
      if (typeof GM_xmlhttpRequest === 'function') {
        GM_xmlhttpRequest({
          method: 'GET',
          url: url,
          timeout: CONFIG.limitsTimeout,
          onload: r => (r.status >= 200 && r.status < 400 ? resolve(r.responseText) : reject(new Error('HTTP ' + r.status))),
          onerror: () => reject(new Error('network')),
          ontimeout: () => reject(new Error('timeout')),
        });
        return;
      }
      fetch(url, { credentials: 'include' }).then(r => r.text()).then(resolve, reject);
    });
  }

  function findHeadingBefore(table, doc) {
    let best = null;
    doc.querySelectorAll('h1, h2, h3, h4, h5').forEach(h => {
      if (h.compareDocumentPosition(table) & Node.DOCUMENT_POSITION_FOLLOWING) best = h;
    });
    if (!best) return '';
    const clone = best.cloneNode(true);
    clone.querySelectorAll('.copy-heading-link-container, button, .aui-icon, [aria-label]').forEach(n => n.remove());
    return clone.textContent.trim().replace(/\s+/g, ' ');
  }

  // Текст ячейки с переносами строк (как cleanHtml в Team Helper)
  function cellText(td) {
    const div = td.cloneNode(true);
    div.querySelectorAll('.tj-source, .tj-hidden').forEach(el => el.remove());
    div.querySelectorAll('p, li, div').forEach(p => p.insertAdjacentText('afterend', '\n'));
    div.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
    return div.textContent.replace(/[ \t ]+/g, ' ').replace(/\n\s*/g, '\n').trim();
  }

  function isStriked(td) {
    return !!td.querySelector('s, del, strike, [style*="line-through"]');
  }

  function readConfluenceTables(doc) {
    const tables = [];
    doc.querySelectorAll('table').forEach(table => {
      if (table.parentElement && table.parentElement.closest('table')) return; // вложенные таблицы пропускаем
      const allRows = Array.from(table.querySelectorAll('tr')).filter(tr => tr.closest('table') === table);
      if (!allRows.length) return;
      let headerRow = allRows.find(tr => tr.closest('thead'));
      if (!headerRow) {
        const first = allRows[0];
        const cells = first.querySelectorAll('th, td');
        const thCount = first.querySelectorAll('th').length;
        headerRow = thCount >= Math.ceil(cells.length / 2) ? first : null;
      }
      const headers = headerRow ? Array.from(headerRow.querySelectorAll('th, td')).map(c => c.textContent.trim()) : [];
      const rows = [];
      allRows.forEach(tr => {
        if (tr === headerRow || tr.closest('thead')) return;
        const cells = Array.from(tr.children).filter(c => /^(TD|TH)$/.test(c.tagName));
        if (cells.length < 2) return;
        rows.push({ cells: cells.map(cellText), striked: isStriked(cells[0]) });
      });
      if (rows.length) tables.push({ heading: findHeadingBefore(table, doc), headers, rows });
    });
    return tables;
  }

  async function loadLimits(force) {
    if (limitIndex && !force) return limitIndex;
    setStatus('Загружаю лимиты из Confluence…');
    const html = await requestPage(CONFIG.limitsPageUrl);
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const tables = readConfluenceTables(doc);
    if (!tables.length) {
      const looksLikeLogin = /login|log in|войти|j_username|os_username/i.test(html.slice(0, 20000));
      throw new Error(looksLikeLogin
        ? 'Confluence вернул страницу входа — откройте doc.office.lan в браузере и залогиньтесь'
        : 'на странице Confluence не найдено таблиц');
    }
    const entries = Core.buildLimitIndex(tables);
    limitIndex = { entries, tables, loadedAt: new Date() };
    const parsed = entries.filter(e => e.min != null || e.max != null).length;
    log('Лимиты загружены:', entries.length, 'строк, распознано', parsed, entries);
    return limitIndex;
  }

  // ------------------------------------------------------------------
  // ОПРЕДЕЛЕНИЕ ПОЛЕЙ JSON
  // ------------------------------------------------------------------
  function normTitle(s) {
    return String(s || '').replace(/[‘’ʼ`]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
  }

  // Таблица на экране: { headers, rows: [[text]] }
  function readScreenTable() {
    const tables = Array.from(document.querySelectorAll('table'));
    for (const table of tables) {
      if (table.closest('#lmf-wrap')) continue;
      const ths = Array.from(table.querySelectorAll('tr.table-head th, thead th'));
      if (!ths.length) continue;
      const headers = ths.map(th => normTitle(th.textContent));
      if (!headers.some(h => h.indexOf('ticket id') !== -1)) continue;
      const rows = Array.from(table.querySelectorAll('tbody tr'))
        .filter(tr => !tr.classList.contains('table-head') && tr.children.length >= Math.min(headers.length, 3))
        .map(tr => Array.from(tr.children).map(td => td.textContent.trim()));
      return { headers, rows };
    }
    return null;
  }

  function findColumn(headers, names) {
    for (const n of names) {
      const exact = headers.indexOf(n);
      if (exact !== -1) return exact;
    }
    for (const n of names) {
      const idx = headers.findIndex(h => h.indexOf(n) !== -1);
      if (idx !== -1) return idx;
    }
    return -1;
  }

  // → { keys: {field: key}, source: {field: 'config'|'screen'|'name'|null} }
  function detectFields(rows) {
    const keys = {};
    const source = {};
    const fields = Object.keys(Core.FIELD_DEFS);

    fields.forEach(f => {
      if (CONFIG.fields[f]) { keys[f] = CONFIG.fields[f]; source[f] = 'config'; }
    });
    if (!keys.ticketId) {
      keys.ticketId = Core.detectKey(rows, 'ticketId');
      source.ticketId = keys.ticketId ? 'name' : null;
    }

    // Сверка с таблицей на экране: строки последнего ответа = строки на экране
    const screen = readScreenTable();
    if (screen && lastRows && lastRows.length) {
      const tidCol = findColumn(screen.headers, CONFIG.domColumns.ticketId);
      const tidKey = keys.ticketId || Core.detectKey(lastRows, 'ticketId');
      const byId = {};
      if (tidKey) lastRows.forEach(r => { byId[String(r[tidKey]).trim()] = r; });
      const pairsFor = col => {
        const pairs = [];
        if (tidCol !== -1 && tidKey) {
          screen.rows.forEach(cells => {
            const id = ((cells[tidCol] || '').match(/\d{5,}/) || [])[0];
            if (id && byId[id]) pairs.push({ row: byId[id], domValue: cells[col] || '' });
          });
        } else if (screen.rows.length === lastRows.length) {
          screen.rows.forEach((cells, i) => pairs.push({ row: lastRows[i], domValue: cells[col] || '' }));
        }
        return pairs;
      };
      Object.keys(CONFIG.domColumns).forEach(f => {
        if (keys[f] || f === 'ticketId') return;
        const col = findColumn(screen.headers, CONFIG.domColumns[f]);
        if (col === -1) return;
        const k = Core.pickKeyByValues(pairsFor(col), f);
        if (k) { keys[f] = k; source[f] = 'screen'; }
      });
    }

    fields.forEach(f => {
      if (keys[f]) return;
      keys[f] = Core.detectKey(rows, f);
      source[f] = keys[f] ? 'name' : null;
    });
    return { keys, source };
  }

  // ------------------------------------------------------------------
  // ОБХОД СТРАНИЦ
  // ------------------------------------------------------------------
  class SessionExpired extends Error {}

  async function fetchListPage(pageNum) {
    const payload = JSON.parse(JSON.stringify(listPayload));
    payload.numberPage = pageNum;
    let lastErr = null;
    for (let attempt = 0; attempt <= CONFIG.pageRetries; attempt++) {
      if (attempt) await sleep(1500 * attempt);
      for (let poll = 0; poll <= CONFIG.pollMax; poll++) {
        let text;
        try {
          const opts = { method: 'POST', headers: getHeaders(), credentials: 'include', body: JSON.stringify(payload) };
          const r = await origFetch.call(W, listUrl, opts);
          if (r.status >= 500 || r.status === 429) { lastErr = new Error('HTTP ' + r.status); break; }
          text = await r.text();
        } catch (e) { lastErr = e; break; }
        const t = text.replace(/^\s+/, '');
        if (t.indexOf('<!DOCTYPE') === 0 || t.indexOf('<html') === 0) throw new SessionExpired('session');
        let resp;
        try { resp = JSON.parse(text); } catch (e) { lastErr = new Error('ответ не JSON'); break; }
        if (resp.is_polling && resp.count < 5 && poll < CONFIG.pollMax) {
          setStatus(`Стр. ${pageNum}: сервер готовит выборку… (${poll + 1}/${CONFIG.pollMax})`);
          await sleep(CONFIG.pollInterval);
          continue;
        }
        return extractRawRows(resp).map(Core.flattenRow);
      }
    }
    throw new Error(`страница ${pageNum}: ${lastErr ? lastErr.message : 'нет ответа'}`);
  }

  const run = {
    running: false,
    stopRequested: false,
    page: 0,
    seen: 0,
    matched: 0,
    keys: null,
    source: null,
    below: [],
    above: [],
    unchecked: [],
    startedAt: null,
    finishedAt: null,
    complete: false,
    error: '',
  };

  const UNCHECKED_REASONS = {
    nosubagent: 'не указан субагент',
    nolimit: 'субагент не найден в Confluence',
    unparsed: 'лимит в Confluence не распознан',
    ambiguous: 'несколько строк с разными лимитами',
    noamount: 'не удалось прочитать сумму',
  };

  function processRow(row, overrides) {
    const k = run.keys;
    const statusText = row[k.status];
    const wanted = CONFIG.statuses.map(Core.normStatus);
    if (wanted.indexOf(Core.normStatus(statusText)) === -1) return;
    run.matched++;

    const rec = {
      ticketId: String(row[k.ticketId] == null ? '' : row[k.ticketId]).trim(),
      subagent: Core.normSpace(row[k.subagent]),
      amountRaw: row[k.amount],
      amount: Core.parseAmount(row[k.amount]),
      currency: k.currency ? Core.normSpace(row[k.currency]) : '',
      country: k.country ? Core.normSpace(row[k.country]) : '',
      userId: k.userId ? Core.normSpace(row[k.userId]) : '',
      status: Core.normSpace(statusText),
      min: null, max: null, diff: null, heading: '', limitName: '', limitRaw: '', reason: '',
    };

    const m = Core.matchLimit(limitIndex.entries, overrides, rec.subagent);
    if (m.status !== 'ok') {
      rec.reason = UNCHECKED_REASONS[m.status] || m.status;
      if (m.entries && m.entries.length) {
        rec.reason += ': ' + m.entries.slice(0, 3).map(e => `«${e.name}» ${e.raw}`.slice(0, 160)).join(' ; ');
      }
      run.unchecked.push(rec);
      return;
    }
    Object.assign(rec, { min: m.limit.min, max: m.limit.max, heading: m.limit.heading, limitName: m.limit.name, limitRaw: m.limit.raw });
    if (rec.amount == null) {
      rec.reason = UNCHECKED_REASONS.noamount + ` («${Core.normSpace(rec.amountRaw)}»)`;
      run.unchecked.push(rec);
      return;
    }
    const ev = Core.evaluate(rec.amount, m.limit);
    rec.diff = ev.diff;
    if (ev.verdict === 'below') run.below.push(rec);
    else if (ev.verdict === 'above') run.above.push(rec);
  }

  async function startRun() {
    if (run.running) return;
    if (!listUrl || !listPayload) {
      showToast('Сначала нажмите APPLY в фильтре — скрипт перехватит запрос списка', 'err');
      return;
    }
    Object.assign(run, {
      running: true, stopRequested: false, page: 0, seen: 0, matched: 0,
      below: [], above: [], unchecked: [], startedAt: new Date(), finishedAt: null, complete: false, error: '',
      keys: null, source: null,
    });
    setBusy(true);
    renderAll();

    try {
      await loadLimits(false);
    } catch (e) {
      run.error = 'Лимиты не загружены: ' + e.message;
      finishRun();
      return;
    }

    const overrides = Core.normOverrides(CONFIG.limitOverrides);
    const seenIds = {};
    try {
      for (let page = 1; page <= CONFIG.maxPages; page++) {
        if (run.stopRequested) break;
        run.page = page;
        setStatus(`Стр. ${page}: загружаю…`);
        const rows = await fetchListPage(page);
        if (!rows.length) { run.complete = true; break; }

        if (!run.keys) {
          const det = detectFields(rows);
          run.keys = det.keys;
          run.source = det.source;
          log('Поля:', det);
          const missing = ['ticketId', 'status', 'amount', 'subagent'].filter(f => !run.keys[f]);
          if (missing.length) {
            throw new Error('не удалось определить поля: ' + missing.join(', ') + '. Откройте «Диагностику» и пропишите ключи в CONFIG.fields');
          }
        }

        let fresh = 0;
        rows.forEach(row => {
          const id = String(row[run.keys.ticketId] == null ? '' : row[run.keys.ticketId]).trim();
          if (!id || seenIds[id]) return;
          seenIds[id] = 1;
          fresh++;
          run.seen++;
          processRow(row, overrides);
        });
        renderAll();
        // сервер вернул только уже виденные тикеты — дальше страниц нет
        if (!fresh) { run.complete = true; break; }
        await sleep(CONFIG.pageDelay);
      }
    } catch (e) {
      run.error = e instanceof SessionExpired
        ? 'Сессия истекла — войдите заново и запустите проверку ещё раз'
        : 'Ошибка: ' + e.message;
      log(e);
    }
    finishRun();
  }

  function finishRun() {
    run.running = false;
    run.finishedAt = new Date();
    setBusy(false);
    renderAll();
    const total = run.below.length + run.above.length;
    if (run.error) {
      setStatus('⚠ ' + run.error + (run.seen ? ' — результаты неполные' : ''));
      showToast(run.error, 'err');
    } else if (run.stopRequested && !run.complete) {
      setStatus(`⏹ Остановлено на стр. ${run.page}. Нарушений: ${total}`);
    } else {
      setStatus(`✅ Готово: ${run.page} стр., ${run.seen} тикетов, Credited (M): ${run.matched}, нарушений: ${total}`);
      showToast(`Готово. Ниже лимита: ${run.below.length}, выше: ${run.above.length}`, 'ok');
    }
  }

  // ------------------------------------------------------------------
  // EXCEL
  // ------------------------------------------------------------------
  function downloadExcel() {
    if (typeof XLSX === 'undefined') { showToast('Библиотека XLSX не загрузилась (@require)', 'err'); return; }
    const total = run.below.length + run.above.length + run.unchecked.length;
    if (!total) { showToast('Нет данных для выгрузки', 'err'); return; }

    const border = { style: 'thin', color: { rgb: 'BFBFBF' } };
    const borders = { top: border, bottom: border, left: border, right: border };
    const headStyle = (rgb) => ({
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { patternType: 'solid', fgColor: { rgb } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: borders,
    });

    function sheet(header, rows, color, numCols) {
      const ws = XLSX.utils.aoa_to_sheet([header].concat(rows));
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let c = range.s.c; c <= range.e.c; c++) {
        for (let r = range.s.r; r <= range.e.r; r++) {
          const ref = XLSX.utils.encode_cell({ r, c });
          if (!ws[ref]) continue;
          if (r === 0) { ws[ref].s = headStyle(color); continue; }
          ws[ref].s = { border: borders, alignment: { vertical: 'top', wrapText: c === range.e.c } };
          if (numCols.indexOf(c) !== -1 && ws[ref].t === 'n') ws[ref].z = '#,##0.##';
        }
      }
      ws['!cols'] = header.map((h, c) => {
        let w = String(h).length;
        rows.forEach(row => { w = Math.max(w, String(row[c] == null ? '' : row[c]).length); });
        return { wch: Math.min(Math.max(w + 2, 8), 60) };
      });
      ws['!rows'] = [{ hpt: 24 }];
      if (rows.length) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: header.length - 1 } }) };
      return ws;
    }

    const num = v => (v == null || v === '' ? '' : Number(v));
    const vHeader = ['№', 'Ticket ID', 'Subagent', 'Amount by receipt', 'Валюта', 'Min', 'Max', 'Отклонение', 'Country', 'User ID', 'Status', 'Строка в Confluence', 'Раздел Confluence', 'Лимит (как в Confluence)'];
    const vRow = (r, i) => [i + 1, r.ticketId, r.subagent, num(r.amount), r.currency, num(r.min), num(r.max), num(r.diff), r.country, r.userId, r.status, r.limitName, r.heading, r.limitRaw];
    const uHeader = ['№', 'Ticket ID', 'Subagent', 'Amount by receipt', 'Валюта', 'Country', 'User ID', 'Status', 'Причина'];
    const uRow = (r, i) => [i + 1, r.ticketId, r.subagent, r.amount == null ? Core.normSpace(r.amountRaw) : r.amount, r.currency, r.country, r.userId, r.status, r.reason];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet(vHeader, run.below.map(vRow), '2F5597', [3, 5, 6, 7]), 'Ниже лимита');
    XLSX.utils.book_append_sheet(wb, sheet(vHeader, run.above.map(vRow), 'C55A11', [3, 5, 6, 7]), 'Выше лимита');
    XLSX.utils.book_append_sheet(wb, sheet(uHeader, run.unchecked.map(uRow), '7F7F7F', [3]), 'Не удалось проверить');

    // Сводка по субагентам
    const bySub = {};
    const add = (r, key) => {
      const s = bySub[r.subagent] || (bySub[r.subagent] = { below: 0, above: 0, min: r.min, max: r.max });
      s[key]++;
    };
    run.below.forEach(r => add(r, 'below'));
    run.above.forEach(r => add(r, 'above'));
    const sumRows = Object.keys(bySub)
      .sort((a, b) => (bySub[b].below + bySub[b].above) - (bySub[a].below + bySub[a].above))
      .map(s => [s, bySub[s].below, bySub[s].above, bySub[s].below + bySub[s].above, num(bySub[s].min), num(bySub[s].max)]);
    const info = [
      ['Сформировано', fmtDateTime(new Date())],
      ['Период фильтра', `${listPayload && listPayload.dateFrom || ''} — ${listPayload && listPayload.dateTo || ''}`],
      ['Статусы', CONFIG.statuses.join(', ')],
      ['Страниц просмотрено', run.page],
      ['Тикетов просмотрено', run.seen],
      ['Тикетов в нужном статусе', run.matched],
      ['Ниже лимита', run.below.length],
      ['Выше лимита', run.above.length],
      ['Не удалось проверить', run.unchecked.length],
      ['Обход завершён', run.complete ? 'да' : 'нет — ' + (run.error || 'остановлен вручную')],
      [],
    ];
    const ws = sheet(['Subagent', 'Ниже', 'Выше', 'Всего', 'Min', 'Max'], sumRows, '375623', [4, 5]);
    XLSX.utils.sheet_add_aoa(ws, info, { origin: { r: sumRows.length + 2, c: 0 } });
    ws['!cols'][0] = { wch: Math.max(ws['!cols'][0].wch, 26) };
    XLSX.utils.book_append_sheet(wb, ws, 'Сводка');

    const d = new Date();
    const fn = `limit_mistakes_${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}_${pad2(d.getHours())}-${pad2(d.getMinutes())}.xlsx`;
    XLSX.writeFile(wb, fn);
    showToast('Скачан ' + fn, 'ok');
  }

  function fmtDateTime(d) {
    return `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
  }

  // ------------------------------------------------------------------
  // ДИАГНОСТИКА
  // ------------------------------------------------------------------
  async function runDiagnostics() {
    ui.tab = 'diag';
    renderAll();
    const out = [];
    out.push(`Версия: ${VERSION}`);
    out.push(`Запрос списка: ${listUrl ? listUrl : 'НЕ ПЕРЕХВАЧЕН — нажмите APPLY в фильтре'}`);
    if (listPayload) out.push(`Тело запроса: ${JSON.stringify(listPayload)}`);
    out.push('');

    let rows = lastRows;
    if (!rows && listUrl) {
      try { rows = await fetchListPage(1); } catch (e) { out.push('Не удалось загрузить стр. 1: ' + e.message); }
    }
    if (rows && rows.length) {
      const det = detectFields(rows);
      const src = { config: 'CONFIG.fields', screen: 'по таблице на экране', name: 'по названию ключа' };
      out.push('Определённые поля:');
      Object.keys(Core.FIELD_DEFS).forEach(f => {
        const k = det.keys[f];
        out.push(`  ${f.padEnd(9)} → ${k ? k : '— не найдено'}${k ? `   (${src[det.source[f]] || ''}; пример: ${JSON.stringify(rows[0][k])})` : ''}`);
      });
      out.push('');
      out.push('Все ключи первой строки:');
      Object.keys(rows[0]).forEach(k => out.push(`  ${k}: ${JSON.stringify(rows[0][k])}`.slice(0, 200)));
      log('Первая строка списка:', rows[0], 'поля:', det);
    } else if (listUrl) {
      out.push('Список пуст.');
    }
    out.push('');

    try {
      const li = await loadLimits(false);
      const parsed = li.entries.filter(e => e.min != null || e.max != null);
      out.push(`Confluence: таблиц ${li.tables.length}, строк ${li.entries.length}, лимит распознан в ${parsed.length}`);
      li.tables.slice(0, 5).forEach((t, i) => out.push(`  Таблица ${i + 1} «${t.heading}»: колонки ${JSON.stringify(t.headers)}`));
      out.push('Примеры распознанных строк:');
      parsed.slice(0, 8).forEach(e => out.push(`  ${e.name} → min ${fmtNum(e.min)}, max ${fmtNum(e.max)}${e.striked ? ' (зачёркнута)' : ''}`));
      const bad = li.entries.filter(e => e.min == null && e.max == null);
      if (bad.length) {
        out.push(`Нераспознанные строки (${bad.length}), первые 8:`);
        bad.slice(0, 8).forEach(e => out.push(`  ${e.name} → ${e.raw.slice(0, 150)}`));
      }
    } catch (e) {
      out.push('Confluence: ошибка — ' + e.message);
    }
    ui.diagText = out.join('\n');
    setStatus('Диагностика готова');
    renderAll();
  }

  // ------------------------------------------------------------------
  // ИНТЕРФЕЙС
  // ------------------------------------------------------------------
  const ui = { ready: false, tab: 'below', diagText: '' };

  const CSS = `
    #lmf-fab{position:fixed;bottom:24px;right:160px;z-index:2147483640;display:flex;align-items:center;gap:7px;padding:10px 16px;border-radius:26px;background:rgba(18,14,35,.9);backdrop-filter:blur(16px);border:1px solid rgba(155,130,255,.5);color:#cbbcff;font:700 13px Arial,sans-serif;cursor:pointer;box-shadow:0 6px 24px rgba(0,0,0,.5);user-select:none;}
    #lmf-fab:hover{box-shadow:0 8px 32px rgba(0,0,0,.6),0 0 16px rgba(155,130,255,.3);}
    #lmf-fab .dot{width:7px;height:7px;border-radius:50%;background:#9b82ff;}
    #lmf-fab.busy .dot{animation:lmfpulse 1.2s ease-in-out infinite;box-shadow:0 0 6px #9b82ff;}
    #lmf-fab .cnt{background:#e5484d;color:#fff;border-radius:10px;padding:0 6px;font-size:11px;display:none;}
    @keyframes lmfpulse{0%,100%{opacity:1}50%{opacity:.3}}
    #lmf-wrap{position:fixed;inset:0;z-index:2147483641;background:rgba(4,6,16,.55);backdrop-filter:blur(5px);display:none;align-items:center;justify-content:center;}
    #lmf-wrap.open{display:flex;}
    #lmf-panel{width:94vw;max-width:1250px;height:88vh;background:rgba(16,14,30,.94);border:1px solid rgba(155,130,255,.4);border-radius:16px;display:flex;flex-direction:column;overflow:hidden;font:13px Arial,sans-serif;color:#e6e2ff;box-shadow:0 24px 64px rgba(0,0,0,.7);}
    #lmf-hdr{display:flex;align-items:center;gap:10px;padding:12px 18px;border-bottom:1px solid rgba(155,130,255,.25);}
    #lmf-hdr .t{font-weight:800;font-size:14px;color:#cbbcff;}
    #lmf-hdr .v{font-size:10px;color:#8a7fc0;border:1px solid rgba(155,130,255,.35);border-radius:4px;padding:1px 6px;}
    #lmf-close{margin-left:auto;background:none;border:none;color:rgba(255,255,255,.4);font-size:18px;cursor:pointer;}
    #lmf-close:hover{color:#cbbcff;}
    #lmf-ctrl{display:flex;flex-wrap:wrap;gap:8px;padding:12px 18px;border-bottom:1px solid rgba(255,255,255,.06);align-items:center;}
    .lmf-btn{padding:8px 14px;border-radius:8px;border:1px solid rgba(155,130,255,.45);background:rgba(155,130,255,.14);color:#d9d0ff;font:700 12px Arial,sans-serif;cursor:pointer;}
    .lmf-btn:hover:not(:disabled){background:rgba(155,130,255,.28);}
    .lmf-btn:disabled{opacity:.35;cursor:not-allowed;}
    .lmf-btn.primary{background:rgba(155,130,255,.32);border-color:#9b82ff;}
    .lmf-btn.stop{border-color:rgba(250,80,80,.4);background:rgba(250,80,80,.1);color:#ff8a8a;}
    .lmf-btn.xl{border-color:rgba(63,200,120,.45);background:rgba(63,200,120,.12);color:#7ee0a5;}
    #lmf-capture{font-size:11px;padding:6px 10px;border-radius:7px;margin-left:auto;}
    #lmf-capture.ok{color:#7ee0a5;background:rgba(63,200,120,.08);border:1px solid rgba(63,200,120,.3);}
    #lmf-capture.no{color:#f0a500;background:rgba(240,165,0,.08);border:1px solid rgba(240,165,0,.3);}
    #lmf-metrics{display:grid;grid-template-columns:repeat(6,1fr);gap:8px;padding:12px 18px;}
    .lmf-m{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);border-radius:9px;padding:8px;text-align:center;}
    .lmf-m b{display:block;font-size:20px;color:#fff;font-variant-numeric:tabular-nums;}
    .lmf-m span{font-size:9px;color:#9a93c8;text-transform:uppercase;letter-spacing:.6px;}
    .lmf-m.below b{color:#7cb4ff;} .lmf-m.above b{color:#ffb070;} .lmf-m.unk b{color:#bbb;}
    #lmf-tabs{display:flex;gap:4px;padding:0 18px;border-bottom:1px solid rgba(255,255,255,.08);}
    .lmf-tab{padding:8px 14px;border:none;border-bottom:2px solid transparent;background:none;color:#9a93c8;font:700 12px Arial,sans-serif;cursor:pointer;}
    .lmf-tab.on{color:#e6e2ff;border-bottom-color:#9b82ff;}
    #lmf-body{flex:1;overflow:auto;padding:0 18px 12px;}
    #lmf-body table{width:100%;border-collapse:collapse;font-size:12px;}
    #lmf-body th{position:sticky;top:0;background:#1d1936;color:#b9adf5;text-align:left;padding:8px 6px;font-size:11px;border-bottom:1px solid rgba(155,130,255,.3);z-index:1;}
    #lmf-body td{padding:6px;border-bottom:1px solid rgba(255,255,255,.05);vertical-align:top;}
    #lmf-body tr:hover td{background:rgba(155,130,255,.06);}
    #lmf-body td.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap;}
    #lmf-body td.neg{color:#7cb4ff;} #lmf-body td.pos{color:#ffb070;}
    #lmf-body .tid{color:#cbbcff;font-weight:700;cursor:pointer;text-decoration:underline dotted;}
    #lmf-body .muted{color:#8a84b0;font-size:11px;}
    #lmf-body .empty{padding:30px;text-align:center;color:#8a84b0;}
    #lmf-body pre{white-space:pre-wrap;font:12px/1.5 Consolas,monospace;color:#d8d2ff;margin:12px 0;}
    #lmf-foot{padding:8px 18px;border-top:1px solid rgba(255,255,255,.06);font-size:11px;color:#a8a1d6;min-height:16px;}
    #lmf-toast{position:fixed;bottom:130px;right:24px;z-index:2147483647;padding:12px 18px;border-radius:12px;color:#eee;font:600 12px Arial,sans-serif;max-width:360px;box-shadow:0 8px 24px rgba(0,0,0,.5);pointer-events:none;}
  `;

  function showToast(msg, type) {
    if (!document.body) return;
    const old = document.getElementById('lmf-toast');
    if (old) old.remove();
    const t = document.createElement('div');
    t.id = 'lmf-toast';
    t.style.background = type === 'ok' ? 'rgba(0,90,40,.92)' : type === 'err' ? 'rgba(110,20,20,.92)' : 'rgba(30,25,70,.92)';
    t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(() => t.remove(), 4500);
  }

  function setStatus(msg) {
    const el = document.getElementById('lmf-foot');
    if (el) el.textContent = msg;
  }

  function setBusy(on) {
    const $ = id => document.getElementById(id);
    if (!$('lmf-start')) return;
    $('lmf-start').disabled = on;
    $('lmf-stop').style.display = on ? '' : 'none';
    $('lmf-reload').disabled = on;
    $('lmf-diag').disabled = on;
    $('lmf-fab').classList.toggle('busy', on);
  }

  function refreshCaptureState() {
    const el = document.getElementById('lmf-capture');
    if (!el) return;
    if (listUrl) {
      el.className = 'ok';
      const p = listPayload || {};
      el.textContent = '✅ Фильтр пойман' + (p.dateFrom ? `: ${p.dateFrom} — ${p.dateTo || ''}` : '');
    } else {
      el.className = 'no';
      el.textContent = '⚠ Нажмите APPLY в фильтре — скрипт перехватит запрос';
    }
  }

  function vTable(list, kind) {
    if (!list.length) {
      return `<div class="empty">${run.running ? 'Идёт проверка…' : run.startedAt ? 'Нет тикетов' : 'Нажмите «Проверить все страницы»'}</div>`;
    }
    const rows = list.map((r, i) => `
      <tr>
        <td class="muted">${i + 1}</td>
        <td><span class="tid" data-tid="${esc(r.ticketId)}" title="Скопировать">${esc(r.ticketId)}</span></td>
        <td>${esc(r.subagent)}${r.limitName && Core.normName(r.limitName) !== Core.normName(r.subagent) ? `<div class="muted">в Confluence: ${esc(r.limitName)}</div>` : ''}</td>
        <td class="num">${fmtNum(r.amount)} ${esc(r.currency)}</td>
        <td class="num">${fmtNum(r.min)}</td>
        <td class="num">${fmtNum(r.max)}</td>
        <td class="num ${kind === 'below' ? 'neg' : 'pos'}">${r.diff > 0 ? '+' : ''}${fmtNum(r.diff)}</td>
        <td>${esc(r.country)}</td>
        <td class="muted">${esc(r.heading)}</td>
      </tr>`).join('');
    return `<table><thead><tr><th>№</th><th>Ticket ID</th><th>Subagent</th><th>Amount by receipt</th><th>Min</th><th>Max</th><th>Отклонение</th><th>Country</th><th>Раздел Confluence</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  function uTable(list) {
    if (!list.length) return `<div class="empty">${run.startedAt ? 'Все тикеты проверены' : 'Пока пусто'}</div>`;
    const rows = list.map((r, i) => `
      <tr>
        <td class="muted">${i + 1}</td>
        <td><span class="tid" data-tid="${esc(r.ticketId)}" title="Скопировать">${esc(r.ticketId)}</span></td>
        <td>${esc(r.subagent)}</td>
        <td class="num">${r.amount == null ? esc(Core.normSpace(r.amountRaw)) : fmtNum(r.amount)} ${esc(r.currency)}</td>
        <td>${esc(r.country)}</td>
        <td>${esc(r.reason)}</td>
      </tr>`).join('');
    return `<table><thead><tr><th>№</th><th>Ticket ID</th><th>Subagent</th><th>Amount by receipt</th><th>Country</th><th>Причина</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  let renderPending = false;
  function renderAll() {
    if (!ui.ready || renderPending) return;
    renderPending = true;
    requestAnimationFrame(() => {
      renderPending = false;
      const $ = id => document.getElementById(id);
      $('m-pages').textContent = run.page || '—';
      $('m-seen').textContent = run.seen || '—';
      $('m-matched').textContent = run.matched || '—';
      $('m-below').textContent = run.below.length;
      $('m-above').textContent = run.above.length;
      $('m-unk').textContent = run.unchecked.length;
      $('t-below').textContent = `Ниже лимита (${run.below.length})`;
      $('t-above').textContent = `Выше лимита (${run.above.length})`;
      $('t-unk').textContent = `Не удалось проверить (${run.unchecked.length})`;
      document.querySelectorAll('.lmf-tab').forEach(b => b.classList.toggle('on', b.dataset.tab === ui.tab));
      const total = run.below.length + run.above.length;
      const cnt = $('lmf-fab').querySelector('.cnt');
      cnt.textContent = total;
      cnt.style.display = total ? '' : 'none';
      $('lmf-xlsx').disabled = !(total + run.unchecked.length);

      const body = $('lmf-body');
      if (ui.tab === 'below') body.innerHTML = vTable(run.below, 'below');
      else if (ui.tab === 'above') body.innerHTML = vTable(run.above, 'above');
      else if (ui.tab === 'unk') body.innerHTML = uTable(run.unchecked);
      else body.innerHTML = `<pre>${esc(ui.diagText || 'Нажмите «Диагностика»')}</pre>`;
    });
  }

  function initUI() {
    GM_addStyle(CSS);

    const fab = document.createElement('div');
    fab.id = 'lmf-fab';
    fab.title = 'Alt+L — открыть/закрыть';
    fab.innerHTML = '<span class="dot"></span><span>Limits</span><span class="cnt"></span>';
    document.body.appendChild(fab);

    const wrap = document.createElement('div');
    wrap.id = 'lmf-wrap';
    wrap.innerHTML = `
      <div id="lmf-panel">
        <div id="lmf-hdr">
          <span class="t">🔎 Limit Mistake Finder</span><span class="v">${VERSION}</span>
          <span class="v">${esc(CONFIG.statuses.join(', '))}</span>
          <button id="lmf-close" title="Закрыть (Esc)">✕</button>
        </div>
        <div id="lmf-ctrl">
          <button class="lmf-btn primary" id="lmf-start">▶ Проверить все страницы</button>
          <button class="lmf-btn stop" id="lmf-stop" style="display:none">⏹ Стоп</button>
          <button class="lmf-btn xl" id="lmf-xlsx" disabled>⬇ Скачать Excel</button>
          <button class="lmf-btn" id="lmf-reload" title="Заново загрузить страницу лимитов из Confluence">🔄 Обновить лимиты</button>
          <button class="lmf-btn" id="lmf-diag">🩺 Диагностика</button>
          <span id="lmf-capture"></span>
        </div>
        <div id="lmf-metrics">
          <div class="lmf-m"><b id="m-pages">—</b><span>Страниц</span></div>
          <div class="lmf-m"><b id="m-seen">—</b><span>Тикетов</span></div>
          <div class="lmf-m"><b id="m-matched">—</b><span>${esc(CONFIG.statuses[0] || '')}</span></div>
          <div class="lmf-m below"><b id="m-below">0</b><span>Ниже лимита</span></div>
          <div class="lmf-m above"><b id="m-above">0</b><span>Выше лимита</span></div>
          <div class="lmf-m unk"><b id="m-unk">0</b><span>Не проверено</span></div>
        </div>
        <div id="lmf-tabs">
          <button class="lmf-tab on" data-tab="below" id="t-below">Ниже лимита</button>
          <button class="lmf-tab" data-tab="above" id="t-above">Выше лимита</button>
          <button class="lmf-tab" data-tab="unk" id="t-unk">Не удалось проверить</button>
          <button class="lmf-tab" data-tab="diag">Диагностика</button>
        </div>
        <div id="lmf-body"></div>
        <div id="lmf-foot">Готов к работе</div>
      </div>`;
    document.body.appendChild(wrap);

    const toggle = open => wrap.classList.toggle('open', open == null ? !wrap.classList.contains('open') : open);
    fab.addEventListener('click', () => toggle());
    document.getElementById('lmf-close').addEventListener('click', () => toggle(false));
    wrap.addEventListener('mousedown', e => { if (e.target === wrap) toggle(false); });
    document.addEventListener('keydown', e => {
      if (e.altKey && (e.code === 'KeyL')) { e.preventDefault(); toggle(); }
      if (e.key === 'Escape' && wrap.classList.contains('open')) toggle(false);
    });

    document.getElementById('lmf-start').addEventListener('click', () => {
      ui.tab = ui.tab === 'diag' ? 'below' : ui.tab;
      startRun();
    });
    document.getElementById('lmf-stop').addEventListener('click', () => {
      run.stopRequested = true;
      setStatus('Останавливаю после текущей страницы…');
    });
    document.getElementById('lmf-xlsx').addEventListener('click', downloadExcel);
    document.getElementById('lmf-reload').addEventListener('click', async () => {
      try {
        const li = await loadLimits(true);
        const parsed = li.entries.filter(e => e.min != null || e.max != null).length;
        setStatus(`Лимиты обновлены: ${li.entries.length} строк, распознано ${parsed}`);
        showToast('Лимиты обновлены', 'ok');
      } catch (e) {
        setStatus('⚠ Лимиты не загружены: ' + e.message);
        showToast('Лимиты не загружены: ' + e.message, 'err');
      }
    });
    document.getElementById('lmf-diag').addEventListener('click', runDiagnostics);
    document.getElementById('lmf-tabs').addEventListener('click', e => {
      const b = e.target.closest('.lmf-tab');
      if (!b) return;
      ui.tab = b.dataset.tab;
      renderAll();
    });
    document.getElementById('lmf-body').addEventListener('click', e => {
      const t = e.target.closest('.tid');
      if (!t) return;
      copyText(t.dataset.tid).then(() => showToast('Скопировано: ' + t.dataset.tid, 'ok'));
    });

    ui.ready = true;
    refreshCaptureState();
    renderAll();
  }

  if (document.body) initUI();
  else document.addEventListener('DOMContentLoaded', initUI, { once: true });
})();
