// ==UserScript==
// @name         TH Management — Limit Mistake Finder
// @namespace    th-management-limit-mistake-finder
// @version      1.12
// @description  Проходит все страницы текущего фильтра, берёт тикеты в статусе Credited (M) и сверяет Amount by receipt с лимитами метода (с учётом страны) со страницы Confluence «ЛИМИТЫ РУЧНЫЕ». Тикеты ниже/выше лимита выводятся списком и выгружаются в Excel.
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

  const VERSION = '1.12';

  // ------------------------------------------------------------------
  // НАСТРОЙКИ
  // ------------------------------------------------------------------
  const CONFIG = {
    // Страница Confluence с лимитами (та же, что в Team Helper → limitsFinder)
    limitsPageUrl: 'https://doc.office.lan/spaces/MENA/pages/373073072/%D0%9B%D0%98%D0%9C%D0%98%D0%A2%D0%AB+%D0%A0%D0%A3%D0%A7%D0%9D%D0%AB%D0%95',
    limitsTimeout: 20000,

    // Какие статусы проверяем. Сравнение без регистра и пробелов, числовой
    // префикс вида "207 - " отбрасывается.
    statuses: ['Credited (M)', 'Credited (Fraud) (M)'],

    // Ключи полей в JSON-ответе списка. null — определить автоматически:
    // сначала по таблице на экране (Ticket ID / Subagent / Amount by
    // receipt / Status), потом по названию ключа. Если «Диагностика»
    // показала, что поле определилось неверно, впишите ключ сюда,
    // например amount: 'amountByReceipt'.
    // amount — это Amount by receipt (подтверждено); summa — сумма транзакции.
    fields: {
      ticketId: null,
      status: null,
      amount: 'amount',
      method: null,
      subagent: null,
      ref: null,
      refId: null,
      subagentId: null,
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
      method: ['method'],
      subagent: ['subagent'],
      country: ['country'],
      userId: ['user id'],
    },

    // Ручные лимиты: перекрывают Confluence. Ключ — method или subagent
    // (как в тикете), значение — { min, max } (любую границу можно не
    // указывать). Пример: 'Sham Cash SYP': { min: 100, max: 30000 },
    limitOverrides: {
    },

    // Бренды/проекты в Confluence (в скобках после метода или в
    // «Переключалке»). Дополняется автоматически тем, что встречается в
    // скобках 2+ раз. По nameRef тикета выбирается строка своего проекта.
    knownBrands: ['1xcasino', 'afropari', 'bizbet', 'rolsbet', 'melbet', 'betwinner', 'megapari', 'paripesa',
      '22bet', 'linebet', 'helabet', '888starz', 'betandyou', '1xbit', 'dbbet', 'xparibet', 'betpari'],

    // Проекты (nameRef), которые считаются webDefault: строка — точное
    // совпадение после нормализации, RegExp — по шаблону. «1xbet22.com» →
    // «1xbet22com» подходит под /^1xbet/
    webDefaultRefs: ['webdefault', /1xbet/],

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
        .replace(/[\u200b-\u200f\u2060\ufeff\u00ad]/g, '') // невидимые символы из Confluence
        .replace(/[\u00a0\u202f\u2007\u2009]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    }

    function normStatus(s) {
      return normSpace(s).toLowerCase().replace(/^\d+\s*[-–—:]?\s*/, '').replace(/\s+/g, '');
    }

    // всё, кроме букв и цифр, — в пробел: «L'bankalik» → «l bankalik»
    // диакритика снимается: «Hızlı» = «Hizli», «YapıKredi» = «YapiKredi»
    function normName(s) {
      return normSpace(s).toLowerCase()
        .replace(/ё/g, 'е')
        .replace(/ı/g, 'i')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '').normalize('NFC')
        .replace(/[^\p{L}\p{N}]+/gu, ' ')
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
    // Из ячейки «DEPOSIT LIMITS min 15 - max 20000 WITHDRAWAL min 50 - …»
    // оставляем раздел депозита
    function depositSection(t) {
      const wd = t.search(/withdraw|вывод|payout/i);
      if (wd < 0) return t;
      const dep = t.search(/deposit|депозит|пополн/i);
      if (dep > wd) {
        // вывод идёт первым: берём от депозита до следующего упоминания вывода
        const rest = t.slice(dep);
        const wd2 = rest.search(/withdraw|вывод|payout/i);
        return wd2 > 0 ? rest.slice(0, wd2) : rest;
      }
      return t.slice(0, wd);
    }

    function parseLimitText(text) {
      const t = depositSection(normSpace(text));
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
    const COL_SKIP_RE = /вывод|withdraw|payout|коммент|comment|примеч|note|контакт|contact|ссылк|link|комисс|fee|переключ|доп|настройк|страна|country|субы|субагенты|subs/;
    const COL_LIMIT_RE = /лимит|limit|min|max|мин|макс|сумм|amount|range|диапазон|депозит|deposit/;
    const COL_COUNTRY_RE = /страна|country/;
    const HEADER_FIRST_CELL_RE = /^(метод|методы|method|methods|субагент|subagent|name|название|наименование)$/;

    // Строка заголовков таблицы Confluence: первая ячейка — «Метод» и т.п.,
    // либо (для первой строки таблицы) в строке нет ни одной цифры
    function isHeaderRow(cells, isFirst) {
      if (!cells.length) return false;
      if (HEADER_FIRST_CELL_RE.test(normName(cells[0]))) return true;
      return !!isFirst && cells.length >= 2 && cells.every(c => !/\d/.test(c));
    }

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

    // ── Страны: название в тикете (англ.) ↔ таблица/колонка в Confluence ──
    const COUNTRY_ALIASES = {
      morocco: ['morocco', 'марокко', 'maroc'],
      egypt: ['egypt', 'египет'],
      uae: ['uae', 'united arab emirates', 'оаэ', 'эмираты', 'объединенные арабские эмираты'],
      somalia: ['somalia', 'сомали'],
      lebanon: ['lebanon', 'ливан'],
      syria: ['syria', 'сирия', 'syrian arab republic'],
      iraq: ['iraq', 'ирак'],
      jordan: ['jordan', 'иордания'],
      saudi: ['saudi arabia', 'ksa', 'саудовская аравия', 'сауд аравия', 'саудия'],
      kuwait: ['kuwait', 'кувейт'],
      qatar: ['qatar', 'катар'],
      bahrain: ['bahrain', 'бахрейн'],
      oman: ['oman', 'оман'],
      yemen: ['yemen', 'йемен'],
      libya: ['libya', 'ливия'],
      tunisia: ['tunisia', 'тунис'],
      algeria: ['algeria', 'алжир'],
      sudan: ['sudan', 'судан'],
      palestine: ['palestine', 'палестина'],
      mauritania: ['mauritania', 'мавритания'],
      djibouti: ['djibouti', 'джибути'],
      turkey: ['turkey', 'turkiye', 'турция'],
      iran: ['iran', 'иран'],
      afghanistan: ['afghanistan', 'афганистан'],
      pakistan: ['pakistan', 'пакистан'],
      bangladesh: ['bangladesh', 'бангладеш'],
      india: ['india', 'индия'],
      // СНГ
      kyrgyzstan: ['kyrgyzstan', 'kyrgyz republic', 'kirgizia', 'кыргызстан', 'киргизия', 'киргизстан'],
      kazakhstan: ['kazakhstan', 'казахстан'],
      uzbekistan: ['uzbekistan', 'узбекистан'],
      tajikistan: ['tajikistan', 'таджикистан'],
      turkmenistan: ['turkmenistan', 'туркменистан'],
      azerbaijan: ['azerbaijan', 'азербайджан'],
      armenia: ['armenia', 'армения'],
      georgia: ['georgia', 'грузия'],
      russia: ['russia', 'russian federation', 'россия', 'рф'],
      ukraine: ['ukraine', 'украина'],
      belarus: ['belarus', 'беларусь', 'белоруссия'],
      moldova: ['moldova', 'молдова', 'молдавия'],
      mongolia: ['mongolia', 'монголия'],
      // Африка
      nigeria: ['nigeria', 'нигерия'],
      kenya: ['kenya', 'кения'],
      ethiopia: ['ethiopia', 'эфиопия'],
      ghana: ['ghana', 'гана'],
      tanzania: ['tanzania', 'танзания'],
      uganda: ['uganda', 'уганда'],
      cameroon: ['cameroon', 'камерун'],
      senegal: ['senegal', 'сенегал'],
      ivorycoast: ['cote d ivoire', 'côte d ivoire', 'ivory coast', 'кот д ивуар'],
      zambia: ['zambia', 'замбия'],
      rwanda: ['rwanda', 'руанда'],
      southafrica: ['south africa', 'юар', 'южная африка'],
      chad: ['chad', 'чад'],
      niger: ['niger', 'нигер'],
      mali: ['mali', 'мали'],
      burkina: ['burkina faso', 'буркина фасо'],
      guinea: ['guinea', 'гвинея'],
      benin: ['benin', 'бенин'],
      togo: ['togo', 'того'],
      congo: ['congo', 'конго'],
      // Азия
      indonesia: ['indonesia', 'индонезия'],
      malaysia: ['malaysia', 'малайзия'],
      philippines: ['philippines', 'филиппины'],
      vietnam: ['vietnam', 'viet nam', 'вьетнам'],
      thailand: ['thailand', 'таиланд', 'тайланд'],
      myanmar: ['myanmar', 'мьянма'],
      cambodia: ['cambodia', 'камбоджа'],
      nepal: ['nepal', 'непал'],
      srilanka: ['sri lanka', 'шри ланка'],
      // Латинская Америка
      brazil: ['brazil', 'бразилия'],
      argentina: ['argentina', 'аргентина'],
      chile: ['chile', 'чили'],
      peru: ['peru', 'перу'],
      colombia: ['colombia', 'колумбия'],
      mexico: ['mexico', 'мексика'],
      ecuador: ['ecuador', 'эквадор'],
      bolivia: ['bolivia', 'боливия'],
      venezuela: ['venezuela', 'венесуэла'],
      paraguay: ['paraguay', 'парагвай'],
      uruguay: ['uruguay', 'уругвай'],
      nicaragua: ['nicaragua', 'никарагуа'],
      panama: ['panama', 'панама'],
      honduras: ['honduras', 'гондурас'],
      guatemala: ['guatemala', 'гватемала'],
      costarica: ['costa rica', 'коста рика'],
      elsalvador: ['el salvador', 'сальвадор'],
      dominican: ['dominican republic', 'доминиканская республика', 'доминикана'],
      jamaica: ['jamaica', 'ямайка'],
      haiti: ['haiti', 'гаити'],
      trinidad: ['trinidad and tobago', 'trinidad', 'тринидад и тобаго', 'тринидад'],
      belize: ['belize', 'белиз'],
      canada: ['canada', 'канада'],
      usa: ['united states', 'usa', 'сша'],
      malawi: ['malawi', 'малави'],
      mozambique: ['mozambique', 'мозамбик'],
      zimbabwe: ['zimbabwe', 'зимбабве'],
      botswana: ['botswana', 'ботсвана'],
      namibia: ['namibia', 'намибия'],
      sierraleone: ['sierra leone', 'сьерра леоне'],
      liberia: ['liberia', 'либерия'],
      gambia: ['gambia', 'гамбия'],
      madagascar: ['madagascar', 'мадагаскар'],
      taiwan: ['taiwan', 'тайвань'],
      hongkong: ['hong kong', 'гонконг'],
      singapore: ['singapore', 'сингапур'],
      laos: ['laos', 'лаос'],
      // Европа
      cyprus: ['cyprus', 'кипр'],
    };

    function hasWords(text, words) {
      return (' ' + text + ' ').indexOf(' ' + words + ' ') !== -1;
    }

    // Текст → набор канонических стран, которые в нём упомянуты
    const COUNTRY_NORM = Object.keys(COUNTRY_ALIASES).map(k => [k, COUNTRY_ALIASES[k].map(a => normName(a))]);
    function countryKeys(text) {
      const n = normName(text);
      if (!n) return [];
      return COUNTRY_NORM.filter(([, aliases]) => aliases.some(a => hasWords(n, a))).map(([k]) => k);
    }

    // tables: [{ heading, headers:[], rows:[{ cells:[text], striked }] }]
    const COL_SWITCH_RE = /переключ|switch/;
    const COL_SUBS_RE = /(^| )(субы|субагенты|subs|subagents)( |$)/;

    // Ячейка «Субы»: список субагентов строки и блоки «исключения по
    // лимитам (500-30000 KGS): <субы>» со своим лимитом. Читается из
    // Confluence при каждой загрузке — ничего не зашито в скрипт.
    // «Aggron Kyrgyzstan MBANK 67653» и «19902 (Huskar D17 Tunisia)» →
    // имя без номера + ID
    function subItem(line) {
      const norm = normName(line);
      if (!norm || !/[a-zа-яё]{2}/.test(norm)) return null;
      const words = norm.split(' ');
      const id = words.find(w => /^\d{3,}$/.test(w)) || '';
      return { norm, base: words.filter(w => !/^\d{3,}$/.test(w)).join(' '), id };
    }

    // Список агентов внутри любой ячейки строки: «Агенты:» и дальше строки
    // вида «19902 (Huskar D17 Tunisia)» / «Name Country Method 12345»
    const AGENTS_HDR_RE = /^(агенты|субагенты|субы|agents|subagents)\s*:?\s*$/i;
    const AGENT_LINE_RE = /^\s*(\d{3,}\s*\(.+\)|.*[a-zа-яё].*\s\d{3,})\s*$/i;
    function parseAgentLists(text) {
      const lines = String(text || '').split('\n').map(normSpace).filter(Boolean);
      const out = [];
      for (let i = 0; i < lines.length; i++) {
        if (!AGENTS_HDR_RE.test(lines[i])) continue;
        for (let j = i + 1; j < lines.length && AGENT_LINE_RE.test(lines[j]); j++) {
          const it = subItem(lines[j]);
          if (it) out.push(it);
        }
      }
      return out;
    }

    function parseSubsCell(text) {
      const lines = String(text || '').split('\n').map(normSpace).filter(Boolean);
      const out = { subs: [], exceptions: [] };
      let cur = null;
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (/исключени|exception/i.test(line)) {
          let hdr = line;
          let lim = parseLimitText(hdr);
          // заголовок мог переноситься на следующую строку: «… (500-30000» / «KGS):»
          if (lim.min == null && lim.max == null && lines[i + 1]) {
            const two = line + ' ' + lines[i + 1];
            const lim2 = parseLimitText(two);
            if (lim2.min != null || lim2.max != null) { hdr = two; lim = lim2; i++; }
          }
          cur = { min: lim.min, max: lim.max, raw: hdr, subs: [] };
          out.exceptions.push(cur);
          continue;
        }
        // хвост перенесённого заголовка исключения: «KGS):»
        if (cur && !cur.subs.length && /^[^()]*\)\s*:?\s*$/.test(line)) { cur.raw += ' ' + line; continue; }
        const item = subItem(line);
        if (item && item.norm.length >= 4) (cur ? cur.subs : out.subs).push(item);
      }
      return out;
    }

    // Subagent тикета есть в списке? По имени (номер в конце списка не
    // обязателен) или по ID субагента
    function subListed(item, sn, sid) {
      if (sid && item.id === sid) return true;
      if (!sn) return false;
      return item.base === sn || item.norm === sn || (sn.length >= 6 && hasWords(item.norm, sn));
    }

    function parenTexts(name) {
      return (String(name).match(/\(([^)]*)\)/g) || []).map(x => normName(x)).filter(Boolean);
    }

    // knownBrands — бренды из CONFIG; к ним добавляется всё, что в скобках
    // названий встречается 2+ раз («(1XCasino)», «(BizBet)»…)
    function buildLimitIndex(tables, knownBrands) {
      const parenCount = {};
      tables.forEach(tbl => tbl.rows.forEach(row => {
        parenTexts(row.cells[0]).forEach(p => { parenCount[p] = (parenCount[p] || 0) + 1; });
      }));
      const brandSet = {};
      (knownBrands || []).forEach(b => { const n = normName(b); if (n) brandSet[n] = 1; });
      Object.keys(parenCount).forEach(p => { if (parenCount[p] >= 2) brandSet[p] = 1; });
      delete brandSet.webdefault;

      const entries = [];
      tables.forEach(tbl => {
        const countryCol = tbl.headers.findIndex((x, i) => i > 0 && COL_COUNTRY_RE.test(normName(x)));
        const switchCol = tbl.headers.findIndex((x, i) => i > 0 && COL_SWITCH_RE.test(normName(x)));
        const subsCol = tbl.headers.findIndex((x, i) => i > 0 && COL_SUBS_RE.test(normName(x)));
        // проект таблицы — из заголовка: «Египет 1XCasino», «Турция Melbet BT»,
        // «Марокко BizBet Africa (Ar) 365», «… расшарены с рефки 1 на 292».
        // Таблица с одним названием страны — webDefault
        const hn = normName(tbl.heading);
        const headingBrands = Object.keys(brandSet).filter(b => hasWords(hn, b));
        const shared = String(tbl.heading || '').match(/(?:рефк|ref)[^]*?на\s+(\d+)/i);
        const headingRefs = shared ? [shared[1]] : hn.split(' ').filter(w => /^\d{3,}$/.test(w));
        tbl.rows.forEach(row => {
          const name = normSpace(row.cells[0]);
          if (!name) return;
          const norm = normName(name);
          if (!norm) return;
          if (HEADER_FIRST_CELL_RE.test(norm)) return; // повтор строки заголовков
          const rawParts = String(row.cells[0]).split(/[\n,;/|]+/).filter(x => normSpace(x));
          const parts = rawParts.map(normName).filter(Boolean);
          // названия без бренда: «ATTIJARI (1XCasino)» → «attijari»,
          // «Orange Cash (1XCasino) Axinite…» → «orange cash»
          let aliases = [norm].concat(parts,
            rawParts.map(x => normName(x.replace(/\([^)]*\)/g, ' '))),
            rawParts.map(x => normName(x.split('(')[0])));
          // и без общего слова «Bank»: «CDM Bank» → «cdm»
          aliases = aliases.concat(aliases.map(dropGeneric))
            .filter((a, i, arr) => a && arr.indexOf(a) === i);
          const lim = parseLimitRow(tbl.headers, row.cells);
          const countryCell = countryCol > 0 ? normSpace(row.cells[countryCol]) : '';
          const heading = tbl.heading || '';
          const switchTokens = switchCol > 0 ? normName(row.cells[switchCol]).split(' ').filter(Boolean) : [];
          // бренд строки: скобки и слова названия + слова «Переключалки»
          const tagSource = parenTexts(name).concat(norm.split(' '), switchTokens);
          const brandTags = tagSource.filter((t, i) => brandSet[t] && tagSource.indexOf(t) === i);
          headingBrands.forEach(b => { if (brandTags.indexOf(b) === -1) brandTags.push(b); });
          // списки субов: колонка «Субы» и «Агенты: …» в любой другой ячейке
          const subsInfo = subsCol > 0 ? parseSubsCell(row.cells[subsCol]) : { subs: [], exceptions: [] };
          row.cells.forEach((c, i) => { if (i > 0 && i !== subsCol) subsInfo.subs.push(...parseAgentLists(c)); });
          entries.push({
            name,
            norm,
            compact: norm.replace(/ /g, ''),
            parts,
            aliases,
            compacts: aliases.map(a => a.replace(/ /g, '')),
            brandTags,
            refIds: headingRefs,
            switchTokens,
            // строка отдельного суба: «Astara Azerbaijan BirBank 53968»
            ownSub: /\s\d{4,}\s*$/.test(name) ? subItem(name) : null,
            min: lim.min,
            max: lim.max,
            heading,
            country: countryCell || heading,
            countryText: normName(countryCell + ' ' + heading),
            countryKeys: countryKeys(countryCell + ' ' + heading),
            striked: !!row.striked,
            raw: row.cells.slice(1).map(normSpace).filter(Boolean).join(' | '),
            rawNorm: normName(row.cells.slice(1).join(' ')),
            subsInfo,
          });
        });
      });
      entries.brands = Object.keys(brandSet);
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

    // Строки Confluence для одного варианта названия: точное → часть
    // ячейки → вхождение целыми словами (самые длинные названия)
    // Строки Confluence для одного варианта названия — уровни по
    // убыванию точности: точное → без пробелов/апострофов → название из
    // Confluence входит в запрос (самые длинные) → запрос входит в название
    // Общие слова, без которых название метода остаётся узнаваемым
    const GENERIC_WORDS = ['bank', 'банк'];
    function dropGeneric(n) {
      return n.split(' ').filter(w => GENERIC_WORDS.indexOf(w) === -1).join(' ');
    }

    // Слова, которые не могут быть названием метода: страны и валюты
    const COUNTRY_WORDS = {};
    Object.keys(COUNTRY_ALIASES).forEach(k => COUNTRY_ALIASES[k].forEach(a => { COUNTRY_WORDS[normName(a)] = 1; }));
    function isStopName(x, codes) {
      return !!COUNTRY_WORDS[x] || (codes || []).indexOf(x) !== -1;
    }

    // codes — коды валют тикета (в нижнем регистре)
    function candidateTiers(entries, n, codes) {
      const tiers = [];
      tiers.push(entries.filter(e => e.aliases.indexOf(n) !== -1));
      const nc = n.replace(/ /g, '');
      tiers.push(entries.filter(e => e.compacts.indexOf(nc) !== -1)); // «L'bankalik» = «Lbankalik»
      if (n.length >= 3 && !isStopName(n, codes)) {
        // по уровню на каждую длину названия, от самых длинных: если
        // «Al Barid Bank (BizBet)» чужого проекта, следующим уровнем
        // проверяется «Al Barid»
        const inner = [];
        entries.forEach(e => {
          const len = e.aliases
            .filter(x => x.length >= 3 && !isStopName(x, codes) && hasWords(n, x))
            .reduce((m, x) => Math.max(m, x.length), 0);
          if (len) inner.push({ e, len });
        });
        inner.map(x => x.len)
          .filter((l, i, arr) => arr.indexOf(l) === i)
          .sort((a, b) => b - a)
          .forEach(l => tiers.push(inner.filter(x => x.len === l).map(x => x.e)));
        tiers.push(entries.filter(e => e.aliases.some(x => hasWords(x, n))));
      }
      return tiers.filter(t => t.length);
    }

    function findCandidates(entries, n) {
      return candidateTiers(entries, n)[0] || [];
    }

    function ticketCodes(ticket) {
      return [ticket.currency].concat(ticket.currencies || [])
        .map(c => normSpace(c).toLowerCase()).filter(Boolean);
    }

    // Варианты названия для поиска: method, method без кода валюты, subagent
    function buildQueries(ticket) {
      const out = [];
      const add = (text, via) => {
        const n = normName(text);
        if (n && !out.some(q => q.n === n)) out.push({ n, via, text: normSpace(text) });
      };
      const codes = ticketCodes(ticket);
      add(ticket.method, 'method');
      if (normSpace(ticket.method)) {
        const stripped = normSpace(ticket.method).split(' ')
          .filter(w => !/^[A-Z]{3}$/.test(w) && codes.indexOf(w.toLowerCase()) === -1)
          .join(' ');
        add(stripped, 'method');
        add(dropGeneric(normName(stripped)), 'method'); // «Optima Bank» → «optima»
      }
      add(ticket.subagent, 'subagent');
      // части Subagent («Orange Baojun» → «orange», «baojun»): подряд идущие
      // слова, от длинных к коротким; страны, валюты и «bank» пропускаем
      const words = normName(ticket.subagent).split(' ').filter(Boolean);
      for (let len = words.length - 1; len >= 1; len--) {
        for (let i = 0; i + len <= words.length; i++) {
          const part = words.slice(i, i + len).join(' ');
          if (part.length < 3 || isStopName(part, codes) || GENERIC_WORDS.indexOf(part) !== -1) continue;
          add(part, 'subagent');
        }
      }
      return out;
    }

    // Проект тикета (nameRef/refId) → строки Confluence этого проекта.
    // webDefault — строки без бренда (или с webdefault в «Переключалке»);
    // другой проект — строки с его названием или refId в «Переключалке».
    // Проекты, которые считаются webDefault: сам webDefault и зеркала 1xbet
    // («1xbet22.com» → «1xbet22com»). Строка — точное совпадение, RegExp — test
    const DEFAULT_WEBDEFAULT_REFS = ['webdefault', /1xbet/];

    function isWebDefaultRef(ref, refId, patterns) {
      if (!ref) return refId === '1';
      return (patterns || DEFAULT_WEBDEFAULT_REFS).some(p => (p instanceof RegExp ? p.test(ref) : normName(p).replace(/ /g, '') === ref));
    }

    function brandFilter(cands, ticket) {
      const ref = normName(ticket.ref).replace(/ /g, '');
      const refId = normSpace(ticket.refId);
      if (!ref && !refId) return { list: cands, matched: false };
      const isProject = e => e.brandTags.length > 0 || (e.refIds || []).length > 0;
      // «webdefault» или «default» в «Переключалке» («jibi default») —
      // явная строка webDefault, если она не из таблицы другого проекта
      const webDefaultRows = () => {
        const plain = cands.filter(e => !isProject(e));
        const explicit = plain.filter(e => e.switchTokens.indexOf('webdefault') !== -1 || e.switchTokens.indexOf('default') !== -1);
        return explicit.length ? explicit : plain;
      };
      if (isWebDefaultRef(ref, refId, ticket.webDefaultRefs)) {
        const list = webDefaultRows();
        return list.length ? { list, matched: true } : { list: cands, matched: false };
      }
      const list = cands.filter(e => (ref && (e.brandTags.indexOf(ref) !== -1 || e.compact.indexOf(ref) !== -1))
        || (refId && (e.switchTokens.indexOf(refId) !== -1 || (e.refIds || []).indexOf(refId) !== -1)));
      if (list.length) return { list, matched: true };
      // у проекта нет своей строки — берём webDefault (проект не совпал →
      // при сумме в лимите тикет уйдёт в «Проверить вручную»)
      const fallback = webDefaultRows();
      return { list: fallback.length ? fallback : cands, matched: false };
    }

    // Строки, название которых целиком есть в Subagent (самые длинные):
    // «Arcturus CLIQ Jordan» → «Cliq». Только с распознанным лимитом и
    // только в стране тикета
    function subagentNamedRows(entries, ticket, codes, tk, tn) {
      const sn = normName(ticket.subagent);
      if (!sn || !tn) return null;
      let bestLen = 0;
      const inner = [];
      entries.forEach(e => {
        const len = e.aliases
          .filter(a => a.length >= 3 && !isStopName(a, codes) && GENERIC_WORDS.indexOf(a) === -1 && hasWords(sn, a))
          .reduce((m, a) => Math.max(m, a.length), 0);
        if (len) { inner.push({ e, len }); bestLen = Math.max(bestLen, len); }
      });
      const list = inner.filter(x => x.len === bestLen).map(x => x.e)
        .filter(e => (tk.length ? tk.some(k => e.countryKeys.indexOf(k) !== -1) : hasWords(e.countryText, tn)));
      if (!list.length) return null;
      const bf = brandFilter(list, ticket);
      if (!bf.list.some(hasLimit)) return null;
      return { cands: bf.list, countryMatched: true, brandMatched: bf.matched };
    }

    // ticket: { method, subagent, country, currency, ref, refId } (или строка)
    // → { status: 'ok'|'nolimit'|'unparsed'|'ambiguous'|'nosubagent',
    //     limit, entries, via, countryMatched }
    function matchLimit(entries, overrides, ticket) {
      if (typeof ticket === 'string') ticket = { subagent: ticket };
      const queries = buildQueries(ticket);
      if (!queries.length) return { status: 'nosubagent' };
      for (const q of queries) {
        if (overrides && overrides[q.n]) {
          const o = overrides[q.n];
          return {
            status: 'ok', via: q.via, countryMatched: true,
            limit: { name: q.text, min: o.min, max: o.max, heading: 'CONFIG.limitOverrides', country: '', raw: '' },
          };
        }
      }
      // Перебираем варианты названия и уровни совпадения; лучший — где
      // совпали и страна (2 балла), и проект (1 балл); при равенстве —
      // самый точный (первый) вариант
      const hasRef = !!(normName(ticket.ref) || normSpace(ticket.refId));
      const codes = ticketCodes(ticket);
      const tk = countryKeys(ticket.country);
      const tn = normName(ticket.country);
      const maxScore = (tn ? 2 : 0) + (hasRef ? 1 : 0);
      // Внутри одного варианта названия строка с распознанным лимитом
      // важнее строки «лимиты в строчке ниже»; между вариантами решает
      // только балл (иначе ник агента мог бы перебить нужный метод)
      let best = null;
      for (const q of queries) {
        let qBest = null;
        for (const tier of candidateTiers(entries, q.n, codes)) {
          let list = tier;
          let countryMatched = false;
          if (tn) {
            const byCountry = list.filter(e => (tk.length
              ? tk.some(k => e.countryKeys.indexOf(k) !== -1)
              : hasWords(e.countryText, tn)));
            if (byCountry.length) { list = byCountry; countryMatched = true; }
          }
          const bf = brandFilter(list, ticket);
          const score = (countryMatched ? 2 : 0) + (bf.matched ? 1 : 0);
          const parsedAny = bf.list.some(hasLimit);
          if (!qBest || score > qBest.score || (score === qBest.score && parsedAny && !qBest.parsedAny)) {
            qBest = { cands: bf.list, via: q.via, countryMatched, brandMatched: bf.matched, score, parsedAny };
          }
          if (score >= maxScore && parsedAny) break;
        }
        if (qBest && (!best || qBest.score > best.score)) best = qBest;
        if (best && best.score >= maxScore) break;
      }
      if (!best) return { status: 'nolimit' };

      // Subagent прямо называет метод («Arcturus CLIQ Jordan» → «Cliq»), а
      // method указывает на другую строку («Bank al Etihad») — важнее Subagent
      let info = '';
      const subRows = subagentNamedRows(entries, ticket, codes, tk, tn);
      if (subRows && normSpace(ticket.method) && best.via === 'method'
          && !subRows.cands.some(e => best.cands.indexOf(e) !== -1)
          && (subRows.brandMatched || !best.brandMatched) && (subRows.countryMatched || !best.countryMatched)) {
        info = `лимит по Subagent: method «${normSpace(ticket.method)}» указывает на «${best.cands[0].name}»`;
        best = { cands: subRows.cands, via: 'subagent', countryMatched: subRows.countryMatched, brandMatched: subRows.brandMatched };
      }
      // валюта тикета: «Sham Cash» + USD → строка «Sham Cash USD», а не SYP.
      // Применяется, только если отсекает часть строк и что-то оставляет
      let cands = best.cands;
      // Сначала валюта платежа (currencyKode), потом все коды (валюта
      // счёта currencyUserKode может отличаться: платёж SYP, счёт USD)
      const primaryCode = normSpace(ticket.currency).toLowerCase();
      for (const set of [primaryCode ? [primaryCode] : [], codes]) {
        if (!set.length || cands.length < 2) continue;
        const byCur = cands.filter(e => set.some(c => hasWords(e.norm, c) || hasWords(e.rawNorm, c)));
        if (byCur.length && byCur.length < cands.length) { cands = byCur; break; }
      }
      const via = best.via;
      const countryMatched = best.countryMatched;
      const brandMatched = best.brandMatched;
      const sn = normName(ticket.subagent);
      const sid = normSpace(ticket.subagentId);
      let note = '';

      // строки отдельного суба («Astara Azerbaijan BirBank 53968»): свой суб —
      // только они, чужой — отбрасываем, если что-то остаётся
      if (cands.some(e => e.ownSub) && (sn || sid)) {
        const mine = cands.filter(e => e.ownSub && subListed(e.ownSub, sn, sid));
        if (mine.length) cands = mine;
        else {
          const rest = cands.filter(e => !e.ownSub);
          if (rest.length) cands = rest;
        }
      }

      // Отличительные слова: строки одного метода различаются словами
      // («D17 QR-code Tunisia» / «D17 Tunisia»). Берём строку, чьё
      // отличительное слово есть в method/Subagent тикета; если таких слов
      // в тикете нет — строку без них (обычный D17, а не QR)
      const distinctLimits = list => Object.keys(list.filter(hasLimit).reduce((m, e) => { m[e.min + '|' + e.max] = 1; return m; }, {})).length;
      if (distinctLimits(cands) > 1) {
        const tw = normName(ticket.method + ' ' + ticket.subagent).split(' ');
        // бренды сюда не входят — проект выбирает brandFilter
        const wordsOf = e => {
          const bw = e.brandTags.join(' ').split(' ');
          return e.norm.split(' ').filter(w => w.length >= 2 && !/^\d+$/.test(w) && !isStopName(w, codes)
            && GENERIC_WORDS.indexOf(w) === -1 && bw.indexOf(w) === -1);
        };
        const sets = cands.map(wordsOf);
        const all = sets.reduce((u, ws) => u.concat(ws), []);
        const distinguishing = all.filter((w, i) => all.indexOf(w) === i && !sets.every(ws => ws.indexOf(w) !== -1));
        if (distinguishing.length) {
          const hit = cands.filter((e, i) => sets[i].some(w => distinguishing.indexOf(w) !== -1 && tw.indexOf(w) !== -1));
          const plain = cands.filter((e, i) => !sets[i].some(w => distinguishing.indexOf(w) !== -1));
          if (hit.length && hit.length < cands.length) cands = hit;
          else if (!hit.length && plain.length && plain.length < cands.length) cands = plain;
        }
      }

      // «Субы» / «Агенты:»: исключение со своим лимитом → строка, где указан суб
      const hasList = e => e.subsInfo && (e.subsInfo.subs.length || e.subsInfo.exceptions.length);
      const withSubs = cands.filter(hasList);
      let subMatched = false;
      if ((sn || sid) && withSubs.length) {
        for (const e of cands) {
          for (const ex of (e.subsInfo ? e.subsInfo.exceptions : [])) {
            if ((ex.min != null || ex.max != null) && ex.subs.some(it => subListed(it, sn, sid))) {
              return {
                status: 'ok', via, countryMatched, brandMatched, subMatched: true,
                limit: Object.assign({}, e, { min: ex.min, max: ex.max, name: e.name + ' — исключение из «Субы»', raw: ex.raw + ' | ' + e.raw }),
              };
            }
          }
        }
        const bySub = cands.filter(e => e.subsInfo && e.subsInfo.subs.some(it => subListed(it, sn, sid)));
        if (bySub.length) { cands = bySub; subMatched = true; }
        else {
          // суба нет ни в одном списке: строки со списком — для своих агентов,
          // остаётся строка без списка (если есть). Ненадёжно → примечание
          const noList = cands.filter(e => !hasList(e));
          if (noList.length && noList.length < cands.length) {
            note = `суба нет в списке агентов строки «${withSubs[0].name}», взята строка без списка`;
            cands = noList;
          }
        }
      }
      const subsChecked = withSubs.length > 0 && !subMatched && !note;

      // последний разделитель: слово из Subagent в названии строки
      // («Anamudi Azerbaijan BirBank» → «Bank köçürməsi - BirBank»)
      if (sn && distinctLimits(cands) > 1) {
        const words = sn.split(' ').filter(w => w.length >= 4 && !isStopName(w, codes) && GENERIC_WORDS.indexOf(w) === -1);
        const byWord = cands.filter(e => words.some(w => e.aliases.some(a => hasWords(a, w))));
        if (byWord.length && byWord.length < cands.length) cands = byWord;
      }

      const active = cands.filter(e => !e.striked);
      const pool = active.length ? active : cands;
      const parsed = pool.filter(hasLimit);
      if (!parsed.length) return { status: 'unparsed', entries: pool, via, countryMatched, brandMatched, subMatched, note, info };
      const distinct = {};
      parsed.forEach(e => { if (!distinct[e.min + '|' + e.max]) distinct[e.min + '|' + e.max] = e; });
      const keys = Object.keys(distinct);
      if (keys.length > 1) return { status: 'ambiguous', entries: parsed, via, countryMatched, brandMatched, subMatched, subsChecked, note, info };
      return { status: 'ok', limit: distinct[keys[0]], via, countryMatched, brandMatched, subMatched, note, info };
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
        keys: ['amountByReceipt', 'amount_by_receipt', 'amountReceipt', 'receiptAmount', 'sumByReceipt', 'sumReceipt', 'amount'],
        pats: [[/receipt/i, /amount|sum/i], [/amount/i], [/^sum/i]],
        valid: v => v !== '' && v != null && typeof v !== 'boolean' && parseAmount(v) != null,
      },
      method: {
        keys: ['method', 'nameMethod', 'methodName', 'paymentMethod', 'namePaymentMethod'],
        pats: [[/method/i, /name/i], [/method/i]],
        valid: isText,
      },
      subagent: {
        keys: ['nameSubagent', 'subagentName', 'nameSubAgent', 'subAgentName', 'nameQuerySubagent', 'subagent', 'subAgent'],
        pats: [[/sub_?agent/i, /name/i], [/sub_?agent/i]],
        valid: isText,
      },
      ref: {
        keys: ['nameRef', 'refName', 'nameProject'],
        pats: [[/ref/i, /name/i]],
        valid: isText,
      },
      subagentId: {
        keys: ['subAgenId', 'subAgentId', 'idSubagent', 'idSubAgent', 'subagentId'],
        pats: [[/^sub_?agen_?t?_?id$/i], [/^id_?sub_?agent$/i]],
        valid: v => /^\d+$/.test(String(v == null ? '' : v).trim()),
      },
      refId: {
        keys: ['refId', 'idRef', 'ref_id'],
        pats: [[/^ref_?id$/i], [/^id_?ref$/i]],
        valid: v => /^\d+$/.test(String(v == null ? '' : v).trim()),
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
        keys: ['currencyKode', 'nameCurrency', 'currencyCode', 'codeCurrency', 'currency'],
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
      parseLimitText, parseLimitRow, parseSubsCell, isHeaderRow, countryKeys, buildQueries, findCandidates, brandFilter,
      buildLimitIndex, normOverrides, matchLimit, evaluate,
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
      const rowCells = tr => Array.from(tr.children).filter(c => /^(TD|TH)$/.test(c.tagName));
      // Заголовок: строка в <thead>, строка из <th> или первая строка без
      // цифр / с «Метод» в первой ячейке (в Confluence шапка бывает из <td>)
      let headerRow = allRows.find(tr => tr.closest('thead'));
      if (!headerRow) {
        const first = allRows[0];
        const cells = rowCells(first);
        const thCount = cells.filter(c => c.tagName === 'TH').length;
        if (thCount >= Math.ceil(cells.length / 2) || Core.isHeaderRow(cells.map(c => c.textContent.trim()), true)) {
          headerRow = first;
        }
      }
      const headers = headerRow ? rowCells(headerRow).map(c => c.textContent.trim()) : [];
      const rows = [];
      allRows.forEach(tr => {
        if (tr === headerRow || tr.closest('thead')) return;
        const cells = rowCells(tr);
        if (cells.length < 2) return;
        const texts = cells.map(cellText);
        if (Core.isHeaderRow(texts, false)) return; // повтор шапки внутри таблицы
        rows.push({ cells: texts, striked: isStriked(cells[0]) });
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
    const entries = Core.buildLimitIndex(tables, CONFIG.knownBrands);
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
    manual: [], // в лимите, но строка Confluence взята другой страны/проекта
    startedAt: null,
    finishedAt: null,
    complete: false,
    error: '',
  };

  const UNCHECKED_REASONS = {
    nosubagent: 'у тикета нет ни method, ни subagent',
    nolimit: 'метод не найден в Confluence',
    unparsed: 'лимит в Confluence не распознан',
    ambiguous: 'несколько строк с разными лимитами',
    noamount: 'не удалось прочитать сумму',
  };

  function isWantedStatus(statusText) {
    return CONFIG.statuses.map(Core.normStatus).indexOf(Core.normStatus(statusText)) !== -1;
  }

  // Строка списка → запись отчёта + вердикт: 'below'|'above'|'ok'|'unchecked'
  function checkRow(row, overrides) {
    const k = run.keys;
    const val = key => (key ? Core.normSpace(row[key]) : '');
    const rec = {
      ticketId: String(row[k.ticketId] == null ? '' : row[k.ticketId]).trim(),
      method: val(k.method),
      subagent: val(k.subagent),
      ref: val(k.ref),
      refId: val(k.refId),
      subagentId: val(k.subagentId),
      amountRaw: row[k.amount],
      amount: Core.parseAmount(row[k.amount]),
      currency: val(k.currency),
      country: val(k.country),
      userId: val(k.userId),
      status: val(k.status),
      min: null, max: null, diff: null, heading: '', limitName: '', limitCountry: '', limitRaw: '',
      countryMatched: false, brandMatched: false, matchedBy: '', reason: '',
    };
    // все коды валют строки (currencyKode, currencyUserKode…) — чтобы
    // отрезать их от method: " Sham Cash SYP" → "Sham Cash"
    const currencies = Object.keys(row)
      .filter(x => /curr/i.test(x) && /^[a-z]{3}$/i.test(Core.normSpace(row[x])))
      .map(x => Core.normSpace(row[x]));

    const m = Core.matchLimit(limitIndex.entries, overrides, {
      method: rec.method, subagent: rec.subagent, country: rec.country, currency: rec.currency, currencies,
      ref: rec.ref, refId: rec.refId, webDefaultRefs: CONFIG.webDefaultRefs, subagentId: rec.subagentId,
    });
    rec.brandMatched = !!m.brandMatched;
    rec.subMatched = !!m.subMatched;
    rec.note = m.note || '';
    rec.info = m.info || ''; // справка, на ручную проверку не отправляет
    if (m.status === 'ambiguous' && m.subsChecked) {
      m.reasonExtra = `Субагента «${rec.subagent}» нет в колонке «Субы» ни одной из строк`;
    }
    rec.matchedBy = m.via || '';
    if (m.status !== 'ok') {
      rec.reason = UNCHECKED_REASONS[m.status] || m.status;
      if (m.reasonExtra) rec.reason += ` (${m.reasonExtra})`;
      if (m.entries && m.entries.length) {
        rec.reason += ': ' + m.entries.slice(0, 4)
          .map(e => `«${e.name}» (${e.country || '?'}) ${e.raw}`.slice(0, 160)).join(' ; ');
      }
      return { rec, verdict: 'unchecked' };
    }
    Object.assign(rec, {
      min: m.limit.min, max: m.limit.max, heading: m.limit.heading, limitName: m.limit.name,
      limitCountry: m.limit.country || m.limit.heading || '', limitRaw: m.limit.raw, countryMatched: !!m.countryMatched,
    });
    if (rec.amount == null) {
      rec.reason = UNCHECKED_REASONS.noamount + ` («${Core.normSpace(rec.amountRaw)}»)`;
      return { rec, verdict: 'unchecked' };
    }
    const ev = Core.evaluate(rec.amount, m.limit);
    rec.diff = ev.diff;
    return { rec, verdict: ev.verdict };
  }

  function processRow(row, overrides) {
    if (!isWantedStatus(row[run.keys.status])) return;
    run.matched++;
    const { rec, verdict } = checkRow(row, overrides);
    if (verdict === 'below') run.below.push(rec);
    else if (verdict === 'above') run.above.push(rec);
    else if (verdict === 'unchecked') run.unchecked.push(rec);
    else if (verdict === 'ok') {
      // в лимите, но лимит взят из строки другой страны или проекта —
      // результат ненадёжен, человеку стоит посмотреть самому
      const why = [];
      if (!rec.countryMatched) why.push(`страна не совпала (тикет: ${rec.country || '—'}, Confluence: ${rec.limitCountry || '—'})`);
      if ((rec.ref || rec.refId) && !rec.brandMatched) why.push(`проект не совпал (тикет: ${rec.ref || rec.refId})`);
      if (rec.note) why.push(rec.note);
      if (why.length) { rec.reason = why.join('; '); run.manual.push(rec); }
    }
  }

  async function startRun() {
    if (run.running) return;
    if (!listUrl || !listPayload) {
      showToast('Сначала нажмите APPLY в фильтре — скрипт перехватит запрос списка', 'err');
      return;
    }
    Object.assign(run, {
      running: true, stopRequested: false, page: 0, seen: 0, matched: 0,
      below: [], above: [], unchecked: [], manual: [], startedAt: new Date(), finishedAt: null, complete: false, error: '',
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
          const missing = ['ticketId', 'status', 'amount'].filter(f => !run.keys[f]);
          if (!run.keys.method && !run.keys.subagent) missing.push('method/subagent');
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
      setStatus(`✅ Готово: ${run.page} стр., ${run.seen} тикетов, в нужном статусе: ${run.matched}, нарушений: ${total}`);
      showToast(`Готово. Ниже лимита: ${run.below.length}, выше: ${run.above.length}`, 'ok');
    }
  }

  // ------------------------------------------------------------------
  // EXCEL
  // ------------------------------------------------------------------
  function downloadExcel() {
    if (typeof XLSX === 'undefined') { showToast('Библиотека XLSX не загрузилась (@require)', 'err'); return; }
    const total = run.below.length + run.above.length + run.unchecked.length + run.manual.length;
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
    const vHeader = ['№', 'Ticket ID', 'Method', 'Subagent', 'Ref', 'Amount by receipt', 'Валюта', 'Min', 'Max', 'Отклонение', 'Country', 'User ID', 'Status', 'Метод в Confluence', 'Страна (Confluence)', 'Страна совпала', 'Проект совпал', 'Лимит (как в Confluence)', 'Примечание'];
    const vRow = (r, i) => [i + 1, r.ticketId, r.method, r.subagent, r.ref, num(r.amount), r.currency, num(r.min), num(r.max), num(r.diff), r.country, r.userId, r.status, r.limitName, r.limitCountry, r.countryMatched ? 'да' : 'нет', r.brandMatched ? 'да' : 'нет', r.limitRaw, [r.note, r.info].filter(Boolean).join('; ')];
    const uHeader = ['№', 'Ticket ID', 'Method', 'Subagent', 'Ref', 'Amount by receipt', 'Валюта', 'Country', 'User ID', 'Status', 'Причина'];
    const uRow = (r, i) => [i + 1, r.ticketId, r.method, r.subagent, r.ref, r.amount == null ? Core.normSpace(r.amountRaw) : r.amount, r.currency, r.country, r.userId, r.status, r.reason];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet(vHeader, run.below.map(vRow), '2F5597', [5, 7, 8, 9]), 'Ниже лимита');
    XLSX.utils.book_append_sheet(wb, sheet(vHeader, run.above.map(vRow), 'C55A11', [5, 7, 8, 9]), 'Выше лимита');
    XLSX.utils.book_append_sheet(wb, sheet(uHeader, run.unchecked.map(uRow), '7F7F7F', [5]), 'Не удалось проверить');
    const mHeader = ['№', 'Ticket ID', 'Method', 'Subagent', 'Ref', 'Amount by receipt', 'Валюта', 'Min', 'Max', 'Country', 'User ID', 'Status', 'Метод в Confluence', 'Страна (Confluence)', 'Почему проверить', 'Лимит (как в Confluence)'];
    const mRow = (r, i) => [i + 1, r.ticketId, r.method, r.subagent, r.ref, num(r.amount), r.currency, num(r.min), num(r.max), r.country, r.userId, r.status, r.limitName, r.limitCountry, r.reason, r.limitRaw];
    XLSX.utils.book_append_sheet(wb, sheet(mHeader, run.manual.map(mRow), 'BF8F00', [5, 7, 8]), 'Проверить вручную');

    // Сводка по методам (method + страна)
    const byMethod = {};
    const add = (r, key) => {
      const id = (r.method || r.subagent) + '\u0000' + r.country;
      const g = byMethod[id] || (byMethod[id] = { method: r.method || r.subagent, country: r.country, below: 0, above: 0, min: r.min, max: r.max });
      g[key]++;
    };
    run.below.forEach(r => add(r, 'below'));
    run.above.forEach(r => add(r, 'above'));
    const sumRows = Object.values(byMethod)
      .sort((a, b) => (b.below + b.above) - (a.below + a.above))
      .map(g => [g.method, g.country, g.below, g.above, g.below + g.above, num(g.min), num(g.max)]);
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
      ['Проверить вручную (в лимите, но строка другой страны/проекта)', run.manual.length],
      ['Обход завершён', run.complete ? 'да' : 'нет — ' + (run.error || 'остановлен вручную')],
      [],
    ];
    const ws = sheet(['Method', 'Country', 'Ниже', 'Выше', 'Всего', 'Min', 'Max'], sumRows, '375623', [5, 6]);
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
    let det = null;
    if (rows && rows.length) {
      det = detectFields(rows);
      const src = { config: 'CONFIG.fields', screen: 'по таблице на экране', name: 'по названию ключа' };
      out.push('Определённые поля:');
      Object.keys(Core.FIELD_DEFS).forEach(f => {
        const k = det.keys[f];
        out.push(`  ${f.padEnd(9)} → ${k ? k : '— не найдено'}${k ? `   (${src[det.source[f]] || ''}; пример: ${JSON.stringify(rows[0][k])})` : ''}`);
      });
      out.push('');
      log('Первая строка списка:', rows[0], 'поля:', det);
    } else if (listUrl) {
      out.push('Список пуст.');
      out.push('');
    }

    let li = null;
    try {
      li = await loadLimits(false);
      const parsed = li.entries.filter(e => e.min != null || e.max != null);
      out.push(`Confluence: таблиц ${li.tables.length}, строк ${li.entries.length}, лимит распознан в ${parsed.length}`);
      li.tables.slice(0, 5).forEach((t, i) => out.push(`  Таблица ${i + 1} «${t.heading}»: колонки ${JSON.stringify(t.headers)}`));
      out.push('Примеры распознанных строк:');
      parsed.slice(0, 8).forEach(e => out.push(`  ${e.name} [${e.country}] → min ${fmtNum(e.min)}, max ${fmtNum(e.max)}${e.striked ? ' (зачёркнута)' : ''}`));
      out.push(`Бренды (для выбора строки по nameRef): ${(li.entries.brands || []).join(', ') || '—'}`);
      const withSubs = li.entries.filter(e => e.subsInfo && (e.subsInfo.subs.length || e.subsInfo.exceptions.length));
      const exCount = withSubs.reduce((n, e) => n + e.subsInfo.exceptions.length, 0);
      out.push(`Строк со списком субов («Субы» / «Агенты:»): ${withSubs.length}, блоков исключений по лимитам: ${exCount}`);
      out.push(`Таблицы проектов (по заголовку): ${li.tables.map(t => t.heading).filter(h => { const es = li.entries.filter(e => e.heading === h); return es.length && es.every(e => e.refIds.length || e.brandTags.some(b => Core.normName(h).indexOf(b) !== -1)); }).join(', ') || '—'}`);
      withSubs.filter(e => e.subsInfo.exceptions.length).slice(0, 5).forEach(e => e.subsInfo.exceptions.forEach(ex => {
        out.push(`  «${e.name}» [${e.country}]: исключение min ${fmtNum(ex.min)} – max ${fmtNum(ex.max)}, субов: ${ex.subs.length}`);
      }));
      const noCountry = li.tables.filter(t => {
        const ci = t.headers.findIndex((x, i) => i > 0 && /страна|country/i.test(x));
        const cells = ci > 0 ? t.rows.map(r => r.cells[ci] || '').join(' ') : '';
        return !Core.countryKeys(t.heading + ' ' + cells).length;
      }).map(t => t.heading || '(без заголовка)');
      if (noCountry.length) {
        out.push(`Таблицы Confluence без распознанной страны (${noCountry.length}): ${noCountry.slice(0, 30).join(', ')}`);
      }
      const bad = li.entries.filter(e => e.min == null && e.max == null);
      if (bad.length) {
        out.push(`Нераспознанные строки (${bad.length}), первые 10:`);
        bad.slice(0, 10).forEach(e => out.push(`  ${e.name} [${e.country}] → ${e.raw.slice(0, 150)}`));
      }
    } catch (e) {
      out.push('Confluence: ошибка — ' + e.message);
    }
    out.push('');

    // Проверка сопоставления на тикетах нужного статуса со стр. 1
    if (det && li && det.keys.status) {
      const prevKeys = run.keys;
      run.keys = det.keys;
      const overrides = Core.normOverrides(CONFIG.limitOverrides);
      const sample = rows.filter(r => isWantedStatus(r[det.keys.status])).slice(0, 10);
      out.push(`Проверка сопоставления — тикеты ${CONFIG.statuses.join(', ')} со стр. 1 (до 10 шт., найдено ${sample.length}):`);
      out.push('  (формат: #тикет | method | subagent | страна | проект | сумма)');
      if (!sample.length) out.push('  на странице нет тикетов в этом статусе');
      const verdicts = { below: 'НИЖЕ лимита', above: 'ВЫШЕ лимита', ok: 'в лимите' };
      sample.forEach(r => {
        const { rec, verdict } = checkRow(r, overrides);
        const head = `  #${rec.ticketId} | ${rec.method || '—'} | ${rec.subagent || '—'} | ${rec.country || '—'} | ${rec.ref || '—'} | ${fmtNum(rec.amount)} ${rec.currency}`;
        if (verdict === 'unchecked') {
          out.push(`${head}\n      → не проверен: ${rec.reason}`);
        } else {
          out.push(`${head}\n      → «${rec.limitName}» [${rec.limitCountry}${rec.countryMatched ? '' : ', страна НЕ совпала'}${rec.brandMatched ? `, проект ${rec.ref || rec.refId} совпал` : (rec.ref || rec.refId ? ', проект НЕ совпал' : '')}${rec.subMatched ? ', суб найден в «Субы»' : ''}] `
            + `min ${fmtNum(rec.min)} – max ${fmtNum(rec.max)} (по ${rec.matchedBy}) → ${verdicts[verdict]}`);
        }
      });
      run.keys = prevKeys;
      out.push('');
    }

    if (rows && rows.length) {
      out.push('Все ключи первой строки:');
      Object.keys(rows[0]).forEach(k => out.push(`  ${k}: ${JSON.stringify(rows[0][k])}`.slice(0, 200)));
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
    #lmf-body .warn{color:#f0a500;}
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
        <td>${esc(r.method || r.subagent)}<div class="muted">${esc([r.method ? r.subagent : '', r.ref].filter(Boolean).join(' · '))}</div></td>
        <td class="num">${fmtNum(r.amount)} ${esc(r.currency)}</td>
        <td class="num">${fmtNum(r.min)}</td>
        <td class="num">${fmtNum(r.max)}</td>
        <td class="num ${kind === 'below' ? 'neg' : kind === 'above' ? 'pos' : ''}">${kind === 'manual' ? 'в лимите' : (r.diff > 0 ? '+' : '') + fmtNum(r.diff)}</td>
        <td>${esc(r.country)}</td>
        <td>${esc(r.limitName)}<div class="muted">${esc(r.limitCountry)}${r.countryMatched ? '' : ' <span class="warn">⚠ страна не совпала</span>'}${(r.ref || r.refId) && !r.brandMatched ? ' <span class="warn">⚠ проект не совпал</span>' : ''}</div></td>
      </tr>`).join('');
    return `<table><thead><tr><th>№</th><th>Ticket ID</th><th>Method / Subagent</th><th>Amount by receipt</th><th>Min</th><th>Max</th><th>Отклонение</th><th>Country</th><th>Строка Confluence</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  function uTable(list) {
    if (!list.length) return `<div class="empty">${run.startedAt ? 'Все тикеты проверены' : 'Пока пусто'}</div>`;
    const rows = list.map((r, i) => `
      <tr>
        <td class="muted">${i + 1}</td>
        <td><span class="tid" data-tid="${esc(r.ticketId)}" title="Скопировать">${esc(r.ticketId)}</span></td>
        <td>${esc(r.method || r.subagent)}<div class="muted">${esc([r.method ? r.subagent : '', r.ref].filter(Boolean).join(' · '))}</div></td>
        <td class="num">${r.amount == null ? esc(Core.normSpace(r.amountRaw)) : fmtNum(r.amount)} ${esc(r.currency)}</td>
        <td>${esc(r.country)}</td>
        <td>${esc(r.reason)}</td>
      </tr>`).join('');
    return `<table><thead><tr><th>№</th><th>Ticket ID</th><th>Method / Subagent</th><th>Amount by receipt</th><th>Country</th><th>Причина</th></tr></thead><tbody>${rows}</tbody></table>`;
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
      $('t-man').textContent = `Проверить вручную (${run.manual.length})`;
      document.querySelectorAll('.lmf-tab').forEach(b => b.classList.toggle('on', b.dataset.tab === ui.tab));
      const total = run.below.length + run.above.length;
      const cnt = $('lmf-fab').querySelector('.cnt');
      cnt.textContent = total;
      cnt.style.display = total ? '' : 'none';
      $('lmf-xlsx').disabled = !(total + run.unchecked.length + run.manual.length);

      const body = $('lmf-body');
      if (ui.tab === 'below') body.innerHTML = vTable(run.below, 'below');
      else if (ui.tab === 'above') body.innerHTML = vTable(run.above, 'above');
      else if (ui.tab === 'unk') body.innerHTML = uTable(run.unchecked);
      else if (ui.tab === 'man') body.innerHTML = vTable(run.manual, 'manual');
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
          <div class="lmf-m"><b id="m-matched">—</b><span title="${esc(CONFIG.statuses.join(', '))}">Нужный статус</span></div>
          <div class="lmf-m below"><b id="m-below">0</b><span>Ниже лимита</span></div>
          <div class="lmf-m above"><b id="m-above">0</b><span>Выше лимита</span></div>
          <div class="lmf-m unk"><b id="m-unk">0</b><span>Не проверено</span></div>
        </div>
        <div id="lmf-tabs">
          <button class="lmf-tab on" data-tab="below" id="t-below">Ниже лимита</button>
          <button class="lmf-tab" data-tab="above" id="t-above">Выше лимита</button>
          <button class="lmf-tab" data-tab="unk" id="t-unk">Не удалось проверить</button>
          <button class="lmf-tab" data-tab="man" id="t-man" title="В лимите, но лимит взят из строки другой страны или проекта">Проверить вручную</button>
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
