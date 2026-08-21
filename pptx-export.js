/* ---------------------------------------------------------------------------
   Export selected Explore charts to PowerPoint as native, editable charts.

   Loaded by explore.html only. PptxGenJS (libs/pptxgen.bundle.js) is ~460 KB,
   so it is injected on first use rather than on page load.

   The chart data is READ BACK OUT OF THE RENDERED DOM, not recomputed from
   DATA_W1/DATA_W2. Every renderer in explore.html funnels into pairBar() and
   singleBar(), and each bar carries its true value in `data-w`, so scraping
   inherits sort order, top-N slicing, break-by grouping and routed bases for
   free. If the markup those two functions emit ever changes, change
   readRows() below to match.

   Never read style.width or getComputedStyle here: bar widths animate over
   0.9s and are 0 for the first frame. `data-w` is the real number.

   libs/pptxgen.bundle.js is hand-patched: PptxGenJS 4.0.1's chart writer bakes
   a stray apostrophe into the embedded worksheet's <table ref="..."> attribute
   (createExcelWorksheet in src/gen-charts.ts), which PowerPoint tolerates but
   Keynote's stricter XML parser rejects, so exported charts fail to open
   correctly on Mac. Fixed upstream in gitbrent/PptxGenJS PR #1327 (still
   unmerged as of 4.0.1), so the one-character removal is applied by hand here.
   Re-vendoring the bundle from a fresh build will reintroduce the bug unless
   #1327 has landed upstream by then.
   --------------------------------------------------------------------------- */
(function () {
  'use strict';

  // PREVIEW ONLY. This is off on the live site and adds nothing to the page
  // unless ?export=1 is present. index.html forwards the flag from the shell
  // URL to this iframe, so https://brighty36.github.io/world_cup_survey/?export=1
  // turns it on; opening explore.html?export=1 directly works too.
  if (!/(?:^|&)export=1(?:&|$)/.test(location.search.slice(1))) return;

  var LIB = 'libs/pptxgen.bundle.js';

  // Site tokens, flattened to the solid fills a PowerPoint series can hold.
  var THEME = {
    bg: '0E3320',
    panel: '0C2C1A',
    text: 'F1F4EC',
    mute: 'A8B3A6',
    cream: 'F7E6C8',
    magenta: 'D6006C',
    grid: '2E6B3E',
    font: 'Arial'
  };
  var SERIES = {
    t: { name: 'Trade', color: THEME.magenta },
    p: { name: 'Public', color: THEME.cream },
    a: { name: 'All', color: THEME.mute }
  };
  var DIMS = { S1: 'Age', S2: 'Gender', S5: 'Region', S6: 'Sport-following' };
  var AUDS = { compare: 'Trade / Public / All', trade: 'Trade', public: 'Public', all: 'All respondents' };

  // Categories per slide. A break-by chart can run to 40+ rows; past this the
  // labels stop being readable, so the chart continues on a further slide.
  var MAX_CATS = 16;

  /* ------------------------------------------------------------------ utils */

  function txt(el) { return el ? (el.textContent || '').replace(/\s+/g, ' ').trim() : ''; }

  // Label text without the "Indicative" flag chip.
  function labelOf(el) {
    if (!el) return '';
    var c = el.cloneNode(true);
    Array.prototype.forEach.call(c.querySelectorAll('.indicative'), function (x) { x.remove(); });
    return txt(c);
  }

  function loadLib() {
    if (loadLib._p) return loadLib._p;
    loadLib._p = new Promise(function (resolve, reject) {
      if (window.PptxGenJS) return resolve(window.PptxGenJS);
      var s = document.createElement('script');
      s.src = LIB;
      s.onload = function () {
        window.PptxGenJS ? resolve(window.PptxGenJS) : reject(new Error('PptxGenJS did not register'));
      };
      s.onerror = function () { reject(new Error('Could not load ' + LIB)); };
      document.head.appendChild(s);
    });
    return loadLib._p;
  }

  /* -------------------------------------------------------------- extraction */

  // The audience / break-by controls are built into each card by
  // buildCardAud(); read the selections back off them rather than reaching
  // into explore.html's `state`.
  function cardAudience(card) {
    var on = card.querySelector('.card-seg button.on');
    return on ? on.dataset.aud : 'compare';
  }
  function cardBreak(card) {
    var sel = card.querySelector('select.card-break');
    return sel ? sel.value : '';
  }

  // One rendered bar -> { label, key, value }. `key` is the t/p/a series class.
  function readRows(el, out, prefix) {
    if (el.classList.contains('pair')) {
      var lab = labelOf(el.querySelector('.pair-lab'));
      ['t', 'p', 'a'].forEach(function (k) {
        var row = el.querySelector('.pair-row.' + k);
        if (!row) return;
        var fill = row.querySelector('.bar-fill');
        if (!fill) return;
        out.push({ label: prefix + lab, key: k, value: parseFloat(fill.dataset.w) || 0 });
      });
      return true;
    }
    if (el.classList.contains('bar-row')) {
      var fill2 = el.querySelector('.bar-fill');
      if (!fill2) return true;
      var k2 = ['t', 'p', 'a'].filter(function (c) { return fill2.classList.contains(c); })[0] || 'a';
      out.push({
        label: prefix + labelOf(el.querySelector('.bar-lab span')),
        key: k2,
        value: parseFloat(fill2.dataset.w) || 0
      });
      return true;
    }
    return false;
  }

  // The two bespoke cards.
  function readStats(root, out) {
    Array.prototype.forEach.call(root.querySelectorAll('.stat'), function (st) {
      var name = txt(st.querySelector('.vlab')).split('·')[0].trim();
      var k = name === 'Trade' ? 't' : name === 'Public' ? 'p' : 'a';
      out.push({ label: name, key: k, value: parseFloat(txt(st.querySelector('.v'))) || 0 });
    });
  }
  function readDiverging(root, out) {
    Array.prototype.forEach.call(root.querySelectorAll('.div-side'), function (sd) {
      var name = txt(sd.querySelector('.div-name'));
      var k = name === 'Trade' ? 't' : name === 'Public' ? 'p' : 'a';
      out.push({ label: name, key: k, value: parseFloat(txt(sd.querySelector('.div-val')).replace('+', '')) || 0 });
    });
  }

  // card -> a chart model, or null if the card has nothing rendered.
  function extractCard(card) {
    var key = card.dataset.chart;
    var root = card.querySelector('#chart-' + key);
    if (!root) return null;

    var flat = [];
    if (key === 'likely') readStats(root, flat);
    else if (key === 'netexcite') readDiverging(root, flat);
    else {
      Array.prototype.forEach.call(root.children, function (child) {
        if (readRows(child, flat, '')) return;
        // A break-by option block: a heading div, then the bars for that option.
        // Separator is a middot, not an em dash: no em dashes in client output.
        if (!child.querySelector('.pair, .bar-row')) return;
        var head = labelOf(child.firstElementChild);
        Array.prototype.forEach.call(child.children, function (g) {
          readRows(g, flat, head ? head + ' · ' : '');
        });
      });
    }
    if (!flat.length) return null;

    // Pivot the flat rows into ordered categories x series.
    var cats = [], seen = {};
    flat.forEach(function (r) { if (!(r.label in seen)) { seen[r.label] = true; cats.push(r.label); } });

    var brk = cardBreak(card);
    var aud = cardAudience(card);
    // S1/S2/S5 are collected on the Public panel only; renderBreakBy draws
    // every bar with the 'a' class, so name the series for what it really is.
    var publicOnlyBreak = !!brk && brk !== 'S6';

    var order = ['t', 'p', 'a'], byKey = {};
    flat.forEach(function (r) {
      if (!byKey[r.key]) byKey[r.key] = {};
      byKey[r.key][r.label] = r.value;
    });
    var series = order.filter(function (k) { return byKey[k]; }).map(function (k) {
      return {
        name: publicOnlyBreak ? 'Public panel' : SERIES[k].name,
        color: SERIES[k].color,
        values: cats.map(function (c) { return byKey[k][c] == null ? 0 : byKey[k][c]; })
      };
    });

    var section = card.closest('section.block');
    var wave = card.closest('.wave-panel');
    var bits = [];
    if (wave) bits.push(wave.dataset.wave === 'w2' ? 'Wave 2' : 'Wave 1');
    if (section) bits.push(txt(section.querySelector('.kicker')).replace(/^\d+/, '').trim());
    bits.push('Audience: ' + (publicOnlyBreak ? 'Public panel' : (AUDS[aud] || aud)));
    if (brk) bits.push('Broken by ' + DIMS[brk]);

    return {
      key: key,
      title: txt(card.querySelector('.card-head h3')),
      subtitle: bits.filter(Boolean).join('  ·  '),
      footnote: txt(card.querySelector('[data-foot]')),
      categories: cats,
      series: series,
      axis: key === 'netexcite' ? 'net' : 'percent'
    };
  }

  /* ------------------------------------------------------------------- deck */

  function chunk(model) {
    var n = model.categories.length;
    if (n <= MAX_CATS) return [model];
    var parts = [], total = Math.ceil(n / MAX_CATS);
    for (var i = 0, p = 1; i < n; i += MAX_CATS, p++) {
      parts.push({
        key: model.key,
        title: model.title + ' (' + p + ' of ' + total + ')',
        subtitle: model.subtitle,
        footnote: model.footnote,
        axis: model.axis,
        categories: model.categories.slice(i, i + MAX_CATS),
        series: model.series.map(function (s) {
          return { name: s.name, color: s.color, values: s.values.slice(i, i + MAX_CATS) };
        })
      });
    }
    return parts;
  }

  function addChartSlide(pptx, m) {
    var slide = pptx.addSlide();
    slide.background = { color: THEME.bg };

    slide.addText(m.title, {
      x: 0.45, y: 0.28, w: 9.1, h: 0.52,
      fontFace: THEME.font, fontSize: 18, bold: true, color: THEME.text, valign: 'top'
    });
    slide.addShape(pptx.ShapeType.rect, { x: 0.45, y: 0.82, w: 0.9, h: 0.03, fill: { color: THEME.magenta } });
    slide.addText(m.subtitle, {
      x: 0.45, y: 0.88, w: 9.1, h: 0.26,
      fontFace: THEME.font, fontSize: 9.5, color: THEME.mute, valign: 'top'
    });

    // Horizontal bars put category 1 at the bottom, so reverse to match the
    // top-to-bottom reading order on the site.
    var cats = m.categories.slice().reverse();
    var data = m.series.map(function (s) {
      return { name: s.name, labels: cats, values: s.values.slice().reverse() };
    });

    var catFont = m.categories.length > 12 ? 8 : m.categories.length > 8 ? 9 : 10.5;
    var opts = {
      x: 0.45, y: 1.22, w: 9.1, h: 3.72,
      barDir: 'bar',
      barGapWidthPct: 45,
      chartColors: m.series.map(function (s) { return s.color; }),
      showValue: true,
      dataLabelColor: THEME.text,
      dataLabelFontFace: THEME.font,
      dataLabelFontSize: 8.5,
      dataLabelFormatCode: m.axis === 'net' ? '0;-0' : '0"%"',
      dataLabelPosition: 'outEnd',
      catAxisLabelColor: THEME.text,
      catAxisLabelFontFace: THEME.font,
      catAxisLabelFontSize: catFont,
      catAxisLineShow: false,
      valAxisLabelColor: THEME.mute,
      valAxisLabelFontFace: THEME.font,
      valAxisLabelFontSize: 9,
      valAxisLineShow: false,
      valGridLine: { color: THEME.grid, size: 0.75, style: 'solid' },
      catGridLine: { style: 'none' },
      showLegend: m.series.length > 1,
      legendPos: 't',
      legendColor: THEME.text,
      legendFontFace: THEME.font,
      legendFontSize: 10,
      plotArea: { fill: { color: THEME.bg } },
      chartArea: { fill: { color: THEME.bg }, border: { pt: 0, color: THEME.bg } }
    };
    if (m.axis === 'percent') { opts.valAxisMinVal = 0; opts.valAxisMaxVal = 100; }

    slide.addChart(pptx.ChartType.bar, data, opts);

    if (m.footnote) {
      slide.addText(m.footnote, {
        x: 0.45, y: 5.02, w: 7.7, h: 0.42,
        fontFace: THEME.font, fontSize: 7.5, color: THEME.mute, valign: 'top'
      });
    }
    slide.addText('Voice of the Trade', {
      x: 8.15, y: 5.02, w: 1.4, h: 0.24,
      fontFace: THEME.font, fontSize: 7.5, color: THEME.mute, align: 'right'
    });
  }

  function waveMeta(w) {
    try {
      var d = (w === 'w2') ? DATA_W2 : DATA_W1;
      return d && d.meta ? d.meta : null;
    } catch (e) { return null; }
  }

  function addCoverSlide(pptx, models, waves) {
    var slide = pptx.addSlide();
    slide.background = { color: THEME.bg };
    slide.addText('2026 World Cup Brand Tracker', {
      x: 0.7, y: 1.5, w: 8.6, h: 0.7,
      fontFace: THEME.font, fontSize: 32, bold: true, color: THEME.text
    });
    slide.addShape(pptx.ShapeType.rect, { x: 0.7, y: 2.3, w: 1.4, h: 0.05, fill: { color: THEME.magenta } });
    slide.addText(models.length + (models.length === 1 ? ' chart' : ' charts') + ' exported from Explore the data', {
      x: 0.7, y: 2.5, w: 8.6, h: 0.35, fontFace: THEME.font, fontSize: 13, color: THEME.cream
    });

    var lines = waves.map(function (w) {
      var m = waveMeta(w);
      var label = w === 'w2' ? 'Wave 2' : 'Wave 1';
      if (!m) return label;
      return label + ' · fielded ' + m.fieldwork + ' · trade n=' + m.n_trades + ', public n=' + m.n_public;
    });
    slide.addText(lines.join('\n'), {
      x: 0.7, y: 2.95, w: 8.6, h: 0.9, fontFace: THEME.font, fontSize: 11, color: THEME.mute, lineSpacingMultiple: 1.4
    });
    slide.addText('Voice of the Trade · Potentia Sports', {
      x: 0.7, y: 4.85, w: 8.6, h: 0.3, fontFace: THEME.font, fontSize: 10, color: THEME.mute
    });
  }

  function stamp() {
    var d = new Date(), p = function (n) { return (n < 10 ? '0' : '') + n; };
    return '' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate());
  }

  function buildDeck(cards) {
    var models = [], waves = [], skipped = [];
    cards.forEach(function (card) {
      var m = extractCard(card);
      if (!m) { skipped.push(card.dataset.chart); return; }
      models.push(m);
      var wp = card.closest('.wave-panel');
      var w = wp ? wp.dataset.wave : 'w1';
      if (waves.indexOf(w) === -1) waves.push(w);
    });
    if (!models.length) throw new Error('Nothing to export: none of the selected charts have rendered.');

    var pptx = new window.PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'Voice of the Trade';
    pptx.company = 'Potentia Sports';
    pptx.title = '2026 World Cup Brand Tracker';

    addCoverSlide(pptx, models, waves.sort());
    models.forEach(function (m) { chunk(m).forEach(function (part) { addChartSlide(pptx, part); }); });

    return pptx.writeFile({ fileName: 'VOTT-WC-Tracker-charts-' + stamp() + '.pptx' })
      .then(function () { return { count: models.length, skipped: skipped }; });
  }

  /* ---------------------------------------------------------------- controls */

  /* No floating bar. The shell sizes this page's iframe to its full content
     height and never scrolls it (index.html sets scrolling="no"), so the
     iframe's viewport IS the whole document: position:fixed pins to the bottom
     of a 14,000px box, miles below the visible area. The page's own sticky
     .jumpbox is inert for the same reason. So the export controls are repeated
     inline — once per card, next to the tick box, and once per wave toolbar —
     and whichever copy you can see is the one you use. */

  function selectedCards() {
    return Array.prototype.filter.call(
      document.querySelectorAll('.card[data-chart]'),
      function (c) { var b = c.querySelector('.card-pick input'); return b && b.checked; }
    );
  }

  function clearPicks() {
    Array.prototype.forEach.call(document.querySelectorAll('.card-pick input'), function (i) { i.checked = false; });
    refresh();
  }

  function refresh() {
    var n = selectedCards().length;
    Array.prototype.forEach.call(document.querySelectorAll('.pptx-pack'), function (p) {
      p.classList.toggle('on', n > 0);
      p.querySelector('.pptx-n').textContent = n;
      p.querySelector('.pptx-go').title = 'Export the ' + n + (n === 1 ? ' selected chart' : ' selected charts') + ' to PowerPoint';
    });
  }

  function setBusy(busy, label) {
    Array.prototype.forEach.call(document.querySelectorAll('.pptx-pack button'), function (b) { b.disabled = busy; });
    Array.prototype.forEach.call(document.querySelectorAll('.pptx-go .pptx-lab'), function (l) {
      l.textContent = label || 'Export';
    });
  }

  function runExport(pack) {
    var cards = selectedCards();
    if (!cards.length) return;
    Array.prototype.forEach.call(document.querySelectorAll('.pptx-msg'), function (m) { m.textContent = ''; });
    var msg = pack.querySelector('.pptx-msg');
    setBusy(true, window.PptxGenJS ? 'Building' : 'Preparing');
    loadLib()
      .then(function () { setBusy(true, 'Building'); return buildDeck(cards); })
      .then(function (res) {
        msg.textContent = 'Downloaded ' + res.count + (res.count === 1 ? ' chart' : ' charts') +
          (res.skipped.length ? ' (skipped ' + res.skipped.join(', ') + ')' : '');
      })
      .catch(function (err) {
        msg.textContent = 'Export failed: ' + err.message;
        console.error('[pptx-export]', err);
      })
      .then(function () { setBusy(false, 'Export'); });
  }

  function makePack(compact) {
    var pack = document.createElement('span');
    pack.className = 'pptx-pack' + (compact ? ' compact' : '');
    pack.innerHTML =
      '<button type="button" class="pptx-go"><span class="pptx-lab">Export</span> <span class="pptx-n">0</span></button>' +
      '<button type="button" class="pptx-clear" title="Clear the selection">Clear</button>' +
      '<span class="pptx-msg"></span>';
    pack.querySelector('.pptx-go').addEventListener('click', function () { runExport(pack); });
    pack.querySelector('.pptx-clear').addEventListener('click', clearPicks);
    return pack;
  }

  // Idempotent: re-run freely as cards and wave panels appear.
  function addControls() {
    Array.prototype.forEach.call(document.querySelectorAll('.card[data-chart]'), function (card) {
      var host = card.querySelector('[data-card-aud]');
      if (!host || card.querySelector('.card-pick')) return;
      var key = card.dataset.chart;
      var lab = document.createElement('label');
      lab.className = 'card-pick';
      lab.title = 'Include this chart in the PowerPoint export';
      lab.innerHTML = '<input type="checkbox" aria-label="Select ' + key + ' for PowerPoint export"><span>Select</span>';
      lab.querySelector('input').addEventListener('change', refresh);
      host.appendChild(lab);
      host.appendChild(makePack(true));
    });

    // .jumpbox is full-bleed; its content sits in an inner .wrap. Go in next to
    // "Reset all" so the controls stay in the content column.
    Array.prototype.forEach.call(document.querySelectorAll('.jumpbox'), function (jb) {
      if (jb.querySelector('.pptx-pack')) return;
      var reset = jb.querySelector('.cb-reset');
      var host = reset ? reset.parentNode : jb.querySelector('.wrap') || jb;
      var pack = makePack(false);
      if (reset) host.insertBefore(pack, reset.nextSibling);
      else host.appendChild(pack);
    });

    refresh();
  }

  function init() {
    addControls();

    // The second wave panel only renders on first tab open, and every bar
    // re-render mutates the DOM, so coalesce to one pass per frame rather than
    // rescanning 32 cards per mutation.
    var queued = false;
    new MutationObserver(function () {
      if (queued) return;
      queued = true;
      requestAnimationFrame(function () { queued = false; addControls(); });
    }).observe(document.body, { childList: true, subtree: true });

    // "Reset all" clears every other per-card choice; clear the ticks with it.
    Array.prototype.forEach.call(document.querySelectorAll('.cb-reset'), function (b) {
      b.addEventListener('click', clearPicks);
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
