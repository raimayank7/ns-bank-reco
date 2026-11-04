/* bank-reco.niveshsahayak.in: bank reconciliation on the shared data home.
   The bank statement (CSV or Excel from net banking, or lines typed by hand) is read in this browser only (NSBank in
   bankcore.js), matched with the bank side of the saved records (NSBooks), and the reconciliation statement shows why the
   two balances differ. From the unmatched lines the user can add bank charges as an expense, add money received from a
   party, and mark matched cheques as cleared. A reconciliation can be saved as a 'brs' record (books.js ignores it). */
(function () {
  'use strict';
  var $ = NS.$, $$ = NS.$$, esc = NS.esc, num = NS.num, r2 = NS.r2, money = NS.money, K = NSBank, t = NSB.t, key = NSB.key;
  var LIB = { xl: ['https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'sha384-vtjasyidUo0kW94K5MXDXntzOJpQgBKXmE7e2Ga4LG0skTTLeBi97eFAXsqewJjw'] };
  var COLS = ['date', 'desc', 'ref', 'dr', 'cr', 'amt', 'type', 'bal'];
  var cur = 'home', ST = null, keep = [], undo = [], saveId = '', tol = 10;
  /* ST = { name, rows, map, header, lines, from, to, bank } (rows kept so the columns can be changed) */

  function route() { cur = location.hash.slice(1) === 'saved' ? 'saved' : 'home'; show(true); }
  function show(focus) {
    NSB.navMark(cur);
    return NSB.load().then(function () {
      var pane = $('[data-pane]');
      NSB.modeLine();
      if (!NSB.biz()) { pane.innerHTML = NSB.noBiz(); return; }
      pane.innerHTML = NSB.bizBar() + (cur === 'saved' ? savedHtml() : homeHtml());
      wire(pane);
      if (focus) NSB.focusH2(pane);
    });
  }
  function D() { return NSB.mine(); }
  function byId(id) { return NSB.data().doc.filter(function (d) { return d.id === id; })[0]; }

  /* ---------- the work ---------- */
  function lines() { return (ST && ST.lines || []).filter(function (l) { return (!ST.from || l.date >= ST.from) && (!ST.to || l.date <= ST.to); }); }
  function calc() {
    var L = lines(), book = K.bookLines(D(), NSB.bizName(), ST.from, ST.to), M = K.match(L, book, { days: tol, keep: keep, no: undo });
    var R = K.brs(K.bookBalance(D(), NSB.bizName(), ST.to), M, ST.bank);
    return { L: L, book: book, M: M, R: R };
  }
  function lastBal(L) { for (var i = L.length - 1; i >= 0; i--) if (L[i].bal != null) return L[i].bal; return null; }

  /* ---------- screen ---------- */
  function homeHtml() {
    var head = '<div class="pane-hd"><h2>' + esc(t('br.title')) + '</h2><div class="quick">' +
      '<label class="btn primary" for="br-file">' + esc(t(ST ? 'br.other' : 'br.open')) + '</label><input type="file" id="br-file" class="sr" accept=".csv,.txt,.xls,.xlsx,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">' +
      '<button type="button" class="btn" data-type>+ ' + esc(t('br.typeline')) + '</button></div></div>';
    if (!ST || !ST.lines.length) return head + '<div class="card"><h3>' + esc(t('br.start')) + '</h3><p>' + esc(t('br.starth')) + '</p>' +
      (ST && !ST.lines.length ? '<p class="sb-note bad">' + esc(t('br.nolines')) + '</p>' + colsHtml() : '') +
      '<ol class="br-steps"><li>' + esc(t('br.s1')) + '</li><li>' + esc(t('br.s2')) + '</li><li>' + esc(t('br.s3')) + '</li></ol></div>';
    var C = calc(), M = C.M, R = C.R, nPairs = Object.keys(M.pairs).length;
    var pend = Object.keys(M.pairs).map(function (i) { return C.book.filter(function (b) { return b.id === M.pairs[i]; })[0]; }).filter(function (b) { return b && b.pending; });
    var info = '<p class="sb-note">' + esc(t('br.from', { f: ST.name || t('br.typed'), n: ST.lines.length, a: NS.fmtDate(ST.lines[0].date), b: NS.fmtDate(ST.lines[ST.lines.length - 1].date) })) + '</p>' +
      (ST.rows ? '<details class="br-cols"><summary>' + esc(t('br.cols')) + '</summary>' + colsHtml() + '</details>' : '');
    var setup = '<div class="sb-period"><div class="f"><label for="br-from">' + esc(t('sb.from')) + '</label><input type="date" id="br-from" value="' + esc(ST.from) + '"></div>' +
      '<div class="f"><label for="br-to">' + esc(t('sb.to')) + '</label><input type="date" id="br-to" value="' + esc(ST.to) + '"></div>' +
      '<div class="f"><label for="br-bal">' + esc(t('br.bankbal', { d: NS.fmtDate(ST.to) })) + '</label><input id="br-bal" inputmode="decimal" value="' + esc(ST.bank == null ? '' : ST.bank) + '"></div>' +
      '<div class="f"><label for="br-tol">' + esc(t('br.tol')) + '</label><input id="br-tol" type="number" min="0" max="60" value="' + tol + '"></div></div>';
    var diffOk = R.diff != null && Math.abs(R.diff) < 0.005;
    var tiles = '<div class="tiles"><div class="tile"><span>' + esc(t('br.matched')) + '</span><b>' + nPairs + '</b></div>' +
      '<a class="tile' + (M.onlyBank.length ? ' warn' : '') + '" href="#br-bank"><span>' + esc(t('br.onlybank')) + '</span><b>' + M.onlyBank.length + '</b></a>' +
      '<a class="tile' + (M.onlyBook.length ? ' warn' : '') + '" href="#br-book"><span>' + esc(t('br.onlybook')) + '</span><b>' + M.onlyBook.length + '</b></a>' +
      '<div class="tile' + (R.diff == null ? '' : diffOk ? ' ok' : ' neg') + '"><span>' + esc(t('br.diff')) + '</span><b>' + (R.diff == null ? '—' : money(R.diff)) + '</b></div></div>';
    var brs = '<h3>' + esc(t('br.brs', { d: NS.fmtDate(ST.to) })) + '</h3><div class="tbl-wrap"><table class="tbl br-brs"><tbody>' +
      row(t('br.r_book'), R.book, 1) + row('+ ' + t('br.r_issued'), R.issued) + row('− ' + t('br.r_dep'), R.deposited) + row('+ ' + t('br.r_cr'), R.credits) + row('− ' + t('br.r_dr'), R.debits) +
      row(t('br.r_should'), R.shouldBe, 1) + (R.bank == null ? '' : row(t('br.r_bank'), R.bank, 1) + row(t('br.diff'), R.diff, 1, diffOk ? 'ok' : 'bad')) + '</tbody></table></div>' +
      '<p class="sb-note">' + esc(R.bank == null ? t('br.needbal') : diffOk ? t('br.ok') : t('br.notok')) + '</p>';
    var bank = '<h3 id="br-bank">' + esc(t('br.onlybank')) + ' (' + M.onlyBank.length + ')</h3><p class="hint">' + esc(t('br.onlybankh')) + '</p>' +
      (M.onlyBank.length ? '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>' + esc(t('br.date')) + '</th><th class="l">' + esc(t('br.details')) + '</th><th class="n">' + esc(t('br.wd')) + '</th><th class="n">' + esc(t('br.dep')) + '</th><th class="l">' + esc(t('br.do')) + '</th></tr></thead><tbody>' +
        M.onlyBank.map(function (l) {
          var opts = C.book.filter(function (b) { return M.onlyBook.indexOf(b) >= 0 && (l.cr ? b.inn : b.out); });
          return '<tr><td>' + NS.fmtDate(l.date) + '</td><td class="l">' + esc(l.desc || '—') + (l.ref ? ' <span class="mono soft">' + esc(l.ref) + '</span>' : '') + '</td><td class="n">' + (l.dr ? money(l.dr, false) : '') + '</td><td class="n">' + (l.cr ? money(l.cr, false) : '') + '</td>' +
            '<td class="l br-do">' + (l.dr ? '<button type="button" class="btn sm" data-addexp="' + l.i + '">' + esc(t('br.addexp')) + '</button>' : '<button type="button" class="btn sm" data-addin="' + l.i + '">' + esc(t('br.addin')) + '</button>') +
            (opts.length ? '<label class="sr" for="br-m' + l.i + '">' + esc(t('br.matchwith')) + '</label><select id="br-m' + l.i + '" data-pick="' + l.i + '"><option value="">' + esc(t('br.matchwith')) + '</option>' +
              opts.map(function (b) { return '<option value="' + esc(b.id) + '">' + esc(NS.fmtDate(b.date) + ' · ' + bookName(b) + ' · ' + money(l.cr ? b.inn : b.out, false)) + '</option>'; }).join('') + '</select>' : '') + '</td></tr>';
        }).join('') + '</tbody></table></div>' : '<p class="sb-empty">' + esc(t('br.none')) + '</p>');
    var book = '<h3 id="br-book">' + esc(t('br.onlybook')) + ' (' + M.onlyBook.length + ')</h3><p class="hint">' + esc(t('br.onlybookh')) + '</p>' +
      (M.onlyBook.length ? '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>' + esc(t('br.date')) + '</th><th class="l">' + esc(t('br.record')) + '</th><th class="n">' + esc(t('br.in')) + '</th><th class="n">' + esc(t('br.out')) + '</th><th class="l">' + esc(t('br.state')) + '</th></tr></thead><tbody>' +
        M.onlyBook.map(function (b) {
          return '<tr><td>' + NS.fmtDate(b.date) + '</td><td class="l">' + esc(bookName(b)) + (b.ref ? ' <span class="mono soft">' + esc(b.ref) + '</span>' : '') + '</td><td class="n">' + (b.inn ? money(b.inn, false) : '') + '</td><td class="n">' + (b.out ? money(b.out, false) : '') + '</td>' +
            '<td class="l">' + esc(b.pending === 'dep' ? t('br.p_dep') : b.pending === 'iss' ? t('br.p_iss') : b.inn ? t('br.p_in') : t('br.p_out')) + '</td></tr>';
        }).join('') + '</tbody></table></div>' : '<p class="sb-empty">' + esc(t('br.none')) + '</p>');
    var paired = '<details class="br-pairs"><summary>' + esc(t('br.matched')) + ' (' + nPairs + ')</summary>' + (nPairs ? '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>' + esc(t('br.date')) + '</th><th class="l">' + esc(t('br.details')) + '</th><th class="n">' + esc(t('br.amt')) + '</th><th class="l">' + esc(t('br.record')) + '</th><th></th></tr></thead><tbody>' +
      C.L.filter(function (l) { return M.pairs[l.i]; }).map(function (l) {
        var b = C.book.filter(function (x) { return x.id === M.pairs[l.i]; })[0];
        return '<tr><td>' + NS.fmtDate(l.date) + '</td><td class="l">' + esc(l.desc || '—') + '</td><td class="n">' + money(l.cr || l.dr, false) + (l.cr ? ' ' + esc(t('br.crs')) : ' ' + esc(t('br.drs'))) + '</td><td class="l">' + esc(NS.fmtDate(b.date) + ' · ' + bookName(b)) + '</td>' +
          '<td><button type="button" class="btn sm" data-unpair="' + l.i + '|' + esc(b.id) + '">' + esc(t('br.unmatch')) + '</button></td></tr>';
      }).join('') + '</tbody></table></div>' : '') + '</details>';
    var clear = pend.length ? '<div class="callout"><b>' + esc(t('br.clearq', { n: pend.length })) + '</b> ' + esc(t('br.clearh')) + ' <button type="button" class="btn sm primary" data-clearchq>' + esc(t('br.clear')) + '</button></div>' : '';
    return head + info + setup + tiles + clear + brs + bank + book + paired +
      '<div class="btns"><button type="button" class="btn primary" data-save>' + esc(t('sb.save')) + '</button><button type="button" class="btn" data-print>' + esc(t('sb.print')) + '</button><button type="button" class="btn" data-pdf>' + esc(t('sb.pdf')) + '</button><button type="button" class="btn" data-csv>' + esc(t('sb.csv')) + '</button>' +
      '<button type="button" class="btn danger" data-reset>' + esc(t('br.reset')) + '</button></div>' +
      '<div class="sb-print">' + sheet(C) + '</div>';
  }
  function row(label, v, strong, cls) { return '<tr' + (cls ? ' class="' + cls + '"' : '') + '><th scope="row" class="l">' + (strong ? '<b>' + esc(label) + '</b>' : esc(label)) + '</th><td class="n">' + (strong ? '<b>' + money(v, false) + '</b>' : money(v, false)) + '</td></tr>'; }
  function bookName(b, en) {
    var d = b.d, who = d.buyer && d.buyer.name, what = d.type === 'contra' ? (en ? (d.to === 'bank' ? 'Cash deposited' : 'Cash withdrawn') : t(d.to === 'bank' ? 'br.c_dep' : 'br.c_wd')) : (en ? NSB.enName(d) : NSB.docName(d));
    return what + (d.no ? ' ' + d.no : '') + (who ? ' · ' + who : d.cat ? ' · ' + d.cat : '');
  }
  function colsHtml() {
    var M = ST && ST.map || {}, n = ST && ST.rows ? Math.max.apply(null, ST.rows.slice(0, 50).map(function (r) { return r.length; })) : 0, hdr = ST && ST.rows && ST.header >= 0 ? ST.rows[ST.header] || [] : [];
    var opt = function (k) { return '<option value="">—</option>' + Array.apply(null, Array(n)).map(function (_, i) { return '<option value="' + i + '"' + (M[k] === i ? ' selected' : '') + '>' + esc((i + 1) + (hdr[i] ? ': ' + String(hdr[i]).slice(0, 24) : '')) + '</option>'; }).join(''); };
    return '<p class="hint">' + esc(t('br.colsh')) + '</p><div class="fg br-map">' + COLS.map(function (k) { return '<div class="f"><label for="br-c-' + k + '">' + esc(t('br.c_' + k)) + '</label><select id="br-c-' + k + '" data-col="' + k + '">' + opt(k) + '</select></div>'; }).join('') + '</div>';
  }

  /* English print sheet */
  function sheet(C) {
    var b = NSB.biz() || {}, R = C.R, M = C.M;
    function tr(l, v, s) { return '<tr' + (s ? ' class="b"' : '') + '><td>' + esc(l) + '</td><td class="n">' + money(v, false) + '</td></tr>'; }
    return '<div class="stsheet" data-sheet><div class="st-hd"><div><h3>' + esc(b.name || '') + '</h3><div class="soft">' + esc([b.addr, b.city, b.gstin && 'GSTIN ' + b.gstin].filter(Boolean).join(', ')) + '</div></div>' +
      '<div class="r"><b>Bank reconciliation statement</b><div class="soft">as on ' + NS.fmtDate(ST.to) + (b.bank ? ' · ' + esc(b.bank) : '') + '</div></div></div>' +
      '<table class="tbl"><tbody>' + tr('Balance as per books (bank account)', R.book, 1) + tr('Add: cheques issued but not yet presented', R.issued) + tr('Less: cheques and amounts deposited but not yet credited', R.deposited) +
      tr('Add: credits in the bank not yet in the books', R.credits) + tr('Less: debits in the bank not yet in the books', R.debits) + tr('Balance as per bank statement (worked out)', R.shouldBe, 1) +
      (R.bank == null ? '' : tr('Balance as per bank statement', R.bank, 1) + tr('Difference', R.diff, 1)) + '</tbody></table>' +
      (M.onlyBook.length ? '<h4>In the books, not in the bank statement</h4><table class="tbl"><thead><tr><th>Date</th><th>Record</th><th class="n">Received</th><th class="n">Paid</th></tr></thead><tbody>' +
        M.onlyBook.map(function (x) { return '<tr><td>' + NS.fmtDate(x.date) + '</td><td>' + esc(bookName(x, 1)) + (x.ref ? ' (' + esc(x.ref) + ')' : '') + '</td><td class="n">' + (x.inn ? money(x.inn, false) : '') + '</td><td class="n">' + (x.out ? money(x.out, false) : '') + '</td></tr>'; }).join('') + '</tbody></table>' : '') +
      (M.onlyBank.length ? '<h4>In the bank statement, not in the books</h4><table class="tbl"><thead><tr><th>Date</th><th>Details</th><th class="n">Withdrawal</th><th class="n">Deposit</th></tr></thead><tbody>' +
        M.onlyBank.map(function (l) { return '<tr><td>' + NS.fmtDate(l.date) + '</td><td>' + esc(l.desc) + (l.ref ? ' (' + esc(l.ref) + ')' : '') + '</td><td class="n">' + (l.dr ? money(l.dr, false) : '') + '</td><td class="n">' + (l.cr ? money(l.cr, false) : '') + '</td></tr>'; }).join('') + '</tbody></table>' : '') +
      '<div class="st-ft">Made with Nivesh Sahayak Bank Reconciliation (bank-reco.niveshsahayak.in) on ' + NS.fmtDate(NS.today()) + '</div></div>';
  }

  /* ---------- saved reconciliations ---------- */
  function savedHtml() {
    var list = D().filter(function (d) { return d.type === 'brs'; }).sort(function (a, b) { return (b.to || '').localeCompare(a.to || ''); });
    return '<div class="pane-hd"><h2>' + esc(t('br.saved')) + '</h2></div>' + (list.length ? '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>' + esc(t('br.asof')) + '</th><th class="l">' + esc(t('br.file')) + '</th><th class="n">' + esc(t('br.r_bank')) + '</th><th class="n">' + esc(t('br.diff')) + '</th><th></th></tr></thead><tbody>' +
      list.map(function (d) {
        return '<tr><td>' + NS.fmtDate(d.to) + '</td><td class="l">' + esc(d.file || t('br.typed')) + ' · ' + esc(t('br.nl', { n: (d.lines || []).length })) + '</td><td class="n">' + (d.bank == null ? '—' : money(d.bank, false)) + '</td><td class="n">' + (d.diff == null ? '—' : money(d.diff, false)) + '</td>' +
          '<td><button type="button" class="btn sm" data-open="' + esc(d.id) + '">' + esc(t('br.reopen')) + '</button> <button type="button" class="btn sm danger" data-del="' + esc(d.id) + '">' + esc(t('sb.del')) + '</button></td></tr>';
      }).join('') + '</tbody></table></div>' : '<div class="card"><p>' + esc(t('br.nosaved')) + '</p></div>');
  }

  /* ---------- reading a statement ---------- */
  function readFile(file) {
    var xl = /\.xlsx?$/i.test(file.name);
    var get = xl ? NS.loadJs(LIB.xl[0], LIB.xl[1]).then(function () { return file.arrayBuffer(); }).then(function (buf) {
      var wb = window.XLSX.read(buf, { type: 'array', cellDates: false }), ws = wb.Sheets[wb.SheetNames[0]];
      return window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    }) : file.text().then(K.csvRows);
    return get.then(function (rows) { setRows(rows, file.name); }).catch(function () { NS.toast(t('br.readfail'), 'bad'); });
  }
  function setRows(rows, name, map) {
    rows = (rows || []).filter(function (r) { return r && r.some(function (c) { return String(c == null ? '' : c).trim(); }); });
    var S = K.statement(rows, map, map ? (ST && ST.header) : undefined);
    ST = { name: name, rows: rows, map: S.map || {}, header: S.header == null ? -1 : S.header, lines: S.lines, from: '', to: '', bank: null };
    if (S.lines.length) {
      var ds = S.lines.map(function (l) { return l.date; }).sort();
      ST.lines.sort(function (a, b) { return a.date.localeCompare(b.date) || a.i - b.i; });
      ST.lines.forEach(function (l, i) { l.i = i; });
      ST.from = ds[0]; ST.to = ds[ds.length - 1]; ST.bank = lastBal(ST.lines);
    }
    keep = []; undo = []; saveId = '';
    show();
    if (S.lines.length) NS.toast(t('br.read', { n: S.lines.length }));
  }
  function typeForm() {
    var f = NSB.dialog(t('br.typeline'), '<p class="hint">' + esc(t('br.typeh')) + '</p><div class="fg"><div class="f"><label for="f-date">' + esc(t('br.date')) + '</label><input type="date" id="f-date" required value="' + esc(NS.today()) + '"></div>' +
      '<div class="f wide"><label for="f-desc">' + esc(t('br.details')) + '</label><input id="f-desc" dir="auto"></div><div class="f"><label for="f-ref">' + esc(t('br.c_ref')) + '</label><input id="f-ref" class="mono"></div>' +
      '<div class="f"><label for="f-dr">' + esc(t('br.wd')) + '</label><input id="f-dr" inputmode="decimal"></div><div class="f"><label for="f-cr">' + esc(t('br.dep')) + '</label><input id="f-cr" inputmode="decimal"></div></div>' +
      '<div class="btns"><button class="btn primary">' + esc(t('br.addline')) + '</button><button type="button" class="btn" onclick="this.closest(\'dialog\').close()">' + esc(t('sb.cancel')) + '</button></div>');
    f.onsubmit = function (e) {
      e.preventDefault();
      var dr = r2(num($('#f-dr', f).value)), cr = r2(num($('#f-cr', f).value)), d = $('#f-date', f).value;
      if (!d || !dr === !cr) { NS.toast(t('br.oneamt'), 'bad'); return; }
      if (!ST || ST.rows) { ST = { name: '', rows: null, map: {}, header: -1, lines: [], from: d, to: d, bank: null }; keep = []; undo = []; saveId = ''; }
      ST.lines.push({ i: ST.lines.length, date: d, desc: $('#f-desc', f).value.trim(), ref: $('#f-ref', f).value.trim(), dr: dr, cr: cr, bal: null });
      ST.lines.sort(function (a, b) { return a.date.localeCompare(b.date); }); ST.lines.forEach(function (l, i) { l.i = i; });
      ST.from = ST.lines[0].date; ST.to = ST.lines[ST.lines.length - 1].date;
      keep = []; undo = [];
      $('#f-desc', f).value = ''; $('#f-ref', f).value = ''; $('#f-dr', f).value = ''; $('#f-cr', f).value = ''; $('#f-desc', f).focus();
      NS.toast(t('br.added')); show();
    };
  }

  /* ---------- putting bank-only lines into the books ---------- */
  function nextNo(type, pre, date) {
    var fy = NS.fy(date), short = fy.slice(2, 4) + '-' + fy.slice(5), max = 0;
    D().forEach(function (d) { if (d.type !== type || NS.fy(d.date) !== fy) return; var m = String(d.no || '').match(/(\d+)\s*$/); if (m) max = Math.max(max, +m[1]); });
    return pre + '/' + short + '/' + String(max + 1).padStart(3, '0');
  }
  function expForm(l) {
    var guess = /int(erest)?\b/i.test(l.desc) ? t('br.e_int') : t('br.e_chg');
    var f = NSB.dialog(t('br.addexp'), '<p class="hint">' + esc(NS.fmtDate(l.date) + ' · ' + (l.desc || '') + ' · ' + money(l.dr)) + '</p><div class="fg"><div class="f wide"><label for="f-cat">' + esc(t('br.e_cat')) + '</label><input id="f-cat" dir="auto" required value="' + esc(guess) + '" list="br-cats"><datalist id="br-cats"><option value="' + esc(t('br.e_chg')) + '"><option value="' + esc(t('br.e_int')) + '"><option value="' + esc(t('br.e_emi')) + '"></datalist></div></div>' +
      '<div class="btns"><button class="btn primary">' + esc(t('sb.save')) + '</button><button type="button" class="btn" onclick="this.closest(\'dialog\').close()">' + esc(t('sb.cancel')) + '</button></div>');
    f.onsubmit = function (e) {
      e.preventDefault();
      NSStore.put('doc', { type: 'exp', app: 1, date: l.date, seller: NSB.sellerCopy(), cat: $('#f-cat', f).value.trim() || t('br.e_chg'), amount: l.dr, mode: 'bank', buyer: { name: '' }, no: l.ref || '', gst: '', itc: false, inter: false, notes: l.desc, total: l.dr }).then(done);
    };
  }
  function inForm(l) {
    var names = NSB.data().party.map(function (p) { return p.name; }).filter(Boolean);
    var f = NSB.dialog(t('br.addin'), '<p class="hint">' + esc(NS.fmtDate(l.date) + ' · ' + (l.desc || '') + ' · ' + money(l.cr)) + '</p><p class="hint">' + esc(t('br.addinh')) + '</p><div class="fg"><div class="f wide"><label for="f-party">' + esc(t('br.party')) + '</label><input id="f-party" dir="auto" required list="br-parties"><datalist id="br-parties">' +
      names.map(function (n) { return '<option value="' + esc(n) + '">'; }).join('') + '</datalist></div></div>' +
      '<div class="btns"><button class="btn primary">' + esc(t('sb.save')) + '</button><button type="button" class="btn" onclick="this.closest(\'dialog\').close()">' + esc(t('sb.cancel')) + '</button></div>');
    f.onsubmit = function (e) {
      e.preventDefault();
      var party = $('#f-party', f).value.trim(); if (!party) return;
      NSStore.put('doc', { type: 'receipt', no: nextNo('receipt', 'RCT', l.date), paper: 'a4', design: 'classic', copies: 1, copyLabel: '', gstAdv: false, rate: '18', pos: '', sac: '', rv: '', rvDate: '', refDate: '', bank: '', balance: '', status: '',
        app: 1, kind: 'pay', date: l.date, seller: NSB.sellerCopy(), amount: l.cr, tds: '', mode: 'bank', ref: l.ref || K.refs(l.desc)[0] || '', against: '', purpose: l.desc, alloc: [],
        buyer: { name: party, addr: '', phone: '', gstin: '', pan: '' }, total: l.cr }).then(done);
    };
  }
  function clearCheques() {
    var C = calc(), jobs = [];
    Object.keys(C.M.pairs).forEach(function (i) {
      var b = C.book.filter(function (x) { return x.id === C.M.pairs[i]; })[0], l = C.L.filter(function (x) { return String(x.i) === String(i); })[0];
      if (!b || !b.pending || !l) return;
      var d = byId(b.id); if (!d) return;
      d.chq = Object.assign({}, d.chq, { st: 'ok', on: l.date });
      jobs.push(NSStore.put('doc', d));
    });
    Promise.all(jobs).then(function () { NS.toast(t('br.cleared', { n: jobs.length })); show(); });
  }
  function done() { NSB.closeDialog(); NS.toast(t('sb.saved')); if (NS.checkReminder) NS.checkReminder(); show(); }

  function save() {
    var C = calc(), R = C.R, rec = saveId && byId(saveId) || { type: 'brs', app: 1 };
    Object.assign(rec, { seller: NSB.sellerCopy(), date: ST.to, from: ST.from, to: ST.to, file: ST.name || '', bank: R.bank, book: R.book, diff: R.diff, tol: tol,
      lines: ST.lines.map(function (l) { return [l.date, l.desc, l.ref, l.dr, l.cr, l.bal]; }), keep: keep, undo: undo,
      pairs: Object.keys(C.M.pairs).map(function (i) { return [+i, C.M.pairs[i]]; }) });
    NSStore.put('doc', rec).then(function (s) { saveId = s && s.id || rec.id; NS.toast(t('sb.saved')); show(); });
  }
  function reopen(d) {
    ST = { name: d.file || '', rows: null, map: {}, header: -1, from: d.from, to: d.to, bank: d.bank,
      lines: (d.lines || []).map(function (a, i) { return { i: i, date: a[0], desc: a[1], ref: a[2], dr: a[3], cr: a[4], bal: a[5] }; }) };
    keep = d.keep || []; undo = d.undo || []; tol = d.tol == null ? 10 : d.tol; saveId = d.id;
    location.hash = ''; route();
  }

  function wire(pane) {
    NSB.wireBiz(pane, show);
    var fi = $('#br-file', pane);
    if (fi) fi.addEventListener('change', function () { if (fi.files[0]) readFile(fi.files[0]); fi.value = ''; });
    var ty = $('[data-type]', pane); if (ty) ty.addEventListener('click', typeForm);
    function val(id, fn) { var x = $(id, pane); if (x) x.addEventListener('change', function () { fn(x.value); show(); }); }
    val('#br-from', function (v) { ST.from = v; }); val('#br-to', function (v) { ST.to = v; });
    val('#br-bal', function (v) { ST.bank = v.trim() === '' ? null : r2(num(v)); });
    val('#br-tol', function (v) { tol = Math.max(0, Math.min(60, parseInt(v, 10) || 0)); });
    $$('[data-col]', pane).forEach(function (s) { s.addEventListener('change', function () {
      var m = Object.assign({}, ST.map); m[s.getAttribute('data-col')] = s.value === '' ? null : +s.value; setRows(ST.rows, ST.name, m);
    }); });
    $$('[data-addexp]', pane).forEach(function (b) { b.addEventListener('click', function () { expForm(ST.lines[+b.getAttribute('data-addexp')]); }); });
    $$('[data-addin]', pane).forEach(function (b) { b.addEventListener('click', function () { inForm(ST.lines[+b.getAttribute('data-addin')]); }); });
    $$('[data-pick]', pane).forEach(function (s) { s.addEventListener('change', function () {
      if (!s.value) return; var i = +s.getAttribute('data-pick');
      undo = undo.filter(function (p) { return !(p[0] === i && p[1] === s.value); }); keep.push([i, s.value]); show();
    }); });
    $$('[data-unpair]', pane).forEach(function (b) { b.addEventListener('click', function () {
      var p = b.getAttribute('data-unpair').split('|'), i = +p[0];
      keep = keep.filter(function (k) { return !(k[0] === i && k[1] === p[1]); }); undo.push([i, p[1]]); show();
    }); });
    var cc = $('[data-clearchq]', pane); if (cc) cc.addEventListener('click', clearCheques);
    var sv = $('[data-save]', pane); if (sv) sv.addEventListener('click', save);
    var rs = $('[data-reset]', pane); if (rs) rs.addEventListener('click', function () { if (confirm(t('br.resetc'))) { ST = null; keep = []; undo = []; saveId = ''; show(); } });
    $$('[data-open]', pane).forEach(function (b) { b.addEventListener('click', function () { var d = byId(b.getAttribute('data-open')); if (d) reopen(d); }); });
    $$('[data-del]', pane).forEach(function (b) { b.addEventListener('click', function () {
      var d = byId(b.getAttribute('data-del'));
      if (d && confirm(t('sb.cdel', { x: t('br.one') + ' ' + NS.fmtDate(d.to) }))) NSStore.del('doc', d.id).then(function () { if (saveId === d.id) saveId = ''; NS.toast(t('sb.deleted')); show(); });
    }); });
    var pr = $('[data-print]', pane), pdf = $('[data-pdf]', pane), cv = $('[data-csv]', pane);
    if (pr) pr.addEventListener('click', NSB.printSheet);
    if (pdf) pdf.addEventListener('click', function () { NSB.shareSheet($('[data-sheet]', pane), NSB.bizName() + ' bank reconciliation ' + ST.to); });
    if (cv) cv.addEventListener('click', function () {
      var C = calc();
      NSB.csv('bank-reconciliation-' + ST.to + '.csv', [['Where', 'Date', 'Details', 'Reference', 'Money in', 'Money out']]
        .concat(C.M.onlyBank.map(function (l) { return ['In the bank, not in the books', l.date, l.desc, l.ref, l.cr || '', l.dr || '']; }))
        .concat(C.M.onlyBook.map(function (b) { return ['In the books, not in the bank', b.date, bookName(b, 1), b.ref, b.inn || '', b.out || '']; })));
    });
  }

  function init() {
    if (!$('#bankreco')) return;
    window.addEventListener('hashchange', route);
    document.addEventListener('ns-lang', function () { show(); });
    NSStore.ready().then(route).then(function () { NSStore.persist().catch(function () {}); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
