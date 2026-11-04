/* NSBank: bank reconciliation for bank-reco.niveshsahayak.in. Pure functions, no screen.
   statement(rows): a bank statement read from CSV or Excel (rows = arrays of cells) -> { map, lines: [{ i, date, desc, ref, dr, cr, bal }] }
     dr = withdrawal (money out of the bank), cr = deposit (money in); dates come back as yyyy-mm-dd, day first as Indian banks print them.
   bookLines(docs, biz, from, to): the business's bank side from the saved records (NSBooks): every line posted to 'bank' in the period
     (a cleared cheque on the day it cleared) plus cheques deposited or issued and not cleared yet, whatever their date.
   match(lines, book, opts): pairs a statement line with a book line of the same amount and direction; same cheque or reference
     number first, then the nearest date within opts.days (default 10). opts.keep = [[lineIndex, bookId]] pairs made by hand, opts.no = pairs the user undid.
   brs(...): the reconciliation statement: balance as per books, items not in the bank yet, items not in the books yet,
     the balance the bank should show, and the difference from the balance it does show. */
(function (G) {
  'use strict';
  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v == null ? '' : v).trim(), neg = /^\(.*\)$/.test(s) || /^-/.test(s) || /\bdr\.?$/i.test(s);
    s = s.replace(/[^0-9.]/g, '');
    var n = parseFloat(s);
    return isFinite(n) ? (neg && !/\bcr\.?$/i.test(String(v)) ? -n : n) : 0;
  }
  function r2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
  function key(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
  var MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12 };
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function ok(y, m, d) {
    if (y < 100) y += 2000;
    if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 1990 && y <= 2100)) return '';
    var t = new Date(Date.UTC(y, m - 1, d));
    return t.getUTCMonth() === m - 1 ? y + '-' + pad(m) + '-' + pad(d) : '';
  }
  /* '05/04/2026', '05-04-26', '05.04.2026', '5 Apr 2026', '05-Apr-26', '2026-04-05', an Excel day number, or a Date */
  function date(v) {
    if (v instanceof Date && !isNaN(v)) return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
    if (typeof v === 'number' && v > 20000 && v < 80000) { var x = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5); return x.toISOString().slice(0, 10); }
    var s = String(v == null ? '' : v).trim(), m;
    if ((m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/))) return ok(+m[1], +m[2], +m[3]);
    if ((m = s.match(/^(\d{1,2})[-\/. ](\d{1,2})[-\/. ](\d{2,4})\b/))) return ok(+m[3], +m[2], +m[1]);
    if ((m = s.match(/^(\d{1,2})[-\/. ]?([A-Za-z]{3,4})[a-z]*[-\/. ,]*(\d{2,4})\b/))) return MON[m[2].toLowerCase()] ? ok(+m[3], MON[m[2].toLowerCase()], +m[1]) : '';
    if ((m = s.match(/^([A-Za-z]{3,4})[a-z]*[ -](\d{1,2}),?[ -](\d{2,4})\b/))) return MON[m[1].toLowerCase()] ? ok(+m[3], MON[m[1].toLowerCase()], +m[2]) : '';
    return '';
  }
  function days(a, b) { return Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 864e5); }

  /* which column is which, from the header row of the statement */
  var HEAD = [
    ['type', /^(dr\s*\/\s*cr|cr\s*\/\s*dr|type|txn type|dr cr|d\/c|c\/d)$/i],
    ['date', /^(txn|tran|transaction|trans|posting|post|value|book)?\.?\s*(date|dt)\b/i],
    ['desc', /narration|description|particulars|details|remarks|transaction details/i],
    ['ref', /(chq|cheque|check)|ref|utr|instrument/i],
    ['dr', /withdraw|debit|^dr\.?\b|paid out|amount out|money out/i],
    ['cr', /deposit|credit|^cr\.?\b|paid in|amount in|money in/i],
    ['amt', /^(txn\s*|transaction\s*)?amount/i],
    ['bal', /balance/i]
  ];
  function findHeader(rows) {
    for (var r = 0; r < Math.min(rows.length, 40); r++) {
      var cells = (rows[r] || []).map(function (c) { return String(c == null ? '' : c).trim(); }), map = {};
      cells.forEach(function (c, i) {
        if (!c) return;
        for (var h = 0; h < HEAD.length; h++) {
          var k = HEAD[h][0];
          if (HEAD[h][1].test(c)) {
            if (k === 'date' && map.date != null && /value/i.test(c)) break;   /* keep the transaction date, not the value date */
            if (map[k] == null) map[k] = i;
            break;
          }
        }
      });
      if (map.date != null && (map.dr != null || map.cr != null || map.amt != null)) return { row: r, map: map };
    }
    return null;
  }
  /* no header found: guess from the data (first date column, last two number columns before a balance) */
  function guess(rows) {
    for (var r = 0; r < Math.min(rows.length, 40); r++) {
      var c = rows[r] || [], di = -1, nums = [];
      c.forEach(function (v, i) { if (di < 0 && date(v)) di = i; else if (di >= 0 && v !== '' && v != null && /^[\s(₹Rs.\-]*[\d,]+(\.\d+)?\)?\s*(cr|dr)?\.?$/i.test(String(v))) nums.push(i); });
      if (di >= 0 && nums.length >= 2) {
        var map = { date: di, desc: di + 1 < (nums[0] || 0) ? di + 1 : null };
        if (nums.length >= 3) { map.dr = nums[nums.length - 3]; map.cr = nums[nums.length - 2]; map.bal = nums[nums.length - 1]; }
        else { map.amt = nums[0]; map.bal = nums[1]; }
        return { row: r - 1, map: map, guessed: true };
      }
    }
    return null;
  }
  function statement(rows, map, from) {
    rows = (rows || []).filter(function (r) { return r && r.some(function (c) { return String(c == null ? '' : c).trim(); }); });
    var h = map ? { row: from == null ? -1 : from, map: map } : findHeader(rows) || guess(rows);
    if (!h) return { map: null, lines: [] };
    var M = h.map, lines = [];
    function cell(r, k) { return M[k] == null ? '' : r[M[k]]; }
    for (var i = h.row + 1; i < rows.length; i++) {
      var r = rows[i], d = date(cell(r, 'date'));
      if (!d) continue;
      var dr = 0, cr = 0;
      if (M.dr != null || M.cr != null) { dr = Math.abs(num(cell(r, 'dr'))); cr = Math.abs(num(cell(r, 'cr'))); }
      else {
        var a = num(cell(r, 'amt')), tp = String(cell(r, 'type') || '') + ' ' + String(cell(r, 'amt') || '');
        if (/\bcr\b|credit|deposit/i.test(tp)) cr = Math.abs(a); else if (/\bdr\b|debit|withdraw/i.test(tp)) dr = Math.abs(a); else if (a < 0) dr = -a; else cr = a;
      }
      if (!dr && !cr) continue;
      var bal = M.bal == null || cell(r, 'bal') === '' || cell(r, 'bal') == null ? null : num(cell(r, 'bal'));
      lines.push({ i: lines.length, date: d, desc: String(cell(r, 'desc') || '').replace(/\s+/g, ' ').trim(), ref: String(cell(r, 'ref') || '').trim(), dr: r2(dr), cr: r2(cr), bal: bal == null ? null : r2(bal) });
    }
    return { map: M, header: h.row, guessed: !!h.guessed, lines: lines };
  }
  /* CSV text -> rows (quotes, commas, semicolons or tabs) */
  function csvRows(text) {
    text = String(text || '').replace(/^﻿/, '');
    var first = text.split(/\r?\n/).slice(0, 20).join('\n'), sep = (first.match(/\t/g) || []).length > (first.match(/,/g) || []).length ? '\t' : (first.match(/;/g) || []).length > (first.match(/,/g) || []).length ? ';' : ',';
    var rows = [], row = [], f = '', q = false;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch; continue; }
      if (ch === '"') q = true;
      else if (ch === sep) { row.push(f); f = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(f); rows.push(row); row = []; f = ''; }
      else f += ch;
    }
    if (f !== '' || row.length) { row.push(f); rows.push(row); }
    return rows;
  }
  /* cheque or reference numbers in a line of text (5 to 12 digits) */
  function refs(s) { return (String(s || '').match(/\d{5,12}/g) || []).map(function (x) { return x.replace(/^0+/, ''); }); }

  /* the bank side of the books */
  function bookLines(docs, biz, from, to) {
    var B = G.NSBooks || (typeof NSBooks !== 'undefined' ? NSBooks : null), out = [];
    B.book(docs, biz, 'bank', from, to).rows.forEach(function (r) {
      var d = r.d;
      out.push({ id: d.id, date: r.date, inn: r.dr, out: r.cr, d: d, chq: d.chq && d.mode === 'cheque' ? d.chq.no || d.ref || '' : '', ref: d.ref || '' });
    });
    /* cheques sent to the bank but not cleared, and cheques we gave that the payee has not cashed yet */
    B.inBooks(docs, biz).forEach(function (d) {
      if (d.mode !== 'cheque' || !d.chq || (d.chq.st !== 'dep' && d.chq.st !== 'iss') || to && d.date > to) return;
      var a = d.chq.st === 'dep' ? 'chq-in' : 'chq-out', dr = 0, cr = 0;
      B.postings(d).forEach(function (l) { if (l.a === a) { dr += l.dr; cr += l.cr; } });
      if (dr || cr) out.push({ id: d.id, date: d.date, inn: r2(dr), out: r2(cr), d: d, pending: d.chq.st, chq: d.chq.no || d.ref || '', ref: d.ref || '' });
    });
    return out.sort(function (a, b) { return a.date.localeCompare(b.date); });
  }

  /* balance as per books on a date, the way a cash book's bank column shows it: the bank account, plus cheques deposited
     and not cleared yet, less cheques issued and not cashed yet */
  function bookBalance(docs, biz, to) {
    var B = G.NSBooks || (typeof NSBooks !== 'undefined' ? NSBooks : null), bal = B.book(docs, biz, 'bank', null, to).close;
    bookLines(docs, biz, to, to).forEach(function (b) { if (b.pending) bal += b.inn - b.out; });
    return r2(bal);
  }

  function match(lines, book, opts) {
    opts = opts || {};
    var tol = opts.days == null ? 10 : opts.days, pairs = {}, usedL = {}, usedB = {}, no = {};
    (opts.no || []).forEach(function (p) { no[p[0] + '|' + p[1]] = 1; });   /* pairs the user undid */
    (opts.keep || []).forEach(function (p) {
      var l = lines[p[0]], b = book.filter(function (x) { return x.id === p[1]; })[0];
      if (l && b && !usedL[l.i] && !usedB[b.id]) { pairs[l.i] = b.id; usedL[l.i] = 1; usedB[b.id] = 1; }
    });
    var cand = [];
    lines.forEach(function (l) {
      if (usedL[l.i]) return;
      var amt = l.cr || l.dr, lr = refs(l.ref + ' ' + l.desc);
      book.forEach(function (b) {
        if (usedB[b.id] || no[l.i + '|' + b.id]) return;
        var bamt = l.cr ? b.inn : b.out;
        if (Math.abs(bamt - amt) > 0.005) return;
        var gap = days(b.date, l.date);
        if (gap < -3 || gap > tol + (b.pending ? 50 : 0)) return;   /* the bank shows it on or after the book date; a pending cheque can take longer */
        var br = refs(b.chq + ' ' + b.ref), same = lr.some(function (x) { return br.indexOf(x) >= 0; });
        cand.push({ l: l.i, b: b.id, s: (same ? 0 : 1000) + Math.abs(gap) });
      });
    });
    cand.sort(function (a, b) { return a.s - b.s; }).forEach(function (c) {
      if (usedL[c.l] || usedB[c.b]) return;
      pairs[c.l] = c.b; usedL[c.l] = 1; usedB[c.b] = 1;
    });
    return {
      pairs: pairs,
      onlyBank: lines.filter(function (l) { return !usedL[l.i]; }),
      onlyBook: book.filter(function (b) { return !usedB[b.id]; })
    };
  }

  /* reconciliation statement on the last day of the period */
  function brs(bookBal, M, bankBal) {
    function sum(a, f) { return r2(a.reduce(function (s, x) { return s + f(x); }, 0)); }
    var issued = sum(M.onlyBook, function (b) { return b.out; }), deposited = sum(M.onlyBook, function (b) { return b.inn; });
    var credits = sum(M.onlyBank, function (l) { return l.cr; }), debits = sum(M.onlyBank, function (l) { return l.dr; });
    var shouldBe = r2(bookBal + issued - deposited + credits - debits);
    return { book: r2(bookBal), issued: issued, deposited: deposited, credits: credits, debits: debits, shouldBe: shouldBe,
      bank: bankBal == null || bankBal === '' ? null : r2(num(bankBal)), diff: bankBal == null || bankBal === '' ? null : r2(num(bankBal) - shouldBe) };
  }

  var API = { num: num, r2: r2, key: key, date: date, days: days, csvRows: csvRows, findHeader: findHeader, statement: statement, refs: refs, bookLines: bookLines, bookBalance: bookBalance, match: match, brs: brs };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else G.NSBank = API;
})(this);
