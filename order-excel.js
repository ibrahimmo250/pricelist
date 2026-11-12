/* Builds Excel files for orders (used by the distributor page and the admin page).
   Prices in the site are VAT-inclusive; the sheet shows them excluding 14% VAT with totals below. */
(function () {
  var VAT = 0.14;
  function load() {
    if (window.XLSX) return Promise.resolve();
    return new Promise(function (res, rej) {
      var s = document.createElement('script');
      s.src = 'xlsx.mini.min.js';
      s.onload = res;
      s.onerror = function () { rej(new Error('xlsx')); };
      document.head.appendChild(s);
    });
  }
  var r2 = function (n) { return Math.round(n * 100) / 100; };
  var ex = function (p) { return r2(Number(p) / (1 + VAT)); };
  var num = function (v, f) { return { t: 'n', v: v, z: '#,##0.00', f: f }; };
  var fmtDate = function (d) { d = new Date(d); return isNaN(d) ? '' : d.toLocaleDateString('en-GB'); };
  var safe = function (s) { return String(s || '').replace(/[^\w؀-ۿ-]+/g, '_').slice(0, 40); };

  function orderSheet(o) {
    var rows = [['رقم الطلب', o.id], ['الموزع', o.co], ['التاريخ', fmtDate(o.date)]];
    if (o.note) rows.push(['ملاحظات', o.note]);
    rows.push([]);
    rows.push(['#', 'الكود (12NC)', 'الوصف (English)', 'الوصف بالعربي', 'الكمية', 'سعر الوحدة (من غير ضريبة)', 'الإجمالي (من غير ضريبة)']);
    var first = rows.length + 1, sub = 0;
    o.items.forEach(function (it, i) {
      var u = ex(it.price), q = Number(it.qty); sub += u * q;
      rows.push([i + 1, String(it.code), it.name || '', it.ar || '', q, u, r2(u * q)]);
    });
    var last = rows.length; rows.push([]);
    var t = rows.length + 1;
    rows.push(['', '', '', '', '', 'الإجمالي (من غير ضريبة)', 0]);
    rows.push(['', '', '', '', '', 'ضريبة 14%', 0]);
    rows.push(['', '', '', '', '', 'الإجمالي شامل الضريبة', 0]);
    var ws = XLSX.utils.aoa_to_sheet(rows);
    for (var r = first; r <= last; r++) {
      ws['B' + r] = { t: 's', v: String(ws['B' + r].v) };
      ws['F' + r] = num(ws['F' + r].v);
      ws['G' + r] = num(ws['G' + r].v, 'E' + r + '*F' + r);
    }
    sub = r2(sub); var vat = r2(sub * VAT);
    ws['G' + t] = num(sub, 'SUM(G' + first + ':G' + last + ')');
    ws['G' + (t + 1)] = num(vat, 'ROUND(G' + t + '*0.14,2)');
    ws['G' + (t + 2)] = num(r2(sub + vat), 'G' + t + '+G' + (t + 1));
    ws['!cols'] = [{ wch: 5 }, { wch: 15 }, { wch: 40 }, { wch: 42 }, { wch: 9 }, { wch: 24 }, { wch: 24 }];
    return ws;
  }
  function book(sheets) {
    var wb = XLSX.utils.book_new();
    sheets.forEach(function (s) { XLSX.utils.book_append_sheet(wb, s[1], s[0]); });
    wb.Workbook = { Views: [{ RTL: true }] };
    return wb;
  }
  // one order -> one file
  window.exportOrderXlsx = function (o) {
    return load().then(function () {
      XLSX.writeFile(book([['Order', orderSheet(o)]]), 'Order_' + safe(o.id) + '_' + safe(o.co) + '.xlsx');
    });
  };
  // many orders -> one file: a lines sheet with every item, plus one sheet per order
  window.exportOrdersXlsx = function (list, name) {
    return load().then(function () {
      var rows = [['التاريخ', 'رقم الطلب', 'الموزع', 'الحالة', 'الكود (12NC)', 'الوصف (English)', 'الوصف بالعربي', 'الكمية', 'سعر الوحدة (من غير ضريبة)', 'الإجمالي (من غير ضريبة)']];
      list.forEach(function (o) {
        o.items.forEach(function (it) {
          var u = ex(it.price);
          rows.push([fmtDate(o.date), o.id, o.co, o.statusAr || o.status || '', String(it.code), it.name || '', it.ar || '', Number(it.qty), u, r2(u * it.qty)]);
        });
      });
      var ws = XLSX.utils.aoa_to_sheet(rows), n = rows.length;
      for (var r = 2; r <= n; r++) {
        ws['E' + r] = { t: 's', v: String(ws['E' + r].v) };
        ws['I' + r] = num(ws['I' + r].v);
        ws['J' + r] = num(ws['J' + r].v, 'H' + r + '*I' + r);
      }
      var sub = r2(rows.slice(1).reduce(function (a, x) { return a + x[9]; }, 0));
      XLSX.utils.sheet_add_aoa(ws, [[], ['', '', '', '', '', '', '', '', 'الإجمالي (من غير ضريبة)', 0], ['', '', '', '', '', '', '', '', 'ضريبة 14%', 0], ['', '', '', '', '', '', '', '', 'الإجمالي شامل الضريبة', 0]], { origin: -1 });
      var t = n + 2;
      ws['J' + t] = num(sub, 'SUM(J2:J' + n + ')');
      ws['J' + (t + 1)] = num(r2(sub * VAT), 'ROUND(J' + t + '*0.14,2)');
      ws['J' + (t + 2)] = num(r2(sub * (1 + VAT)), 'J' + t + '+J' + (t + 1));
      ws['!cols'] = [{ wch: 11 }, { wch: 13 }, { wch: 22 }, { wch: 12 }, { wch: 15 }, { wch: 38 }, { wch: 40 }, { wch: 8 }, { wch: 22 }, { wch: 22 }];
      var sheets = [['كل الأصناف', ws]];
      list.slice(0, 200).forEach(function (o, i) { sheets.push([(i + 1) + '-' + safe(o.id).slice(0, 25), orderSheet(o)]); });
      XLSX.writeFile(book(sheets), (name || 'Orders') + '.xlsx');
    });
  };
})();
