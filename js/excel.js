/* ==========================================================================
   excel.js - builds the formatted workbook with ExcelJS
   Sheet 1 "Invoice Summary" : the 12 requested columns, one row per item
   Sheet 2 "Invoice Check"   : one row per invoice, tie-out to the printed bill
   ========================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.InvoiceExcel = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const INK = 'FF0F3D47';      // deep teal (header)
  const INK_SOFT = 'FFE4EEF0';  // banded invoice tint
  const LINE = 'FFB9CCD1';
  const MUTED = 'FF6B7F85';
  const FONT = 'Arial';

  // Indian digit grouping (12,34,567.89). Zero shown as a dash in the columns that are mostly nil.
  const FMT_AMT = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00';
  const FMT_AMT_DASH = '[>=100000]##\\,##\\,##0.00;[=0]"–";##,##0.00';

  const COLS = [
    { key: 'invoiceNo', header: 'Invoice No', width: 12, align: 'center' },
    { key: 'gst', header: 'GST number', width: 20, align: 'left' },
    { key: 'hsn', header: 'HSN Code', width: 12, align: 'center' },
    { key: 'itemName', header: 'Item Name', width: 40, align: 'left' },
    { key: 'qty', header: 'Qty', width: 8, align: 'center' },
    { key: 'packaging', header: 'Packaging', width: 13, align: 'center' },
    { key: 'discount', header: 'Discount Allocated', width: 15, align: 'right', fmt: FMT_AMT_DASH },
    { key: 'amount', header: 'Total Amount', width: 16, align: 'right', fmt: FMT_AMT },
    { key: 'cgst', header: 'CGST', width: 13, align: 'right', fmt: FMT_AMT_DASH },
    { key: 'sgst', header: 'SGST', width: 13, align: 'right', fmt: FMT_AMT_DASH },
    { key: 'igst', header: 'IGST', width: 13, align: 'right', fmt: FMT_AMT_DASH },
    { key: 'total', header: 'Total Incl. GST', width: 17, align: 'right', fmt: FMT_AMT },
  ];

  const colLetter = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const thin = (argb) => ({ style: 'thin', color: { argb } });

  /**
   * @param ExcelJS   the ExcelJS library
   * @param results   array of normalise() results (only usable ones)
   * @param opts      { tied: boolean, generatedOn: Date }
   */
  function build(ExcelJS, results, opts) {
    opts = opts || {};
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Invoice Summariser';
    wb.created = opts.generatedOn || new Date();

    results = results.slice().sort((a, b) => {
      const x = a.summary.invoiceNo, y = b.summary.invoiceNo;
      return typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true });
    });

    /* ---------------- Sheet 1 : Invoice Summary ---------------- */
    const ws = wb.addWorksheet('Invoice Summary', {
      views: [{ state: 'frozen', ySplit: 4, showGridLines: false }],
      pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '4:4', margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } },
      headerFooter: { oddFooter: '&LInvoice Summary&RPage &P of &N' },
    });
    ws.columns = COLS.map((c) => ({ key: c.key, width: c.width }));
    const nCols = COLS.length;
    const last = colLetter(nCols);

    const rowsAll = results.flatMap((r) => r.rows);
    const first = rowsAll.length ? results[0].summary.invoiceNo : '';
    const lastNo = rowsAll.length ? results[results.length - 1].summary.invoiceNo : '';
    const range = results.length > 1 ? 'Invoices ' + first + ' to ' + lastNo : 'Invoice ' + first;

    ws.mergeCells('A1:' + last + '1');
    const t = ws.getCell('A1');
    t.value = 'Invoice Summary';
    t.font = { name: FONT, size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: INK } };
    t.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(1).height = 30;

    ws.mergeCells('A2:' + last + '2');
    const sub = ws.getCell('A2');
    const d = (opts.generatedOn || new Date());
    const dd = String(d.getDate()).padStart(2, '0') + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + d.getFullYear();
    sub.value = range + '   |   ' + results.length + ' invoices, ' + rowsAll.length + ' line items   |   Generated ' + dd;
    sub.font = { name: FONT, size: 9, color: { argb: MUTED } };
    sub.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(2).height = 18;
    ws.getRow(3).height = 6;

    const H = 4;
    const hr = ws.getRow(H);
    COLS.forEach((c, i) => {
      const cell = hr.getCell(i + 1);
      cell.value = c.header;
      cell.font = { name: FONT, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: INK } };
      cell.alignment = { vertical: 'middle', horizontal: c.align === 'right' ? 'right' : 'center', wrapText: true };
      cell.border = { bottom: { style: 'medium', color: { argb: 'FFE09F3E' } } };
    });
    hr.height = 32;

    let r = H + 1;
    const firstData = r;
    results.forEach((res, idx) => {
      const band = idx % 2 === 0;
      res.rows.forEach((row, k) => {
        const xr = ws.getRow(r);
        COLS.forEach((c, i) => {
          const cell = xr.getCell(i + 1);
          if (c.key === 'total') {
            cell.value = { formula: 'H' + r + '-G' + r + '+I' + r + '+J' + r + '+K' + r, result: row.total };
          } else {
            cell.value = row[c.key];
          }
          cell.font = { name: FONT, size: 10, color: { argb: 'FF1B2A2E' } };
          if (c.fmt) cell.numFmt = c.fmt;
          if (c.key === 'invoiceNo') { cell.numFmt = '0'; cell.font = { name: FONT, size: 10, bold: true, color: { argb: 'FF1B2A2E' } }; }
          if (c.key === 'hsn') cell.numFmt = '@';
          if (c.key === 'total') cell.font = { name: FONT, size: 10, bold: true, color: { argb: 'FF1B2A2E' } };
          if (c.key === 'gst' && row.gst === 'non-registered') cell.font = { name: FONT, size: 10, italic: true, color: { argb: 'FF9A6A12' } };
          cell.alignment = { vertical: 'middle', horizontal: c.align, indent: c.align === 'left' ? 1 : 0, wrapText: c.key === 'itemName' };
          if (band) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: INK_SOFT } };
          cell.border = {
            top: k === 0 ? thin('FF7F9CA3') : thin(LINE),
            bottom: thin(LINE),
          };
        });
        xr.height = 20;
        r++;
      });
    });
    const lastData = r - 1;

    // totals row (SUBTOTAL so it follows any filter the user applies)
    const tr = ws.getRow(r);
    ws.mergeCells('A' + r + ':F' + r);
    const tl = tr.getCell(1);
    tl.value = 'Total';
    tl.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
    const sums = { 7: 'discount', 8: 'amount', 9: 'cgst', 10: 'sgst', 11: 'igst', 12: 'total' };
    const sumVal = (k) => Math.round(rowsAll.reduce((s, x) => s + x[k], 0) * 100) / 100;
    for (let i = 1; i <= nCols; i++) {
      const cell = tr.getCell(i);
      if (sums[i] && rowsAll.length) {
        const L = colLetter(i);
        cell.value = { formula: 'SUBTOTAL(109,' + L + firstData + ':' + L + lastData + ')', result: sumVal(sums[i]) };
        cell.numFmt = i === 8 || i === 12 ? FMT_AMT : FMT_AMT_DASH;
      }
      cell.font = { name: FONT, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: INK } };
      if (i > 6) cell.alignment = { horizontal: 'right', vertical: 'middle' };
    }
    tr.height = 24;

    if (rowsAll.length) ws.autoFilter = { from: { row: H, column: 1 }, to: { row: lastData, column: nCols } };

    // footnotes
    const notes = [
      'Discount: the full bill discount is allocated to the first item of each invoice. GST: split across items in proportion to taxable value.',
      opts.tied
        ? 'Item amounts are printed in whole rupees on the bills; paise have been adjusted so every invoice ties exactly to its printed Invoice Value (see "Invoice Check").'
        : 'Item amounts are as printed (whole rupees); invoice totals may differ from the printed Invoice Value by round-off (see "Invoice Check").',
    ];
    notes.forEach((n, i) => {
      const nr = r + 2 + i;
      ws.mergeCells('A' + nr + ':' + last + nr);
      const c = ws.getCell('A' + nr);
      c.value = n;
      c.font = { name: FONT, size: 8.5, italic: true, color: { argb: MUTED } };
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };
      ws.getRow(nr).height = 16;
    });

    /* ---------------- Sheet 2 : Invoice Check ---------------- */
    const cs = wb.addWorksheet('Invoice Check', {
      views: [{ state: 'frozen', ySplit: 3, showGridLines: false }],
      pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    const CC = [
      { h: 'Invoice No', w: 12, a: 'center' },
      { h: 'Date', w: 13, a: 'center' },
      { h: 'Customer', w: 34, a: 'left' },
      { h: 'GST number', w: 20, a: 'left' },
      { h: 'Items', w: 8, a: 'center' },
      { h: 'Taxable value', w: 16, a: 'right', f: FMT_AMT },
      { h: 'CGST', w: 13, a: 'right', f: FMT_AMT },
      { h: 'SGST', w: 13, a: 'right', f: FMT_AMT },
      { h: 'IGST', w: 12, a: 'right', f: FMT_AMT_DASH },
      { h: 'Total Incl. GST', w: 16, a: 'right', f: FMT_AMT },
      { h: 'Printed Invoice Value', w: 18, a: 'right', f: FMT_AMT },
      { h: 'Difference', w: 12, a: 'right', f: '0.00;-0.00;"–"' },
      { h: 'Status', w: 46, a: 'left' },
    ];
    cs.columns = CC.map((c) => ({ width: c.w }));
    const lastC = colLetter(CC.length);
    cs.mergeCells('A1:' + lastC + '1');
    const ct = cs.getCell('A1');
    ct.value = 'Invoice Check  -  each invoice compared with the total printed on the bill';
    ct.font = { name: FONT, size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
    ct.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: INK } };
    ct.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cs.getRow(1).height = 26;
    cs.getRow(2).height = 6;
    const ch = cs.getRow(3);
    CC.forEach((c, i) => {
      const cell = ch.getCell(i + 1);
      cell.value = c.h;
      cell.font = { name: FONT, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: INK } };
      cell.alignment = { vertical: 'middle', horizontal: c.a === 'right' ? 'right' : 'center', wrapText: true };
      cell.border = { bottom: { style: 'medium', color: { argb: 'FFE09F3E' } } };
    });
    ch.height = 30;

    results.forEach((res, idx) => {
      const s = res.summary;
      const rr = 4 + idx;
      const taxable = Math.round((s.totalAmount - s.discount) * 100) / 100;
      const vals = [s.invoiceNo, s.date, s.customer, s.gst, s.items, taxable, s.cgst, s.sgst, s.igst, s.total, s.printedValue, null, null];
      const row = cs.getRow(rr);
      CC.forEach((c, i) => {
        const cell = row.getCell(i + 1);
        if (i === 11) cell.value = { formula: 'ROUND(K' + rr + '-J' + rr + ',2)', result: Math.round((s.printedValue - s.total) * 100) / 100 };
        else if (i === 12) {
          cell.value = res.status === 'ok' ? 'OK - ties to printed total' : (res.errors.concat(res.warnings)).join('; ');
          const colr = res.status === 'ok' ? 'FF1F7A4D' : res.status === 'warn' ? 'FF9A6A12' : 'FFB3261E';
          cell.font = { name: FONT, size: 10, bold: true, color: { argb: colr } };
        } else cell.value = vals[i];
        if (i !== 12) cell.font = { name: FONT, size: 10, color: { argb: 'FF1B2A2E' }, bold: i === 0, italic: i === 3 && s.gst === 'non-registered' };
        if (i === 3 && s.gst === 'non-registered') cell.font = { name: FONT, size: 10, italic: true, color: { argb: 'FF9A6A12' } };
        if (c.f) cell.numFmt = c.f;
        if (i === 0) cell.numFmt = '0';
        cell.alignment = { vertical: 'middle', horizontal: c.a, indent: c.a === 'left' ? 1 : 0, wrapText: i === 12 };
        if (idx % 2 === 0) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: INK_SOFT } };
        cell.border = { top: thin(LINE), bottom: thin(LINE) };
      });
      row.height = 20;
    });

    return wb;
  }

  return { build, COLS };
});
