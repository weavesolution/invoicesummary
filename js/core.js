/* ==========================================================================
   core.js  -  invoice reading + normalisation
   --------------------------------------------------------------------------
   HOW TO ADD A NEW BILL DESIGN
   1. Copy the "global-dates-v1" parser below and rename it.
   2. Change detect() so it returns true ONLY for the new design.
   3. Change parse() so it returns the same shape (see "Parsed invoice shape").
   4. Call InvoiceCore.register(...) with it. Nothing else needs to change -
      discount allocation, tax split, validation, Excel and UI are shared.

   Parsed invoice shape (what every parser must return)
   {
     invoiceNo, date, customer, customerGstin ('' if none), sellerGstin,
     items: [{ sn, code, name, hsn, packing, qty, rate, amount }],
     totals: { totalAmount, discount, taxable, cgst, sgst, igst, invoiceValue }
   }
   ========================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.InvoiceCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------- small helpers ---------- */
  const num = (s) => {
    if (s === null || s === undefined) return NaN;
    const t = String(s).replace(/[₹,\s]/g, '');
    if (t === '' || t === '-') return 0;
    return /^-?\d+(\.\d+)?$/.test(t) ? parseFloat(t) : NaN;
  };
  const r2 = (v) => Math.round(v * 100 + (v >= 0 ? 1e-6 : -1e-6)) / 100;
  const GSTIN_RE = /^\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/;

  /* ---------- PDF -> positioned text ---------- */
  async function readPdf(pdfjsLib, data) {
    const task = pdfjsLib.getDocument({ data: data, useSystemFonts: true });
    const pdf = await task.promise;
    const pages = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const vp = page.getViewport({ scale: 1 });
      const tc = await page.getTextContent();
      const items = [];
      for (const it of tc.items) {
        const s = (it.str || '').trim();
        if (!s) continue;
        items.push({ x: it.transform[4], y: vp.height - it.transform[5], w: it.width, str: s });
      }
      pages.push({ width: vp.width, height: vp.height, items });
    }
    return { pages };
  }

  /* ---------- layout helpers used by parsers ---------- */
  const Y_TOL = 3.5;
  const sameLine = (page, y, tol) => page.items
    .filter((i) => Math.abs(i.y - y) <= (tol || Y_TOL))
    .sort((a, b) => a.x - b.x);
  const find = (page, re, where) => page.items.find((i) => re.test(i.str) &&
    (!where || ((where.xmin === undefined || i.x >= where.xmin) && (where.xmax === undefined || i.x <= where.xmax))));
  // text to the right of a label on the same line
  function rightOf(page, label, xmax) {
    if (!label) return [];
    return sameLine(page, label.y).filter((i) => i !== label && i.x >= label.x + label.w - 1 && (xmax === undefined || i.x < xmax));
  }
  const joinStr = (arr) => arr.map((i) => i.str).join(' ').replace(/\s+/g, ' ').trim();

  /* ---------- parser registry ---------- */
  const parsers = [];
  function register(p) { parsers.push(p); }

  /* ======================================================================
     DESIGN 1 : Global Dates Trading "Tax Invoice"
     (columns: S.N. | ITEM CODE | ITEM NAME | HSN CODE | PACKING | UNIT OF SALE | RATE | TOTAL AMOUNT)
     Note: the line printed as "Add : GST 2.50%" sits above "Add : CGST 2.50%"
     and is the State GST, so it is read as SGST.
     ====================================================================== */
  register({
    id: 'global-dates-v1',
    name: 'Global Dates Trading - Tax Invoice',
    detect(doc) {
      const all = doc.pages.flatMap((p) => p.items.map((i) => i.str));
      const has = (re) => all.some((s) => re.test(s));
      return has(/^ITEM NAME$/) && has(/^PACKING$/) && has(/^Total Invoice Value$/) && has(/^Add : CGST/);
    },
    parse(doc) {
      const p1 = doc.pages[0];
      const out = { items: [], totals: {} };

      // right-hand block (seller) starts at the Mail ID / Contacts labels
      const rightLbl = p1.items.filter((i) => /^(Mail ID|Contacts)$/.test(i.str)).map((i) => i.x);
      const splitX = rightLbl.length ? Math.min.apply(null, rightLbl) - 4 : 295;

      out.invoiceNo = joinStr(rightOf(p1, find(p1, /^Invoice No$/, { xmax: 120 }), splitX));
      out.date = joinStr(rightOf(p1, find(p1, /^Invoice Date$/, { xmax: 120 }), splitX));
      out.customer = joinStr(rightOf(p1, find(p1, /^Customer Name$/, { xmax: 120 }), splitX));
      const custG = find(p1, /^GSTIN$/, { xmax: splitX });
      out.customerGstin = joinStr(rightOf(p1, custG, splitX)).replace(/\s+/g, '').toUpperCase();
      const sellG = find(p1, /^GSTIN$/, { xmin: splitX });
      out.sellerGstin = joinStr(rightOf(p1, sellG)).replace(/\s+/g, '').toUpperCase();

      // item rows (may continue over several pages)
      for (const page of doc.pages) {
        const stop = page.items.filter((i) => /^(Amount in Words|Total Amount)/.test(i.str) && i.x > 30)
          .map((i) => i.y);
        const tableEnd = stop.length ? Math.min.apply(null, stop) - 2 : page.height;
        const hdr = find(page, /^ITEM NAME$/);
        const tableStart = hdr ? hdr.y + 14 : 0;

        const rowsY = [];
        const seen = [];
        page.items.filter((i) => i.x < 70 && /^\d{1,3}$/.test(i.str) && i.y > tableStart && i.y < tableEnd)
          .sort((a, b) => a.y - b.y).forEach((i) => { if (!seen.some((y) => Math.abs(y - i.y) < 4)) { seen.push(i.y); rowsY.push(i.y); } });

        rowsY.forEach((y, idx) => {
          const cells = sameLine(page, y);
          const iH = cells.findIndex((c, k) => k >= 2 && /^\d{4,8}$/.test(c.str));
          if (iH < 0) { out.items.push({ _bad: true, sn: cells[0].str, raw: joinStr(cells) }); return; }
          const sn = cells[0].str;
          let nameStart = 1, code = '';
          if (/^\d+$/.test(cells[1].str) && iH >= 3) { code = cells[1].str; nameStart = 2; }
          let name = joinStr(cells.slice(nameStart, iH));
          const after = cells.slice(iH + 1);
          if (after.length < 3) { out.items.push({ _bad: true, sn, raw: joinStr(cells) }); return; }
          const amount = after[after.length - 1].str;
          const rate = after[after.length - 2].str;
          const qty = after[after.length - 3].str;
          const packing = joinStr(after.slice(0, after.length - 3));
          // wrapped item-name lines (no serial number) up to the next row
          const nextY = rowsY[idx + 1] !== undefined ? rowsY[idx + 1] : tableEnd;
          page.items.filter((i) => i.y > y + 6 && i.y < nextY - 4 && i.x >= cells[nameStart].x - 4 && i.x < cells[iH].x - 4)
            .sort((a, b) => a.y - b.y || a.x - b.x).forEach((i) => { name += ' ' + i.str; });
          out.items.push({ sn, code, name, hsn: cells[iH].str, packing, qty: num(qty), rate: num(rate), amount: num(amount) });
        });
      }

      // totals block (right side, label -> right-most value on the same line)
      const val = (re) => {
        for (const page of doc.pages) {
          const lbl = find(page, re, { xmin: 250 });
          if (lbl) {
            const r = rightOf(page, lbl).filter((i) => i.x > lbl.x + lbl.w + 20);
            if (r.length) return num(r[r.length - 1].str);
            return 0;
          }
        }
        return NaN;
      };
      out.totals = {
        totalAmount: val(/^Total Amount$/),
        discount: val(/^Discount$/),
        taxable: val(/^Taxable Amount$/),
        sgst: val(/^Add : GST/),
        cgst: val(/^Add : CGST/),
        igst: val(/^Add : IGST/),
        invoiceValue: val(/^Total Invoice Value$/),
      };
      return out;
    },
  });

  /* ---------- run the right parser ---------- */
  function parseDocument(doc) {
    for (const p of parsers) {
      let ok = false;
      try { ok = p.detect(doc); } catch (e) { ok = false; }
      if (ok) {
        const inv = p.parse(doc);
        inv._parser = p.id;
        return inv;
      }
    }
    return null; // unknown design
  }

  /* ======================================================================
     Normalisation (shared by all designs):
       - discount: whole invoice discount goes to item 1
       - GST: split across items in proportion to their taxable value,
         last item absorbs paise so the invoice ties to the printed GST
       - validation against the printed totals
     options.tieToInvoiceValue : push the round-off (paise) into the last
       item's Total Amount so the invoice equals the printed Invoice Value.
     ====================================================================== */
  function normalise(inv, options) {
    options = options || {};
    const warn = [];
    const err = [];
    const T = inv.totals;
    const nz = (v) => (isNaN(v) ? 0 : v);

    if (!inv.invoiceNo) err.push('Invoice number not found');
    if (!inv.items.length) err.push('No item rows found');
    inv.items.filter((i) => i._bad).forEach((i) => err.push('Item row ' + i.sn + ' could not be read: ' + i.raw));
    const items = inv.items.filter((i) => !i._bad);
    items.forEach((i) => {
      if (isNaN(i.amount) || isNaN(i.qty)) err.push('Item ' + i.sn + ' has an unreadable qty/amount');
    });
    ['totalAmount', 'invoiceValue'].forEach((k) => { if (isNaN(T[k])) err.push('Printed ' + k + ' not found'); });

    // GSTIN
    let gst = inv.customerGstin;
    const registered = !!gst;
    if (!registered) gst = 'non-registered';
    else if (!GSTIN_RE.test(gst)) warn.push('Customer GSTIN "' + gst + '" does not look like a valid GSTIN');

    const discount = nz(T.discount);
    const gross = items.reduce((s, i) => s + nz(i.amount), 0);
    const sumTotalAmt = nz(T.totalAmount);
    const tol = 0.5 * items.length + 1; // item amounts are printed rounded to the rupee
    if (Math.abs(gross - sumTotalAmt) > tol) err.push('Item amounts add to ' + gross.toFixed(2) + ' but the bill says Total Amount ' + sumTotalAmt.toFixed(2));

    const printedTaxable = isNaN(T.taxable) ? sumTotalAmt - discount : T.taxable;
    if (Math.abs((sumTotalAmt - discount) - printedTaxable) > tol) warn.push('Total Amount - Discount (' + (sumTotalAmt - discount).toFixed(2) + ') differs from printed Taxable Amount (' + printedTaxable.toFixed(2) + ')');

    const rows = items.map((it, idx) => ({
      invoiceNo: /^\d+$/.test(inv.invoiceNo) ? Number(inv.invoiceNo) : inv.invoiceNo,
      gst: gst,
      hsn: it.hsn,
      itemName: it.name,
      qty: it.qty,
      packaging: it.packing,
      discount: idx === 0 ? discount : 0,
      amount: it.amount,
      cgst: 0, sgst: 0, igst: 0,
    }));
    if (rows.length && rows[0].amount < rows[0].discount) warn.push('Discount is larger than the first item amount');

    // GST allocation
    const taxable = rows.map((r) => r.amount - r.discount);
    const taxableSum = taxable.reduce((s, v) => s + v, 0);
    ['cgst', 'sgst', 'igst'].forEach((k) => {
      const tot = nz(T[k]);
      if (!rows.length) return;
      if (!tot) return;
      let used = 0;
      rows.forEach((r, i) => {
        if (i < rows.length - 1) {
          r[k] = r2(taxableSum ? (tot * taxable[i]) / taxableSum : 0);
          used += r[k];
        } else {
          r[k] = r2(tot - used);
        }
      });
    });

    // round-off tie-out
    const computedTotal = r2(taxableSum + nz(T.cgst) + nz(T.sgst) + nz(T.igst));
    const diff = r2(nz(T.invoiceValue) - computedTotal);
    if (Math.abs(diff) > tol) err.push('Computed invoice total ' + computedTotal.toFixed(2) + ' does not match printed Total Invoice Value ' + nz(T.invoiceValue).toFixed(2));
    let tied = false;
    if (options.tieToInvoiceValue && rows.length && Math.abs(diff) <= tol && Math.abs(diff) > 0.0001) {
      // printed item amounts are rounded to the rupee; rescale them (paise only) so that
      // taxable + GST equals the printed Invoice Value exactly
      const targetGross = r2(nz(T.invoiceValue) - nz(T.cgst) - nz(T.sgst) - nz(T.igst) + discount);
      const g = rows.reduce((a, r) => a + r.amount, 0);
      let used = 0;
      rows.forEach((r, i) => {
        if (i < rows.length - 1) { r.amount = r2(g ? (r.amount * targetGross) / g : r.amount); used += r.amount; }
        else r.amount = r2(targetGross - used);
      });
      tied = true;
    }

    rows.forEach((r) => { r.total = r2(r.amount - r.discount + r.cgst + r.sgst + r.igst); });
    const rowsTotal = r2(rows.reduce((s, r) => s + r.total, 0));

    return {
      rows,
      summary: {
        invoiceNo: rows.length ? rows[0].invoiceNo : inv.invoiceNo,
        date: inv.date,
        customer: inv.customer,
        gst: gst,
        registered,
        items: rows.length,
        totalAmount: r2(rows.reduce((s, r) => s + r.amount, 0)),
        discount: r2(discount),
        cgst: r2(rows.reduce((s, r) => s + r.cgst, 0)),
        sgst: r2(rows.reduce((s, r) => s + r.sgst, 0)),
        igst: r2(rows.reduce((s, r) => s + r.igst, 0)),
        total: rowsTotal,
        printedValue: nz(T.invoiceValue),
        roundOff: r2(nz(T.invoiceValue) - rowsTotal),
        tied,
      },
      warnings: warn,
      errors: err,
      status: err.length ? 'error' : warn.length ? 'warn' : 'ok',
    };
  }

  return { register, parsers, readPdf, parseDocument, normalise, r2, num };
});
