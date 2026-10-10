/* app.js - wires the page to InvoiceCore + InvoiceExcel */
(function () {
  'use strict';
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';

  const $ = (id) => document.getElementById(id);
  const drop = $('drop'), picker = $('picker'), tie = $('tie');
  const entries = []; // { name, state:'ok'|'warn'|'error'|'unknown', parsed, result, message }
  let busy = 0;

  const inr = (v) => v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  /* ---------- file intake ---------- */
  async function collect(fileList) {
    const out = [];
    for (const f of fileList) {
      const lower = f.name.toLowerCase();
      if (lower.endsWith('.pdf')) out.push({ name: f.name, get: () => f.arrayBuffer() });
      else if (lower.endsWith('.zip')) {
        try {
          const zip = await JSZip.loadAsync(f);
          zip.forEach((path, entry) => {
            if (entry.dir || /(^|\/)(__MACOSX|\.)/.test(path) || !path.toLowerCase().endsWith('.pdf')) return;
            out.push({ name: path.split('/').pop(), get: () => entry.async('arraybuffer') });
          });
        } catch (e) {
          entries.push({ name: f.name, state: 'error', message: 'Could not open this ZIP file.' });
        }
      } else {
        entries.push({ name: f.name, state: 'error', message: 'Only PDF or ZIP files are supported.' });
      }
    }
    return out;
  }

  async function handle(fileList) {
    busy++;
    setBusy(true);
    const items = await collect(Array.from(fileList));
    for (const it of items) {
      try {
        const doc = await InvoiceCore.readPdf(pdfjsLib, new Uint8Array(await it.get()));
        const parsed = InvoiceCore.parseDocument(doc);
        if (!parsed) {
          entries.push({ name: it.name, state: 'unknown', message: 'Bill layout not recognised yet. Send this PDF to have the layout added.' });
        } else {
          entries.push({ name: it.name, parsed });
        }
      } catch (e) {
        entries.push({ name: it.name, state: 'error', message: 'Could not read this PDF (' + (e && e.message ? e.message : 'unknown error') + ').' });
      }
    }
    busy--;
    refresh();
  }

  /* ---------- (re)normalise everything and draw ---------- */
  function refresh() {
    const seen = new Set();
    entries.forEach((e) => {
      if (!e.parsed) return;
      e.result = InvoiceCore.normalise(e.parsed, { tieToInvoiceValue: tie.checked });
      const no = String(e.result.summary.invoiceNo);
      if (e.result.rows.length && seen.has(no)) {
        e.duplicate = true;
        e.state = 'error';
        e.message = 'Invoice ' + no + ' was already added - skipped.';
      } else {
        e.duplicate = false;
        seen.add(no);
        e.state = e.result.status;
        e.message = e.result.errors.concat(e.result.warnings).join(' ');
      }
    });
    render();
  }

  function usable() {
    return entries.filter((e) => e.result && !e.duplicate && e.result.rows.length);
  }

  function setBusy(on) {
    drop.classList.toggle('busy', on);
    $('dropTitle').textContent = on ? 'Reading invoices...' : 'Drop invoice PDFs or a ZIP here';
  }

  function render() {
    setBusy(busy > 0);
    const has = entries.length > 0;
    $('results').hidden = !has;
    if (!has) return;

    // files list
    const label = { ok: 'Read', warn: 'Check', error: 'Not added', unknown: 'Not recognised' };
    $('files').innerHTML = entries.map((e) => {
      const st = e.state || 'error';
      const inv = e.result && !e.duplicate && e.result.rows.length ? ' <span class="muted">Invoice ' + esc(e.result.summary.invoiceNo) + ' &middot; ' + e.result.rows.length + ' item' + (e.result.rows.length > 1 ? 's' : '') + '</span>' : '';
      return '<li class="file ' + st + '"><span class="chip">' + label[st] + '</span><span class="fname">' + esc(e.name) + inv + '</span>' +
        (e.message ? '<span class="msg">' + esc(e.message) + '</span>' : '') + '</li>';
    }).join('');

    const good = usable();
    const rows = good.flatMap((e) => e.result.rows);
    const sum = (k) => rows.reduce((s, r) => s + r[k], 0);
    $('sInv').textContent = good.length;
    $('sItems').textContent = rows.length;
    $('sAmt').textContent = inr(sum('amount') - sum('discount'));
    $('sGst').textContent = inr(sum('cgst') + sum('sgst') + sum('igst'));
    $('sTot').textContent = inr(sum('total'));
    $('stats').hidden = !good.length;
    $('dl').disabled = !good.length;

    // preview table
    const sorted = good.slice().sort((a, b) => {
      const x = a.result.summary.invoiceNo, y = b.result.summary.invoiceNo;
      return typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true });
    });
    const body = [];
    sorted.forEach((e, i) => {
      e.result.rows.forEach((r, k) => {
        body.push('<tr class="' + (i % 2 ? '' : 'band') + (k === 0 ? ' first' : '') + '">' +
          '<td class="c b">' + esc(r.invoiceNo) + '</td>' +
          '<td class="' + (r.gst === 'non-registered' ? 'nr' : '') + '">' + esc(r.gst) + '</td>' +
          '<td class="c">' + esc(r.hsn) + '</td><td>' + esc(r.itemName) + '</td><td class="c">' + r.qty + '</td><td class="c">' + esc(r.packaging) + '</td>' +
          '<td class="n">' + (r.discount ? inr(r.discount) : '&ndash;') + '</td><td class="n">' + inr(r.amount) + '</td>' +
          '<td class="n">' + (r.cgst ? inr(r.cgst) : '&ndash;') + '</td><td class="n">' + (r.sgst ? inr(r.sgst) : '&ndash;') + '</td><td class="n">' + (r.igst ? inr(r.igst) : '&ndash;') + '</td>' +
          '<td class="n b">' + inr(r.total) + '</td></tr>');
      });
    });
    $('tbody').innerHTML = body.join('');
    $('tableWrap').hidden = !body.length;
  }

  /* ---------- download ---------- */
  async function download() {
    const good = usable();
    if (!good.length) return;
    const wb = InvoiceExcel.build(ExcelJS, good.map((e) => e.result), { tied: tie.checked, generatedOn: new Date() });
    const buf = await wb.xlsx.writeBuffer();
    const nos = good.map((e) => e.result.summary.invoiceNo).sort((a, b) => (typeof a === 'number' && typeof b === 'number' ? a - b : String(a).localeCompare(String(b), undefined, { numeric: true })));
    const name = 'Invoice_Summary_' + nos[0] + (nos.length > 1 ? '_to_' + nos[nos.length - 1] : '') + '.xlsx';
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  /* ---------- events ---------- */
  picker.addEventListener('change', () => { if (picker.files.length) handle(picker.files); picker.value = ''; });
  $('pick').addEventListener('click', () => picker.click());
  drop.addEventListener('click', (e) => { if (e.target === drop || e.target.closest('.dropInner') && !e.target.closest('button')) picker.click(); });
  drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); picker.click(); } });
  ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => { if (e.dataTransfer.files.length) handle(e.dataTransfer.files); });
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());
  tie.addEventListener('change', refresh);
  $('dl').addEventListener('click', download);
  $('clear').addEventListener('click', () => { entries.length = 0; render(); });
})();
