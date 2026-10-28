/* global XLSX */
(() => {
  'use strict';

  const STATE = {
    running: false,
    stopRequested: false,
    delayMs: 1200,
    rows: [],
    currentIndex: 0
  };

  function log(msg) {
    const panel = document.getElementById('egate-importer-panel');
    const box = document.getElementById('egate-importer-log');
    const line = `[${new Date().toLocaleTimeString('pt-BR')}] ${msg}`;
    if (box) {
      box.value += line + '\n';
      box.scrollTop = box.scrollHeight;
    } else {
      console.log('[eGate Importer]', msg);
    }
    if (panel) panel.dataset.last = String(Date.now());
  }

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  function removeReadonly() {
    document.querySelectorAll('input[readonly], textarea[readonly]').forEach(el => {
      el.removeAttribute('readonly');
      el.readOnly = false;
    });
  }

  function setNativeValue(element, value) {
    const proto = Object.getPrototypeOf(element);
    const desc = Object.getOwnPropertyDescriptor(proto, 'value') ||
                 Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
    if (desc && desc.set) desc.set.call(element, value);
    else element.value = value;
  }

  function dispatchAll(element) {
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
    element.dispatchEvent(new Event('blur', { bubbles: true }));
  }

  // ======= CORRIGIDO AQUI =======
  // Agora trata também strings no formato M/D/YY ou M/D/YYYY (ex: "1/18/26" => "18/01/2026")
  function fmtDateBR(v) {
    if (!v) return '';

    // Excel/SheetJS: Date real
    if (v instanceof Date && !isNaN(v)) {
      const dd = String(v.getDate()).padStart(2, '0');
      const mm = String(v.getMonth() + 1).padStart(2, '0');
      const yy = String(v.getFullYear());
      return `${dd}/${mm}/${yy}`;
    }

    // Número (raro aqui): tenta converter se parecer serial Excel
    if (typeof v === 'number' && isFinite(v)) {
      // Excel serial (1900-based, com bug do 1900). SheetJS geralmente já devolve Date quando cellDates=true.
      // Mesmo assim, se vier número, tenta converter com epoch 1899-12-30.
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const d = new Date(excelEpoch.getTime() + v * 86400000);
      if (!isNaN(d)) {
        const dd = String(d.getUTCDate()).padStart(2, '0');
        const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
        const yy = String(d.getUTCFullYear());
        return `${dd}/${mm}/${yy}`;
      }
    }

    // Strings
    if (typeof v === 'string') {
      const s = v.trim();
      if (!s) return '';

      // Já está em DD/MM/AAAA
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;

      // YYYY-MM-DD => DD/MM/AAAA
      const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;

      // M/D/YY ou M/D/YYYY ou DD/MM/YY etc.
      // Ex: 1/18/26 (mm/dd/yy) => 18/01/2026
      const md = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
      if (md) {
        let a = parseInt(md[1], 10);
        let b = parseInt(md[2], 10);
        let y = md[3].length === 2 ? parseInt(md[3], 10) : parseInt(md[3], 10);

        // expande ano 2 dígitos (assume 2000-2099)
        if (md[3].length === 2) y = 2000 + y;

        // decidir se é dd/mm ou mm/dd:
        // - se a > 12 => a é dia
        // - se b > 12 => b é dia (caso 1/18/26)
        // - se ambos <= 12 => assume BR (dd/mm)
        let day, month;
        if (a > 12 && b <= 12) { day = a; month = b; }
        else if (b > 12 && a <= 12) { day = b; month = a; }
        else { day = a; month = b; }

        const dd = String(day).padStart(2, '0');
        const mm = String(month).padStart(2, '0');
        const yy = String(y).padStart(4, '0');
        return `${dd}/${mm}/${yy}`;
      }

      return s;
    }

    return String(v);
  }
  // ======= FIM DA CORREÇÃO =======

  function fmtMoneyBR(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number' && isFinite(v)) {
      return v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    const s = String(v).trim();
    if (/^\d{1,3}(\.\d{3})*,\d{2}$/.test(s) || /^\d+,\d{2}$/.test(s)) return s;
    if (/^\d+(\.\d+)?$/.test(s)) {
      const n = Number(s);
      if (isFinite(n)) return n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return s;
  }

  function getEl(id) {
    const el = document.getElementById(id);
    if (!el) throw new Error(`Não encontrei o campo com id="${id}" nesta tela.`);
    return el;
  }

  async function setField(id, value) {
    const el = getEl(id);
    removeReadonly();
    el.scrollIntoView({ block: 'center', behavior: 'smooth' });
    el.focus();
    setNativeValue(el, value);
    dispatchAll(el);
    await sleep(30);
  }

  function findActionButtonsContainer() {
    const candidates = Array.from(document.querySelectorAll('a, button, input[type="button"], input[type="submit"]'));
    const salvar = candidates.find(el => /salvar/i.test((el.innerText || el.value || '').trim()) && /(F8|\(F8\))/i.test((el.innerText || el.value || '')));
    const voltar = candidates.find(el => /voltar/i.test((el.innerText || el.value || '').trim()) && /(F9|\(F9\))/i.test((el.innerText || el.value || '')));
    if (salvar && voltar && salvar.parentElement) return { salvar, voltar, container: salvar.parentElement };

    const salvar2 = candidates.find(el => /salvar/i.test((el.innerText || el.value || '').trim()));
    const voltar2 = candidates.find(el => /voltar/i.test((el.innerText || el.value || '').trim()));
    if (salvar2 && voltar2 && salvar2.parentElement) return { salvar: salvar2, voltar: voltar2, container: salvar2.parentElement };
    return null;
  }

  function ensurePanel() {
    if (document.getElementById('egate-importer-panel')) return;

    const panel = document.createElement('div');
    panel.id = 'egate-importer-panel';
    panel.style.position = 'fixed';
    panel.style.bottom = '16px';
    panel.style.right = '16px';
    panel.style.zIndex = '999999';
    panel.style.width = '360px';
    panel.style.background = '#fff';
    panel.style.border = '1px solid rgba(0,0,0,0.15)';
    panel.style.borderRadius = '10px';
    panel.style.boxShadow = '0 10px 24px rgba(0,0,0,0.18)';
    panel.style.fontFamily = 'Arial, sans-serif';

    panel.innerHTML = `
      <div style="padding:10px 12px; border-bottom:1px solid rgba(0,0,0,0.08); display:flex; align-items:center; justify-content:space-between;">
        <div style="font-weight:700;">Importador e-Gate</div>
        <button id="egate-importer-close" style="border:none;background:transparent;font-size:16px;cursor:pointer;">✕</button>
      </div>
      <div style="padding:10px 12px;">
        <div style="display:flex; gap:8px; align-items:center; margin-bottom:8px;">
          <label style="font-size:12px;">Delay (ms)</label>
          <input id="egate-importer-delay" type="number" min="0" step="100" value="${STATE.delayMs}" style="flex:1; padding:6px 8px; border:1px solid rgba(0,0,0,0.15); border-radius:8px;">
          <button id="egate-importer-stop" style="padding:6px 10px; border-radius:8px; border:1px solid rgba(0,0,0,0.15); background:#f7f7f7; cursor:pointer;">Parar</button>
        </div>
        <textarea id="egate-importer-log" readonly style="width:100%; height:180px; resize:vertical; font-family:Consolas, monospace; font-size:11px; padding:8px; border:1px solid rgba(0,0,0,0.15); border-radius:8px; background:#fafafa;"></textarea>
        <div id="egate-importer-status" style="margin-top:8px; font-size:12px; color:#333;">Pronto.</div>
      </div>
    `;
    document.documentElement.appendChild(panel);

    document.getElementById('egate-importer-close').addEventListener('click', () => panel.remove());
    document.getElementById('egate-importer-stop').addEventListener('click', () => {
      STATE.stopRequested = true;
      log('Parada solicitada.');
    });
    document.getElementById('egate-importer-delay').addEventListener('change', (e) => {
      const v = Number(e.target.value);
      if (!Number.isFinite(v) || v < 0) return;
      STATE.delayMs = v;
      log(`Delay ajustado para ${STATE.delayMs} ms.`);
    });
  }

  function setStatus(text) {
    const el = document.getElementById('egate-importer-status');
    if (el) el.textContent = text;
  }

  function normalizeHeader(h) {
    return String(h || '')
      .trim()
      .toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, '');
  }

  function validateAndMapRows(rows) {
    const required = [
      'numerodocumento',
      'datadocumento',
      'datarecebimento',
      'datavencimento',
      'numeropedido',
      'valordocumento'
    ];

    const keyMap = {};
    if (rows.length) {
      Object.keys(rows[0]).forEach(k => {
        keyMap[normalizeHeader(k)] = k;
      });
    }

    const missing = required.filter(r => !keyMap[r]);
    if (missing.length) {
      throw new Error(
        `Colunas obrigatórias não encontradas: ${missing.join(', ')}.\n` +
        `Use exatamente: NumeroDocumento, DataDocumento, DataRecebimento, DataVencimento, NumeroPedido, ValorDocumento`
      );
    }

    const mapped = rows
      .map((r, idx) => ({
        _idx: idx + 1,
        NumeroDocumento: r[keyMap['numerodocumento']],
        DataDocumento: r[keyMap['datadocumento']],
        DataRecebimento: r[keyMap['datarecebimento']],
        DataVencimento: r[keyMap['datavencimento']],
        NumeroPedido: r[keyMap['numeropedido']],
        ValorDocumento: r[keyMap['valordocumento']]
      }))
      .filter(r => r.NumeroDocumento !== undefined && r.NumeroDocumento !== null && String(r.NumeroDocumento).trim() !== '');

    if (!mapped.length) throw new Error('Nenhuma linha válida encontrada (NumeroDocumento vazio).');
    return mapped;
  }

  // ======= CORRIGIDO AQUI TAMBÉM =======
  // raw:true ajuda a manter Date como Date (evita "1/18/26")
  async function parseFile(file) {
    const name = (file.name || '').toLowerCase();
    if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
      const ab = await file.arrayBuffer();
      const wb = XLSX.read(ab, { type: 'array', cellDates: true });
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];

      const rows = XLSX.utils.sheet_to_json(ws, {
        defval: '',
        raw: true,      // <-- importante
        cellDates: true
      });

      return validateAndMapRows(rows);
    }

    if (name.endsWith('.csv')) {
      const text = await file.text();
      return validateAndMapRows(parseCsv(text));
    }

    throw new Error('Formato não suportado. Use .xlsx ou .csv');
  }
  // ======= FIM DA CORREÇÃO =======

  function parseCsv(text) {
    const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').filter(l => l.trim() !== '');
    if (!lines.length) return [];
    const delim = detectDelimiter(lines[0]);
    const rows = lines.map(line => splitCsvLine(line, delim));
    const header = rows[0].map(h => h.trim());
    const out = [];
    for (let i = 1; i < rows.length; i++) {
      const obj = {};
      for (let c = 0; c < header.length; c++) obj[header[c]] = rows[i][c] ?? '';
      out.push(obj);
    }
    return out;
  }

  function detectDelimiter(line) {
    const semi = (line.match(/;/g) || []).length;
    const comma = (line.match(/,/g) || []).length;
    const tab = (line.match(/\t/g) || []).length;
    if (tab >= semi && tab >= comma) return '\t';
    if (semi >= comma) return ';';
    return ',';
  }

  function splitCsvLine(line, delim) {
    const out = [];
    let cur = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i+1] === '"') { cur += '"'; i++; }
        else inQuotes = !inQuotes;
      } else if (ch === delim && !inQuotes) {
        out.push(cur);
        cur = '';
      } else {
        cur += ch;
      }
    }
    out.push(cur);
    return out;
  }

  function hasFormFields() {
    return !!(document.getElementById('txtNumeroDocumento') &&
              document.getElementById('txtDataDocumento') &&
              document.getElementById('txtDataRecebimento') &&
              document.getElementById('txtDataVencimento') &&
              document.getElementById('txtNumeroPedido') &&
              document.getElementById('lnkAddNota'));
  }

  async function runAutomation(rows) {
    STATE.running = true;
    STATE.stopRequested = false;
    STATE.rows = rows;
    STATE.currentIndex = 0;

    ensurePanel();
    log(`Iniciando importação: ${rows.length} linha(s).`);
    setStatus(`Rodando... 0/${rows.length}`);

    if (!hasFormFields()) {
      STATE.running = false;
      throw new Error('Não encontrei os campos do formulário nesta página. Abra a tela do formulário de notas (onde existe o botão "Salvar (F8)") e tente novamente.');
    }

    for (let i = 0; i < rows.length; i++) {
      if (STATE.stopRequested) {
        log('Importação parada pelo usuário.');
        break;
      }

      const r = rows[i];
      STATE.currentIndex = i;

      try {
        removeReadonly();

        const numero = String(r.NumeroDocumento ?? '').trim();
        const dtDoc = fmtDateBR(r.DataDocumento);
        const dtRec = fmtDateBR(r.DataRecebimento);
        const dtVenc = fmtDateBR(r.DataVencimento);
        const pedido = String(r.NumeroPedido ?? '').trim();
        const valor = fmtMoneyBR(r.ValorDocumento);

        log(`Linha ${i+1}: preenchendo Documento=${numero} Pedido=${pedido} Valor=${valor}`);

        await setField('txtNumeroDocumento', numero);
        await setField('txtNumeroPedido', pedido);
        await setField('txtDataDocumento', dtDoc);
        await setField('txtDataRecebimento', dtRec);
        await setField('txtDataVencimento', dtVenc);

        const elValor = document.getElementById('txtValorDocumento');
        if (elValor) {
          elValor.removeAttribute('readonly');
          elValor.readOnly = false;
          elValor.focus();
          setNativeValue(elValor, valor);
          dispatchAll(elValor);
        }

        const btnAdd = document.getElementById('lnkAddNota');
        if (!btnAdd) throw new Error('Botão lnkAddNota não encontrado.');
        btnAdd.click();

        setStatus(`Rodando... ${i+1}/${rows.length}`);

        await sleep(STATE.delayMs);

        const err = findVisibleError();
        if (err) {
          log(`⚠️ Possível erro após enviar linha ${i+1}: ${err}`);
        } else {
          log(`✅ Linha ${i+1} enviada (clique em lnkAddNota).`);
        }
      } catch (e) {
        log(`❌ Erro na linha ${i+1}: ${e && e.message ? e.message : String(e)}`);
        STATE.stopRequested = true;
        break;
      }
    }

    STATE.running = false;
    setStatus('Finalizado.');
    log('Finalizado.');
  }

  function findVisibleError() {
    const selectors = [
      '.validation-summary-errors',
      '.field-validation-error',
      '.alert.alert-danger',
      '.alert-danger',
      '.text-danger',
      '.toast-error',
      '.ui-state-error'
    ];
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el) {
        const txt = (el.innerText || '').trim();
        if (txt) return txt.slice(0, 200);
      }
    }
    return null;
  }

  function ensureImportButton() {
    if (document.getElementById('egate-btn-importar')) return true;

    const found = findActionButtonsContainer();
    if (!found) return false;

    const { salvar, container } = found;

    const btn = salvar.cloneNode(true);
    btn.id = 'egate-btn-importar';

    btn.removeAttribute('onclick');
    btn.removeAttribute('href');
    btn.removeAttribute('accesskey');
    btn.removeAttribute('data-ajax');
    btn.removeAttribute('data-ajax-method');

    if (btn.tagName.toLowerCase() === 'input') {
      btn.value = 'Importar';
    } else {
      const icon = btn.querySelector('i, span.glyphicon, span.fa');
      btn.innerHTML = '';
      if (icon) btn.appendChild(icon);
      btn.appendChild(document.createTextNode(' Importar'));
    }

    btn.addEventListener('click', async (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      try {
        ensurePanel();
        if (STATE.running) {
          log('Já está rodando. Clique em "Parar" no painel.');
          return;
        }
        const file = await pickFile();
        const rows = await parseFile(file);
        log(`Arquivo carregado: ${file.name} (${rows.length} linha(s) válidas).`);
        await runAutomation(rows);
      } catch (e) {
        ensurePanel();
        log(`Erro: ${e && e.message ? e.message : String(e)}`);
        setStatus('Erro.');
      }
    }, true);

    if (salvar.nextSibling) container.insertBefore(btn, salvar.nextSibling);
    else container.appendChild(btn);

    log('Botão "Importar" inserido na tela.');
    return true;
  }

  function pickFile() {
    return new Promise((resolve, reject) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.xlsx,.xls,.csv';
      input.style.display = 'none';
      document.body.appendChild(input);

      input.addEventListener('change', () => {
        const file = input.files && input.files[0];
        input.remove();
        if (!file) reject(new Error('Nenhum arquivo selecionado.'));
        else resolve(file);
      });

      input.click();
    });
  }

  function boot() {
    ensureImportButton();
    const obs = new MutationObserver(() => {
      if (ensureImportButton()) obs.disconnect();
    });
    obs.observe(document.documentElement, { childList: true, subtree: true });

    setInterval(() => {
      if (STATE.running) removeReadonly();
    }, 800);
  }

  boot();
})();
