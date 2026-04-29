import { sbSaveWorkbook, sbLoadWorkbook, sbSaveMetadata, sbLoadMetadata } from './supabase-service.js';

document.addEventListener('DOMContentLoaded', () => {
    // Check Auth
    if (localStorage.getItem('isLoggedIn') !== 'true') {
        window.location.href = 'index.html';
        return;
    }

    const username = localStorage.getItem('username') || 'Admin';
    const greetingEl = document.getElementById('userGreeting');
    if (greetingEl) greetingEl.textContent = `Olá, ${username}`;

    // Injetar CSS para Checkbox Premium
    const style = document.createElement('style');
    style.textContent = `
        .pay-checkbox {
            appearance: none;
            width: 20px;
            height: 20px;
            border: 2px solid #cbd5e1;
            border-radius: 6px;
            cursor: pointer;
            outline: none;
            transition: all 0.2s ease;
            position: relative;
            display: inline-block;
            vertical-align: middle;
            background: white;
        }
        .pay-checkbox:checked {
            background-color: #22c55e;
            border-color: #22c55e;
            box-shadow: 0 0 8px rgba(34, 197, 94, 0.4);
        }
        .pay-checkbox:checked::after {
            content: '✓';
            position: absolute;
            color: white;
            font-size: 14px;
            font-weight: bold;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
        }
        .pay-checkbox:hover {
            border-color: #22c55e;
        }
        .selected-cell {
            background-color: rgba(34, 197, 94, 0.15) !important;
            outline: 2px solid #22c55e !important;
            outline-offset: -2px;
        }
    `;
    document.head.appendChild(style);

    let currentWorkbook = null;
    let currentCategory = 'Planilha';

    const excelTabs      = document.getElementById('excelTabs');
    const tableContainer = document.querySelector('.table-responsive');
    const btnLogout      = document.getElementById('btnLogout');
    const filterStatus   = document.getElementById('filterStatus');

    if (filterStatus) filterStatus.style.display = 'none';

    // ── Sidebar Category Navigation ─────────────────────────────
    const navPlanilha = document.getElementById('navPlanilha');
    const navAguaDesinsetizacao = document.getElementById('navAguaDesinsetizacao');
    const navAguaDesinsetizacaoEmail = document.getElementById('navAguaDesinsetizacaoEmail');
    const navAgua = document.getElementById('navAgua');
    const navAguaEmail = document.getElementById('navAguaEmail');
    const navBoletoMaos = document.getElementById('navBoletoMaos');
    const navNotaFiscal = document.getElementById('navNotaFiscal');
    const navFinanceiro = document.getElementById('navFinanceiro');
    const navExcluidos = document.getElementById('navExcluidos');

    function updateActiveNav(activeEl) {
        document.querySelectorAll('.sidebar-link').forEach(el => el.classList.remove('active'));
        if (activeEl) activeEl.classList.add('active');
    }

    const navItems = [
        { el: navPlanilha, cat: 'Planilha' },
        { el: navAguaDesinsetizacao, cat: 'agua e desinsetizacao' },
        { el: navAguaDesinsetizacaoEmail, cat: 'agua e desinsetizacao email' },
        { el: navAgua, cat: 'agua' },
        { el: navAguaEmail, cat: 'agua email' },
        { el: navBoletoMaos, cat: 'boleto em maos' },
        { el: navNotaFiscal, cat: 'Nota Fiscal' },
        { el: navFinanceiro, cat: 'Financeiro' },
        { el: navExcluidos, cat: 'Excluídos' }
    ];

    navItems.forEach(item => {
        if (item.el) {
            item.el.addEventListener('click', (e) => {
                e.preventDefault();
                currentCategory = item.cat;
                updateActiveNav(item.el);
                if (currentWorkbook) buildNav(currentWorkbook);
            });
        }
    });

    if (navPlanilha) {
        navPlanilha.addEventListener('click', () => {
            const submenu = document.getElementById('planilhaSubmenu');
            const arrow = navPlanilha.querySelector('.dropdown-arrow');
            if (submenu) {
                const isHidden = submenu.style.display === 'none';
                submenu.style.display = isHidden ? 'flex' : 'none';
                if (arrow) {
                    arrow.style.transform = isHidden ? 'rotate(180deg)' : 'rotate(0deg)';
                }
            }
        });
    }

    // ── Novo Condomínio Modal Logic ────────────────────────────────
    const btnNewCondo = document.getElementById('btnNewCondo');
    const newCondoModal = document.getElementById('newCondoModal');
    const closeNewCondoModal = document.getElementById('closeNewCondoModal');
    const newCondoForm = document.getElementById('newCondoForm');

    if (btnNewCondo && newCondoModal) {
        btnNewCondo.addEventListener('click', (e) => {
            e.preventDefault();
            newCondoModal.style.display = 'flex';
        });
    }

    if (closeNewCondoModal && newCondoModal) {
        closeNewCondoModal.addEventListener('click', () => {
            newCondoModal.style.display = 'none';
        });
    }

    if (newCondoForm) {
        newCondoForm.addEventListener('submit', (e) => {
            e.preventDefault();
            
            const name = document.getElementById('condoName').value.trim();
            const cnpj = document.getElementById('condoCNPJ').value.trim();
            const rua = document.getElementById('condoRua').value.trim();
            const bairro = document.getElementById('condoBairro').value.trim();
            const cep = document.getElementById('condoCEP').value.trim();
            const cidade = document.getElementById('condoCidade').value.trim();
            
            if (!name || !currentWorkbook) return;

            const letter = prompt("Informe a Letra do Grupo para este condomínio (ex: A, B, C):", "A").toUpperCase().trim();
            if (!letter || letter.length > 2) {
                alert("Letra inválida!");
                return;
            }

            const sheetName = `${name} (${letter})`;
            
            if (currentWorkbook.SheetNames.includes(sheetName)) {
                alert("Já existe uma planilha com este nome!");
                return;
            }

            let ws = {
                "!ref": "A1:H40",
                "A1": { v: name },
                "A2": { v: "CNPJ: " + cnpj },
                "A3": { v: `${rua}, ${bairro}, ${cep}, ${cidade}` },
                "A5": { v: "Mês" },
                "B5": { v: "Ano" },
                "C5": { v: "Parcela" },
                "D5": { v: "VALOR DA NOTA" },
                "E5": { v: "M.O" },
                "F5": { v: "M.E" },
                "G5": { v: "VALOR DO INSS" },
                "H5": { v: "À RECEBER" }
            };

            for (let r = 5; r < 35; r++) {
                let row_num = r + 1;
                ws[`E${row_num}`] = { f: `IF(D${row_num}="","",D${row_num}*0.8)` };
                ws[`F${row_num}`] = { f: `IF(D${row_num}="","",D${row_num}*0.2)` };
                ws[`G${row_num}`] = { f: `IF(E${row_num}="","",E${row_num}*0.11)` };
                ws[`H${row_num}`] = { f: `IF(D${row_num}="","",D${row_num}-G${row_num})` };
            }

            currentWorkbook.SheetNames.push(sheetName);
            currentWorkbook.Sheets[sheetName] = ws;

            saveWorkbookToDB(currentWorkbook).then(() => {
                alert("Condomínio criado com sucesso!");
                newCondoModal.style.display = 'none';
                newCondoForm.reset();
                buildNav(currentWorkbook);
            });
        });
    }
    const DB_NAME = 'CristianaExcelDB';
    const DB_VERSION = 1;
    const STORE_NAME = 'workbooks';
    const KEY_NAME = 'current';

    let sheetFormatting = {}; 
    let currentSheetName = ''; 
    let serverCategories = {};
    const userEmail = localStorage.getItem('userEmail');

    async function loadCategoriesFromServer() {
        if (!userEmail) return;
        const meta = await sbLoadMetadata(userEmail);
        serverCategories = meta.categories || {};
        for (let key in serverCategories) {
            localStorage.setItem('condo_cat_' + key, JSON.stringify(serverCategories[key]));
        }
    }

    async function saveCategoriesToServer() {
        if (!userEmail) return;
        await sbSaveMetadata(userEmail, { categories: serverCategories });
    }

    async function saveWorkbookToDB(workbook) {
        if (!workbook || !userEmail) return;

        const table = document.getElementById('spreadsheetTable');
        if (table && currentSheetName) {
            const domStates = [];
            table.querySelectorAll('td').forEach(td => {
                const r = parseInt(td.dataset.r);
                const c = parseInt(td.dataset.c);
                if (!isNaN(r) && !isNaN(c)) {
                    domStates.push({
                        r, c,
                        text: td.textContent,
                        align: td.style.textAlign,
                        weight: td.style.fontWeight,
                        size: td.style.fontSize,
                        display: td.style.display,
                        colspan: td.getAttribute('colspan'),
                        rowspan: td.getAttribute('rowspan')
                    });
                }
            });
            sheetFormatting[currentSheetName] = domStates;
        }

        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        
        // Salva arquivo no Storage e metadados no DB
        await sbSaveWorkbook(userEmail, buffer);
        await sbSaveMetadata(userEmail, { formatting: sheetFormatting });
    }

    async function loadWorkbookFromDB() {
        if (!userEmail) return null;

        // 1. Carrega formatações do DB
        const meta = await sbLoadMetadata(userEmail);
        sheetFormatting = meta.formatting || {};

        // 2. Carrega a planilha do Storage
        const buf = await sbLoadWorkbook(userEmail);
        if (buf) {
            console.log('✅ Planilha carregada da nuvem Supabase.');
            return buf;
        }

        // 3. Se não tiver nada na nuvem ainda, fallback para o arquivo local padrão
        console.log('ℹ️ Sem arquivo na nuvem, carregando planilha padrão...');
        try {
            const res = await fetch('Planilha%20Cris.xlsx');
            if (res.ok) return await res.arrayBuffer();
        } catch (e) {
            console.error("Erro ao carregar planilha padrão:", e);
        }

        return null;
    }

    let autoSaveTimeout = null;
    function triggerAutoSave() {
        clearTimeout(autoSaveTimeout);
        autoSaveTimeout = setTimeout(async () => {
            await saveWorkbookToDB(currentWorkbook);
            console.log('Planilha e formatações salvas na nuvem Supabase.');
        }, 1500);
    }

    // ─────────────────────────────────────────────────────────────
    // HELPERS
    // ─────────────────────────────────────────────────────────────

    function getSheetRange(ws) {
        if (!ws) return null;
        if (ws['!ref']) return ws['!ref'];
        let mnR = 0, mxR = 0, mnC = 0, mxC = 0, found = false;
        for (let k in ws) {
            if (k.startsWith('!')) continue;
            try {
                const c = XLSX.utils.decode_cell(k);
                if (!found) { mnR = mxR = c.r; mnC = mxC = c.c; found = true; }
                else {
                    if (c.r < mnR) mnR = c.r; if (c.r > mxR) mxR = c.r;
                    if (c.c < mnC) mnC = c.c; if (c.c > mxC) mxC = c.c;
                }
            } catch (e) {}
        }
        if (!found) return null;
        return XLSX.utils.encode_range({ s: { r: mnR, c: mnC }, e: { r: mxR, c: mxC } });
    }

    function cellVal(ws, r, c) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell) return '';
        return String(cell.w !== undefined ? cell.w : (cell.v !== undefined ? cell.v : ''));
    }

    // ─────────────────────────────────────────────────────────────
    // AUTO-LOAD
    // ─────────────────────────────────────────────────────────────

    function autoLoadExcel() {
        tableContainer.innerHTML = `
            <div style="text-align:center;padding:60px 20px;color:#94a3b8;">
                <i class="fa-solid fa-spinner fa-spin" style="font-size:32px;margin-bottom:16px;display:block;"></i>
                Carregando planilha, aguarde…
            </div>`;

        loadCategoriesFromServer().finally(() => {
            loadWorkbookFromDB()
                .then(buf => {
                    if (buf) {
                        console.log('✅ Planilha carregada do servidor.');
                        currentWorkbook = XLSX.read(new Uint8Array(buf), { type: 'array' });
                        buildNav(currentWorkbook);
                    } else {
                        tableContainer.innerHTML = `
                            <p style="text-align:center;padding:20px;color:#f87171;">
                                Não foi possível carregar <strong>Planilha Cris.xlsx</strong>.<br>
                                Use <strong>Importar Excel</strong> para carregar manualmente.
                            </p>`;
                    }
                })
                .catch(err => {
                    console.error('Erro crítico ao carregar planilha:', err);
                    tableContainer.innerHTML = `
                        <p style="text-align:center;padding:20px;color:#f87171;">
                            Erro ao carregar planilha. Verifique se o servidor está rodando.<br>
                            <small style="color:#94a3b8;">${err.message || ''}</small>
                        </p>`;
                });
        });
    }

    // ─────────────────────────────────────────────────────────────
    // BUILD NAV  — Letter-based two-level navigation
    //
    // Sheet name patterns discovered:
    //   Index sheets  : "CONDOMÍNIO A"  /  "CONDOMÍNIOS D"  (ends with " X" or " LS")
    //   Leaf  sheets  : "Planilha 1 (A)"  — name ends with  (LETTER)
    //                   OR any other non-index, non-MENU sheet
    //
    // Strategy:
    //   1. Parse all sheet names to build letterToIndex and letterToLeaves maps.
    //   2. Show letter buttons.
    //   3. Clicking a letter shows:
    //        a. The index sheet table (if it exists and has data), with every cell
    //           whose text exactly matches a sheet name made into a clickable button.
    //        b. Below (or instead), a card grid of all leaf sheets for that letter.
    //   4. Clicking a leaf opens the individual sheet in spreadsheet view.
    // ─────────────────────────────────────────────────────────────

    function buildFinanceiroView(workbook) {
        excelTabs.innerHTML = '';
        tableContainer.innerHTML = `
            <div style="text-align:center;padding:30px;color:#94a3b8;">
                <i class="fa-solid fa-spinner fa-spin" style="font-size:24px;margin-bottom:12px;display:block;"></i>
                Calculando lançamentos financeiros…
            </div>`;

        const financeData = {};
        const monthMap = {
            'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4, 'maio': 5, 'junho': 6,
            'julho': 7, 'agosto': 8, 'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
        };

        const sheetToCondo = workbook._sheetToCondo || {};

        workbook.SheetNames.forEach(name => {
            const up = name.toUpperCase().trim();
            if (up === 'MENU' || up.startsWith('CONDOMÍNIO')) return;

            const ws = workbook.Sheets[name];
            if (!ws) return;

            const ref = getSheetRange(ws);
            if (!ref) return;

            const range = XLSX.utils.decode_range(ref);
            const maxRow = Math.min(range.e.r, 1500);
            const maxCol = 18;

            let colHeaderRow = -1;
            for (let r = range.s.r; r <= Math.min(range.s.r + 12, maxRow); r++) {
                for (let c = 0; c <= maxCol; c++) {
                    const v = cellVal(ws, r, c).trim().toLowerCase();
                    if (v === 'mês' || v === 'mes') { colHeaderRow = r; break; }
                }
                if (colHeaderRow >= 0) break;
            }

            if (colHeaderRow < 0) return;

            const tables = [];

            for (let c = 0; c <= maxCol; c++) {
                const v = cellVal(ws, colHeaderRow, c).trim().toLowerCase();
                if (v === 'mês' || v === 'mes') {
                    let tableObj = { mesCol: c, anoCol: c+1, valorNotaCol: -1, aReceberCol: -1, pagoCol: -1 };
                    
                    for (let c2 = c + 1; c2 <= maxCol; c2++) {
                        const v2 = cellVal(ws, colHeaderRow, c2).trim().toLowerCase();
                        if (v2 === 'mês' || v2 === 'mes') break;
                        
                        if (v2 === 'pago') {
                            tableObj.pagoCol = c2;
                        } else if (v2 === 'valor da nota' || v2 === 'valor nota' || v2 === 'vl nota' || v2.includes('valor da nota')) {
                            tableObj.valorNotaCol = c2;
                        } else if (v2 === 'à receber' || v2 === 'a receber' || v2 === 'receber' || v2.includes('receber')) {
                            tableObj.aReceberCol = c2;
                        } else if (v2 === 'ano') {
                            tableObj.anoCol = c2;
                        }
                    }
                    
                    if (tableObj.valorNotaCol === -1) tableObj.valorNotaCol = c + 3;
                    if (tableObj.aReceberCol === -1) tableObj.aReceberCol = c + 7;
                    if (tableObj.pagoCol === -1) tableObj.pagoCol = c + 8;
                    
                    tables.push(tableObj);
                }
            }

            const parseVal = (str) => {
                if (!str) return 0;
                let clean = str.replace(/[^\d.,-]/g, '').trim();
                
                if (clean.includes('.') && clean.includes(',')) {
                    clean = clean.replace(/\./g, '').replace(',', '.');
                } else if (clean.includes(',') && !clean.includes('.')) {
                    clean = clean.replace(',', '.');
                } else if (clean.includes('.') && !clean.includes(',')) {
                    const parts = clean.split('.');
                    if (parts[parts.length - 1].length > 2) {
                        clean = clean.replace(/\./g, '');
                    }
                }
                
                let num = parseFloat(clean);
                return isNaN(num) ? 0 : num;
            };

            for (let r = colHeaderRow + 1; r <= maxRow; r++) {
                tables.forEach(t => {
                    const mes = cellVal(ws, r, t.mesCol).trim();
                    const ano = cellVal(ws, r, t.anoCol).trim();

                    if (!mes || !ano) return;

                    const valorNota = parseVal(cellVal(ws, r, t.valorNotaCol).trim());
                    const p = cellVal(ws, r, t.pagoCol).trim().toUpperCase();
                    const isPaid = (p === 'PG' || p === 'TRUE');

                    let aReceber = 0;
                    if (isPaid) {
                        aReceber = parseVal(cellVal(ws, r, t.aReceberCol).trim());
                    }

                    if (valorNota === 0 && aReceber === 0) return;

                    const key = `${mes} / ${ano}`;
                    if (!financeData[key]) {
                        financeData[key] = { entrada: 0, faturamento: 0, items: [], mes, ano };
                    }

                    financeData[key].faturamento += valorNota;
                    
                    if (isPaid) {
                        financeData[key].entrada += aReceber;
                    }

                    financeData[key].items.push({
                        condominio: sheetToCondo[name] || name,
                        valorNota,
                        aReceber: isPaid ? aReceber : 0,
                        isPaid
                    });
                });
            }


        });

        let html = `
            <div style="padding: 20px; background: #0f172a; border-radius: 12px; margin: 10px;">
                <h2 style="color: white; font-size: 18px; font-weight: 600; margin-bottom: 20px; display: flex; align-items: center; gap: 10px;">
                    <i class="fa-solid fa-file-invoice-dollar" style="color: #22c55e;"></i> Controle Financeiro — Entradas
                </h2>
                
                <div style="display: flex; gap: 16px; margin-bottom: 25px; flex-wrap: wrap;">
                    <div style="background: #1e293b; padding: 20px; border-radius: 12px; border: 1px solid rgba(34, 197, 94, 0.2); flex: 1; min-width: 200px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);">
                        <span style="color: #94a3b8; font-size: 13px; font-weight: 500;">FATURAMENTO TOTAL (NOTAS)</span>
                        <div id="totalFaturamento" style="color: white; font-size: 24px; font-weight: 700; margin-top: 8px;">R$ 0,00</div>
                    </div>
                    <div style="background: #1e293b; padding: 20px; border-radius: 12px; border: 1px solid rgba(34, 197, 94, 0.4); flex: 1; min-width: 200px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); background: rgba(34, 197, 94, 0.05);">
                        <span style="color: #22c55e; font-size: 13px; font-weight: 600;">ENTRADA EFETIVA (PAGO)</span>
                        <div id="totalEntrada" style="color: #22c55e; font-size: 24px; font-weight: 700; margin-top: 8px;">R$ 0,00</div>
                    </div>
                </div>

                <div class="table-responsive" style="border-radius: 8px; overflow: hidden; border: 1px solid rgba(255,255,255,0.05);">
                    <table style="width: 100%; border-collapse: collapse; font-size: 14px; text-align: left;">
                        <thead>
                            <tr style="background: #1e293b; color: #cbd5e1; border-bottom: 1px solid #334155;">
                                <th style="padding: 14px 16px; font-weight: 600;">Mês/Ano</th>
                                <th style="padding: 14px 16px; font-weight: 600;">Faturamento (Notas)</th>
                                <th style="padding: 14px 16px; font-weight: 600;">Entrada (À Receber)</th>
                                <th style="padding: 14px 16px; font-weight: 600; text-align: center;">Ações</th>
                            </tr>
                        </thead>
                        <tbody id="financeiroBody" style="color: #e2e8f0; background: #0f172a;">
        `;

        let globalFaturamento = 0;
        let globalEntrada = 0;

        const sortedKeys = Object.keys(financeData).sort((a, b) => {
            const dataA = financeData[a];
            const dataB = financeData[b];
            const yearDiff = parseInt(dataA.ano) - parseInt(dataB.ano);
            if (yearDiff !== 0) return yearDiff;
            
            const m1 = monthMap[dataA.mes.toLowerCase().trim()] || 0;
            const m2 = monthMap[dataB.mes.toLowerCase().trim()] || 0;
            return m1 - m2;
        });

        if (sortedKeys.length === 0) {
            html += `
                <tr>
                    <td colspan="4" style="padding: 30px; text-align: center; color: #94a3b8;">
                        Nenhum lançamento marcado como <strong>PAGO (PG)</strong> encontrado.
                    </td>
                </tr>
            `;
        } else {
            sortedKeys.forEach(key => {
                const data = financeData[key];
                globalFaturamento += data.faturamento;
                globalEntrada += data.entrada;

                html += `
                    <tr style="border-bottom: 1px solid rgba(255,255,255,0.03); transition: background 0.15s;" onmouseover="this.style.background='rgba(255,255,255,0.02)'" onmouseout="this.style.background='transparent'">
                        <td style="padding: 14px 16px; font-weight: 500;">${key}</td>
                        <td style="padding: 14px 16px;">${data.faturamento.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
                        <td style="padding: 14px 16px; color: #22c55e; font-weight: 600;">${data.entrada.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
                        <td style="padding: 14px 16px; text-align: center;">
                            <button class="view-finance-details" data-key="${encodeURIComponent(key)}" style="background: rgba(99, 102, 241, 0.1); border: 1px solid rgba(99, 102, 241, 0.2); color: #a5b4fc; padding: 6px 12px; border-radius: 6px; font-size: 12px; font-weight: 500; cursor: pointer; transition: all 0.2s;">
                                Detalhes
                            </button>
                        </td>
                    </tr>
                    <tr id="details-${encodeURIComponent(key)}" style="display: none; background: rgba(15, 23, 42, 0.5);">
                        <td colspan="4" style="padding: 15px 25px;">
                            <div style="background: #1e293b; border-radius: 8px; padding: 15px; border: 1px solid rgba(255,255,255,0.05);">
                                <h4 style="margin-top: 0; color: white; font-size: 14px; margin-bottom: 10px;">Lançamentos de ${key}</h4>
                                <table style="width:100%; border-collapse: collapse; font-size: 12px;">
                                    <thead>
                                        <tr style="color:#94a3b8; border-bottom: 1px solid #334155; text-align: left;">
                                            <th style="padding: 8px;">Condomínio</th>
                                            <th style="padding: 8px;">Faturamento</th>
                                            <th style="padding: 8px;">Entrada</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        ${data.items.map(item => `
                                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.03); opacity: ${item.isPaid ? 1 : 0.6};">
                                                <td style="padding: 8px; color: #e2e8f0;">${item.condominio}</td>
                                                <td style="padding: 8px; color: #cbd5e1;">${item.valorNota.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}</td>
                                                <td style="padding: 8px; color: ${item.isPaid ? '#22c55e' : '#94a3b8'};">
                                                    ${item.isPaid ? item.aReceber.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : '<span style="color:#94a3b8;font-size:11px;">R$ 0,00</span>'}
                                                </td>
                                            </tr>
                                        `).join('')}
                                    </tbody>
                                </table>
                            </div>
                        </td>
                    </tr>
                `;
            });
        }

        html += `</tbody></table></div></div>`;
        tableContainer.innerHTML = html;

        const totalFatEl = document.getElementById('totalFaturamento');
        const totalEntEl = document.getElementById('totalEntrada');
        if (totalFatEl) totalFatEl.textContent = globalFaturamento.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        if (totalEntEl) totalEntEl.textContent = globalEntrada.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

        tableContainer.querySelectorAll('.view-finance-details').forEach(btn => {
            btn.addEventListener('click', () => {
                const key = decodeURIComponent(btn.dataset.key);
                const detailsRow = document.getElementById(`details-${encodeURIComponent(key)}`);
                if (detailsRow) {
                    const isHidden = detailsRow.style.display === 'none';
                    detailsRow.style.display = isHidden ? 'table-row' : 'none';
                    btn.textContent = isHidden ? 'Recolher' : 'Detalhes';
                }
            });
        });
    }

    function buildNav(workbook) {
        if (currentCategory === 'Financeiro') {
            buildFinanceiroView(workbook);
            return;
        }
        excelTabs.innerHTML = '';
        const headerCatContainer = document.getElementById('headerCategoryContainer');
        if (headerCatContainer) headerCatContainer.innerHTML = '';

        const names = workbook.SheetNames;
        const letterToIndex  = {};   // letter → index sheet name
        const letterToLeaves = {};   // letter → [leaf sheet names]
        const sheetToCondo   = {};   // sheetName → Condo Name (from A1)

        names.forEach(name => {
            const up = name.toUpperCase().trim();
            if (up === 'MENU') return;

            // Get Condo Name from A1
            const ws = workbook.Sheets[name];
            let condoName = '';
            if (ws && ws['A1']) {
                condoName = String(ws['A1'].w || ws['A1'].v || '').trim();
            }
            sheetToCondo[name] = condoName || name;

            // Index: "CONDOMÍNIO(S) X" ending with 1-2 uppercase letters
            if (/CONDOM[IÍ]N/i.test(name)) {
                const m = up.match(/\s([A-Z]{1,2})$/);
                if (m) {
                    let letter = m[1];
                    if (letter === 'LS') letter = 'S';
                    letterToIndex[letter] = name;
                    return;
                }
            }

            // Leaf: ends with  (X)  e.g. "Planilha 1 (A)"
            const lm = name.match(/\(([A-Za-z]{1,2})\)\s*$/);
            if (lm) {
                let letter = lm[ lm.length - 1 ].toUpperCase();
                if (letter === 'LS') letter = 'S';
                if (!letterToLeaves[letter]) letterToLeaves[letter] = [];
                letterToLeaves[letter].push(name);
                return;
            }

            // Everything else goes to a catch-all group keyed '#'
            if (!letterToLeaves['#']) letterToLeaves['#'] = [];
            letterToLeaves['#'].push(name);
        });

        // Filter by category
        for (let letter in letterToLeaves) {
            letterToLeaves[letter] = letterToLeaves[letter].filter(name => {
                const raw = localStorage.getItem('condo_cat_' + name);
                let cats = [];
                if (raw) {
                    try {
                        cats = JSON.parse(raw);
                        if (!Array.isArray(cats)) cats = [cats];
                    } catch (e) {
                        cats = [raw];
                    }
                }
                
                if (currentCategory === 'Excluídos') {
                    return cats.includes('excluidos') || cats.includes('Excluídos');
                } else if (currentCategory === 'Planilha') {
                    return !cats.includes('excluidos') && !cats.includes('Excluídos');
                } else {
                    // Match category (case insensitive or exact)
                    const match = cats.some(c => c.toLowerCase().trim() === currentCategory.toLowerCase().trim());
                    return match && !cats.includes('excluidos') && !cats.includes('Excluídos');
                }
            });
            if (letterToLeaves[letter].length === 0) {
                delete letterToLeaves[letter];
            }
        }
        for (let letter in letterToIndex) {
            if (!letterToLeaves[letter]) {
                delete letterToIndex[letter];
            }
        }

        // Store sheetToCondo globally or pass it along
        workbook._sheetToCondo = sheetToCondo;

        // Collect all letters from both maps (except '#')
        const allLetters = new Set([
            ...Object.keys(letterToIndex),
            ...Object.keys(letterToLeaves).filter(k => k !== '#')
        ]);
        const sortedLetters = [...allLetters].sort();

        if (sortedLetters.length === 0) {
            buildFallbackDropdown(workbook);
            return;
        }

        // ── Build UI ─────────────────────────────────────────────

        const breadcrumb = makeBreadcrumb();
        const letterBar  = makeLetterBar(sortedLetters, workbook, letterToIndex, letterToLeaves, breadcrumb);

        excelTabs.style.cssText = 'display:flex; flex-direction:column; padding:0;';
        excelTabs.appendChild(breadcrumb);
        excelTabs.appendChild(letterBar);

        // Auto-open first letter
        letterBar.querySelector('.letter-btn')?.click();
    }

    function makeBreadcrumb() {
        const bc = document.createElement('div');
        bc.id = 'navBreadcrumb';
        bc.style.cssText = `
            display:flex; align-items:center; gap:6px;
            padding:10px 20px;
            background:rgba(15,23,42,0.6);
            border-bottom:1px solid rgba(255,255,255,0.08);
            font-size:13px; color:#94a3b8; flex-shrink:0; flex-wrap:wrap;
        `;
        return bc;
    }

    function makeLetterBar(letters, workbook, letterToIndex, letterToLeaves, breadcrumb) {
        const bar = document.createElement('div');
        bar.id = 'letterBar';
        bar.style.cssText = `
            display:flex; flex-wrap:wrap; gap:6px;
            padding:12px 20px;
            background:rgba(15,23,42,0.4);
            border-bottom:1px solid rgba(255,255,255,0.06);
            align-items:center;
        `;

        const lbl = document.createElement('span');
        lbl.textContent = 'Grupo:';
        lbl.style.cssText = 'color:#64748b; font-size:13px; margin-right:4px;';
        bar.appendChild(lbl);

        letters.forEach(letter => {
            const btn = document.createElement('button');
            btn.textContent = letter;
            btn.dataset.letter = letter;
            btn.className = 'letter-btn';
            setLetterBtnStyle(btn, false);

            btn.addEventListener('click', () => {
                bar.querySelectorAll('.letter-btn').forEach(b => setLetterBtnStyle(b, false));
                setLetterBtnStyle(btn, true);
                showLetterGroup(workbook, letter, letterToIndex, letterToLeaves, breadcrumb, bar);
            });

            bar.appendChild(btn);
        });

        return bar;
    }

    function setLetterBtnStyle(btn, active) {
        btn.style.cssText = `
            min-width:36px; height:36px; padding:0 8px; border-radius:8px;
            border:1px solid ${active ? '#6366f1' : 'rgba(99,102,241,0.3)'};
            background:${active ? '#6366f1' : 'rgba(99,102,241,0.1)'};
            color:${active ? 'white' : '#a5b4fc'};
            font-weight:700; font-size:13px; cursor:pointer;
            transition:all 0.18s ease;
            box-shadow:${active ? '0 4px 12px rgba(99,102,241,0.4)' : 'none'};
            display:flex; align-items:center; justify-content:center;
        `;
        if (active) btn.classList.add('active-letter');
        else btn.classList.remove('active-letter');
    }

    // ─────────────────────────────────────────────────────────────
    // LEVEL 2 — Show all condominiums for a letter
    // ─────────────────────────────────────────────────────────────

    function showLetterGroup(workbook, letter, letterToIndex, letterToLeaves, breadcrumb, letterBar) {
        const headerCatContainer = document.getElementById('headerCategoryContainer');
        if (headerCatContainer) headerCatContainer.innerHTML = '';
        // Update breadcrumb
        updateBreadcrumb(breadcrumb, letter, null, () => {
            // clicking letter in breadcrumb reloads the group (same place)
        }, letterBar, workbook, letterToIndex, letterToLeaves);

        const indexSheetName = letterToIndex[letter];
        const leafSheets     = letterToLeaves[letter] || [];
        const ws             = indexSheetName ? workbook.Sheets[indexSheetName] : null;
        const ref            = ws ? getSheetRange(ws) : null;
        const sheetNameSet   = new Set(workbook.SheetNames);
        const sheetToCondo   = workbook._sheetToCondo || {};

        let html = '';

        // ── Part 1: index sheet table (if exists and non-empty) ──
        if (ws && ref) {
            const range  = XLSX.utils.decode_range(ref);
            const maxRow = Math.min(range.e.r, 2000);
            const maxCol = range.e.c;

            html += `
                <div style="padding:12px 16px 4px; background:#1e293b; color:#94a3b8; font-size:12px; text-transform:uppercase; letter-spacing:0.5px; font-weight:600;">
                    📋 Índice — ${indexSheetName}
                </div>`;
            html += `<table class="nav-table">`;

            // Header
            html += `<thead><tr>`;
            for (let c = 0; c <= maxCol; c++) {
                html += `<th>${cellVal(ws, range.s.r, c)}</th>`;
            }
            html += `</tr></thead><tbody>`;

            // Data
            for (let r = range.s.r + 1; r <= maxRow; r++) {
                let hasData = false;
                for (let c = 0; c <= maxCol && !hasData; c++) {
                    if (ws[XLSX.utils.encode_cell({ r, c })]) hasData = true;
                }
                if (!hasData) continue;

                html += `<tr>`;
                for (let c = 0; c <= maxCol; c++) {
                    const v = cellVal(ws, r, c);
                    const t = v.trim();
                    if (t && sheetNameSet.has(t)) {
                        const displayName = sheetToCondo[t] || t;
                        html += `<td>
                            <div style="position:relative; display:flex; align-items:center; gap:6px;">
                                <button class="cond-link" data-sheet="${encodeURIComponent(t)}" style="flex-grow:1;">
                                    <i class="fa-solid fa-building" style="font-size:11px;margin-right:6px;opacity:.7;"></i>${displayName}
                                    <i class="fa-solid fa-arrow-right" style="font-size:10px;margin-left:6px;opacity:.5;"></i>
                                </button>
                                <button class="delete-condo-btn" data-sheet="${encodeURIComponent(t)}" style="background:transparent; border:none; color:#ef4444; cursor:pointer; font-size:12px; padding:4px;" title="Excluir Condomínio">
                                    <i class="fa-solid fa-trash-can"></i>
                                </button>
                            </div>
                        </td>`;
                    } else {
                        html += `<td>${v}</td>`;
                    }
                }
                html += `</tr>`;
            }
            html += `</tbody></table>`;
        }

        // ── Part 2: leaf sheet cards (sheets ending in "(LETTER)") ──
        if (leafSheets.length > 0) {
            html += `
                <div style="padding:16px 16px 4px; background:#0f172a; color:#94a3b8; font-size:12px; text-transform:uppercase; letter-spacing:0.5px; font-weight:600;">
                    🏢 Condomínios — Grupo ${letter} (${leafSheets.length})
                </div>
                <div style="display:flex; flex-wrap:wrap; gap:10px; padding:14px 16px; background:#0f172a;">`;

            leafSheets.forEach(name => {
                const displayName = sheetToCondo[name] || name;
                html += `
                <div style="position:relative; display:inline-block;">
                    <button class="cond-card" data-sheet="${encodeURIComponent(name)}">
                        <i class="fa-solid fa-building" style="font-size:20px;margin-bottom:6px;color:#818cf8;"></i>
                        <span>${displayName}</span>
                    </button>
                    <button class="delete-condo-btn" data-sheet="${encodeURIComponent(name)}" style="position:absolute; top:5px; right:5px; background:rgba(239, 68, 68, 0.1); border:none; color:#ef4444; cursor:pointer; font-size:12px; padding:4px; border-radius:50%; width:24px; height:24px; display:flex; align-items:center; justify-content:center; z-index:10; transition:all 0.2s;" title="Excluir Condomínio">
                        <i class="fa-solid fa-trash-can"></i>
                    </button>
                </div>`;
            });

            html += `</div>`;
        }

        if (!html) {
            html = `<p style="text-align:center;padding:40px;color:#94a3b8;">Nenhuma aba encontrada para o grupo <strong>${letter}</strong>.</p>`;
        }

        tableContainer.innerHTML = html;

        // Attach events
        tableContainer.querySelectorAll('.cond-link, .cond-card').forEach(btn => {
            btn.addEventListener('click', () => {
                const sheetName = decodeURIComponent(btn.dataset.sheet);
                updateBreadcrumb(breadcrumb, letter, sheetToCondo[sheetName] || sheetName, () => {
                    showLetterGroup(workbook, letter, letterToIndex, letterToLeaves, breadcrumb, letterBar);
                }, letterBar, workbook, letterToIndex, letterToLeaves);
                openLeafSheet(workbook, sheetName);
            });
        });

        tableContainer.querySelectorAll('.delete-condo-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const sheetName = decodeURIComponent(btn.dataset.sheet);
                if (confirm(`Tem certeza que deseja excluir o condomínio "${sheetName}"?`)) {
                    localStorage.setItem('condo_cat_' + sheetName, JSON.stringify(['excluidos']));
                    serverCategories[sheetName] = ['excluidos'];
                    saveCategoriesToServer();
                    alert("Condomínio movido para Excluídos!");
                    buildNav(workbook);
                }
            });
        });
    }

    // ─────────────────────────────────────────────────────────────
    // LEVEL 3 — Spreadsheet view of an individual sheet
    // ─────────────────────────────────────────────────────────────

    function openLeafSheet(workbook, sheetName) {
        currentSheetName = sheetName;
        const ws  = workbook.Sheets[sheetName];
        const ref = getSheetRange(ws);

        if (!ws || !ref) {
            tableContainer.innerHTML = `<p style="text-align:center;padding:40px;color:#94a3b8;">A aba <strong>${sheetName}</strong> está vazia.</p>`;
            return;
        }

        const range  = XLSX.utils.decode_range(ref);
        const maxRow = Math.min(range.e.r, 1500);
        const maxCol = 18; // Force exactly 18 columns (A-S) to match the layout

        // ── Detect header rows ──────────────────────────────────
        let colHeaderRow = -1;
        for (let r = range.s.r; r <= Math.min(range.s.r + 12, maxRow); r++) {
            for (let c = 0; c <= maxCol; c++) {
                const v = cellVal(ws, r, c).trim().toLowerCase();
                if (v === 'mês' || v === 'mes') { colHeaderRow = r; break; }
            }
            if (colHeaderRow >= 0) break;
        }
        
        const pagoCols = [];
        if (colHeaderRow >= 0) {
            for (let c = 0; c <= maxCol; c++) {
                const v = cellVal(ws, colHeaderRow, c).trim().toLowerCase();
                if (v === 'pago') {
                    pagoCols.push(c);
                }
            }
        }
        
        let ano2024Row = colHeaderRow > 0 ? colHeaderRow - 1 : -1;
        const infoStart = range.s.r;
        const infoEnd   = ano2024Row >= 0 ? ano2024Row - 1 : infoStart + 2;

        // ── Styles per row ──────────────────────────────────────
        function rowStyle(r) {
            if (r >= infoStart && r <= infoEnd) return 'background:#ddeeff;';
            if (r === colHeaderRow) return 'background:#1F3864; color:#fff;';
            return (r - colHeaderRow) % 2 === 0 ? 'background:#f0f4ff;' : 'background:#fff;';
        }

        function cellStyle(r, c) {
            const base = 'border:1px solid #4472C4; padding:5px 8px; white-space:nowrap; font-size:13px; min-width:80px;';
            if (r >= infoStart && r <= infoEnd) {
                return base + 'font-weight:700; color:#1F3864; background:#ddeeff;';
            }
            if (r === colHeaderRow) {
                return base + 'font-weight:700; color:#fff; background:#1F3864; text-align:center;';
            }
            if (c >= 4 && c <= 7) {
                return base + 'color:#1a6632; font-weight:600; text-align:right; background:#f0fff4;';
            }
            return base + 'text-align:' + (c <= 2 ? 'left' : 'right') + ';';
        }

        // ── Build table ─────────────────────────────────────────
        let savedCats = [];
        const rawCats = localStorage.getItem('condo_cat_' + sheetName);
        if (rawCats) {
            try {
                savedCats = JSON.parse(rawCats);
                if (!Array.isArray(savedCats)) savedCats = [savedCats];
            } catch (e) {
                savedCats = [rawCats];
            }
        }
        
        // Populate header category container
        const headerCatContainer = document.getElementById('headerCategoryContainer');
        if (headerCatContainer) {
            headerCatContainer.innerHTML = `
                <span style="color: #94a3b8; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; margin-right: 4px;">Destino:</span>
                <button class="cat-chip ${savedCats.length === 0 ? 'active' : ''}" data-cat="Planilha">
                    <i class="${savedCats.length === 0 ? 'fa-solid fa-square-check' : 'fa-regular fa-square'} check-icon" style="margin-right: 4px;"></i>
                    <i class="fa-solid fa-ban" style="margin-right: 4px;"></i> Nenhuma
                </button>
                <button class="cat-chip ${savedCats.includes('agua e desinsetizacao') ? 'active' : ''}" data-cat="agua e desinsetizacao">
                    <i class="${savedCats.includes('agua e desinsetizacao') ? 'fa-solid fa-square-check' : 'fa-regular fa-square'} check-icon" style="margin-right: 4px;"></i>
                    <i class="fa-solid fa-droplet" style="margin-right: 4px;"></i> Água/Desinset.
                </button>
                <button class="cat-chip ${savedCats.includes('agua e desinsetizacao email') ? 'active' : ''}" data-cat="agua e desinsetizacao email">
                    <i class="${savedCats.includes('agua e desinsetizacao email') ? 'fa-solid fa-square-check' : 'fa-regular fa-square'} check-icon" style="margin-right: 4px;"></i>
                    <i class="fa-solid fa-envelope" style="margin-right: 4px;"></i> Água/Desins. Email
                </button>
                <button class="cat-chip ${savedCats.includes('agua') ? 'active' : ''}" data-cat="agua">
                    <i class="${savedCats.includes('agua') ? 'fa-solid fa-square-check' : 'fa-regular fa-square'} check-icon" style="margin-right: 4px;"></i>
                    <i class="fa-solid fa-faucet" style="margin-right: 4px;"></i> Água
                </button>
                <button class="cat-chip ${savedCats.includes('agua email') ? 'active' : ''}" data-cat="agua email">
                    <i class="${savedCats.includes('agua email') ? 'fa-solid fa-square-check' : 'fa-regular fa-square'} check-icon" style="margin-right: 4px;"></i>
                    <i class="fa-solid fa-envelope-open-text" style="margin-right: 4px;"></i> Água Email
                </button>
                <button class="cat-chip ${savedCats.includes('boleto em maos') ? 'active' : ''}" data-cat="boleto em maos">
                    <i class="${savedCats.includes('boleto em maos') ? 'fa-solid fa-square-check' : 'fa-regular fa-square'} check-icon" style="margin-right: 4px;"></i>
                    <i class="fa-solid fa-hand-holding-dollar" style="margin-right: 4px;"></i> Boleto Mãos
                </button>
                <button class="cat-chip ${savedCats.includes('Nota Fiscal') ? 'active' : ''}" data-cat="Nota Fiscal">
                    <i class="${savedCats.includes('Nota Fiscal') ? 'fa-solid fa-square-check' : 'fa-regular fa-square'} check-icon" style="margin-right: 4px;"></i>
                    <i class="fa-solid fa-file-invoice-dollar" style="margin-right: 4px;"></i> Nota Fiscal
                </button>
            `;

            // Style the chips if not already styled
            if (!document.getElementById('catChipStyles')) {
                const style = document.createElement('style');
                style.id = 'catChipStyles';
                style.textContent = `
                    .cat-chip {
                        background: rgba(30, 41, 59, 0.7);
                        color: #94a3b8;
                        border: 1px solid rgba(255, 255, 255, 0.08);
                        padding: 6px 12px;
                        border-radius: 20px;
                        font-size: 12px;
                        font-weight: 500;
                        cursor: pointer;
                        transition: all 0.2s;
                    }
                    .cat-chip:hover {
                        background: rgba(255, 255, 255, 0.1);
                        color: #f1f5f9;
                    }
                    .cat-chip.active {
                        background: #6366f1;
                        color: white;
                        border-color: #6366f1;
                        box-shadow: 0 4px 12px rgba(99, 102, 241, 0.3);
                    }
                `;
                document.head.appendChild(style);
            }

            // Attach events
            headerCatContainer.querySelectorAll('.cat-chip').forEach(btn => {
                btn.addEventListener('click', () => {
                    const selectedCat = btn.getAttribute('data-cat');
                    let savedCats = [];
                    const rawCats = localStorage.getItem('condo_cat_' + sheetName);
                    if (rawCats) {
                        try {
                            savedCats = JSON.parse(rawCats);
                            if (!Array.isArray(savedCats)) savedCats = [rawCats];
                        } catch (e) {
                            savedCats = [rawCats];
                        }
                    }
                    
                    if (selectedCat === 'Planilha') {
                        savedCats = [];
                        localStorage.removeItem('condo_cat_' + sheetName);
                    } else {
                        const index = savedCats.indexOf(selectedCat);
                        if (index > -1) {
                            savedCats.splice(index, 1);
                        } else {
                            savedCats.push(selectedCat);
                        }
                        
                        if (savedCats.length === 0) {
                            localStorage.removeItem('condo_cat_' + sheetName);
                            delete serverCategories[sheetName];
                        } else {
                            localStorage.setItem('condo_cat_' + sheetName, JSON.stringify(savedCats));
                            serverCategories[sheetName] = savedCats;
                        }
                        saveCategoriesToServer();
                    }
                    
                    // Update active state and icons
                    headerCatContainer.querySelectorAll('.cat-chip').forEach(b => {
                        const cat = b.getAttribute('data-cat');
                        const icon = b.querySelector('.check-icon');
                        
                        if (cat === 'Planilha') {
                            if (savedCats.length === 0) {
                                b.classList.add('active');
                                if (icon) icon.className = 'fa-solid fa-square-check check-icon';
                            } else {
                                b.classList.remove('active');
                                if (icon) icon.className = 'fa-regular fa-square check-icon';
                            }
                        } else {
                            if (savedCats.includes(cat)) {
                                b.classList.add('active');
                                if (icon) icon.className = 'fa-solid fa-square-check check-icon';
                            } else {
                                b.classList.remove('active');
                                if (icon) icon.className = 'fa-regular fa-square check-icon';
                            }
                        }
                    });
                });
            });
        }

        let html = `<div style="overflow:auto; max-height:calc(100vh - 220px); background:white;">`;
        html += `<table style="border-collapse:collapse; table-layout:fixed; width:100%; font-family:Calibri,Arial,sans-serif;" id="spreadsheetTable">`;
        
        // Define Column Widths
        html += `<colgroup>`;
        html += `<col style="width: 40px;">`; // #
        // Tabela 1
        html += `<col style="width: 100px;">`; // A - Mês
        html += `<col style="width: 80px;">`;  // B - Ano
        html += `<col style="width: 80px;">`;  // C - Parcela
        html += `<col style="width: 120px;">`; // D - VALOR DA NOTA
        html += `<col style="width: 110px;">`; // E - M.O
        html += `<col style="width: 110px;">`; // F - M.E
        html += `<col style="width: 110px;">`; // G - INSS
        html += `<col style="width: 110px;">`; // H - A Receber
        html += `<col style="width: 60px;">`;  // I - PAGO
        // Espaço
        html += `<col style="width: 30px;">`;  // J - Espaço
        // Tabela 2
        html += `<col style="width: 100px;">`; // K - Mês
        html += `<col style="width: 80px;">`;  // L - Ano
        html += `<col style="width: 80px;">`;  // M - Parcela
        html += `<col style="width: 120px;">`; // N - VALOR DA NOTA
        html += `<col style="width: 110px;">`; // O - M.O
        html += `<col style="width: 110px;">`; // P - M.E
        html += `<col style="width: 110px;">`; // Q - INSS
        html += `<col style="width: 110px;">`; // R - A Receber
        html += `<col style="width: 60px;">`;  // S - PAGO
        html += `</colgroup>`;

        // Column letters header
        html += `<thead><tr style="background:#0f172a;">`;
        html += `<th style="padding:4px 6px;color:#64748b;font-size:11px;border:1px solid #334155;min-width:32px;position:sticky;top:0;z-index:10;">#</th>`;
        for (let c = 0; c <= maxCol; c++) {
            html += `<th style="padding:4px 8px;color:#94a3b8;font-size:11px;border:1px solid #334155;position:sticky;top:0;z-index:10;cursor:pointer;" class="col-header" data-c="${c}">${XLSX.utils.encode_col(c)}</th>`;
        }
        html += `</tr></thead><tbody>`;

        for (let r = range.s.r; r <= maxRow; r++) {
            let hasData = false;
            for (let c = 0; c <= maxCol; c++) {
                if (ws[XLSX.utils.encode_cell({r, c})]) { hasData = true; break; }
            }
            if (!hasData && r > colHeaderRow + 24) continue; // Limit empty rows after 24

            const rStyle = rowStyle(r);
            html += `<tr style="${rStyle}">`;
            html += `<td style="padding:3px 6px;color:#64748b;font-size:11px;border:1px solid #334155;text-align:center;font-weight:600;position:sticky;left:0;background:inherit;z-index:5;cursor:pointer;" class="row-header" data-r="${r}">${r + 1}</td>`;

            for (let c = 0; c <= maxCol; c++) {
                // 1. Condomínio Info (Rows infoStart to infoStart+2, Cols A-E)
                if (r >= infoStart && r <= infoStart + 2) {
                    const cs = cellStyle(r, c);
                    const v = cellVal(ws, r, c);
                    if (c === 0) {
                        html += `<td style="${cs} background:#ddeeff; font-weight:bold; text-align:left;" colspan="5">${v}</td>`;
                        continue;
                    }
                    if (c >= 1 && c <= 4) {
                        continue;
                    }
                }

                // 2. Síndico Info (Rows infoStart to infoStart+1, Cols L-M)
                if (r >= infoStart && r <= infoStart + 1) {
                    const cs = cellStyle(r, c);
                    const v = cellVal(ws, r, c);
                    if (c === 11) {
                        html += `<td style="${cs} background:#ddeeff; font-weight:bold; text-align:left;" colspan="3">${v}</td>`;
                        continue;
                    }
                    if (c >= 12 && c <= 13) {
                        continue;
                    }
                }

                // 3. Merged Cells for "Ano 2024" (Based on cell content)
                const v_ano = cellVal(ws, r, c);
                const v_next = cellVal(ws, r, c + 1);
                
                if (v_ano.trim() === 'Ano' && v_next.trim() === 'Ano 2024') {
                    const cs = cellStyle(r, c);
                    html += `<td style="${cs} background:#ddeeff;" contenteditable="true" data-r="${r}" data-c="${c}"></td>`;
                    continue;
                }
                if (v_ano.trim() === 'Ano 2024') {
                    const cs = cellStyle(r, c);
                    html += `<td style="${cs} text-align:center; font-weight:bold; background:#ddeeff;" colspan="2" contenteditable="true" data-r="${r}" data-c="${c}">Ano 2024</td>`;
                    continue;
                }
                if (c > 0 && cellVal(ws, r, c - 1).trim() === 'Ano 2024') {
                    continue;
                }

                const cellRef = XLSX.utils.encode_cell({r, c});
                const v = cellVal(ws, r, c);
                const cs = cellStyle(r, c);
                
                // Determine if cell is editable
                const isDataRow = r > colHeaderRow;
                
                // Column 9 (J) is the empty separator column
                if (c === 9) {
                    html += `<td style="${cs} width:30px; border-top:none; border-bottom:none; border-left:none; border-right:none; background:#f8fafc;" data-r="${r}" data-c="${c}" contenteditable="true"></td>`;
                } 
                // Columns 8 (I) and 19 (T) are the PAGO columns
                else if (isDataRow && pagoCols.includes(c)) {
                    const isChecked = v === 'PG' || v === 'true' || v === true;
                    html += `<td style="${cs} text-align:center; background:#fff;" data-r="${r}" data-c="${c}">
                        <input type="checkbox" class="pay-checkbox" data-r="${r}" data-c="${c}" ${isChecked ? 'checked' : ''}>
                    </td>`;
                }
                else {
                    html += `<td style="${cs} outline:none;" contenteditable="true" data-r="${r}" data-c="${c}">${v}</td>`;
                }
            }
            html += `</tr>`;
        }

        html += `</tbody></table></div>`;
        tableContainer.innerHTML = html;



        const table = document.getElementById('spreadsheetTable');
        if (table) {
            const savedFormat = sheetFormatting[sheetName];
            if (savedFormat) {
                savedFormat.forEach(s => {
                    const td = table.querySelector(`td[data-r="${s.r}"][data-c="${s.c}"]`);
                    if (td) {
                        if (s.text !== undefined && !pagoCols.includes(s.c)) td.textContent = s.text;
                        if (s.align) td.style.textAlign = s.align;
                        if (s.weight) td.style.fontWeight = s.weight;
                        if (s.size) td.style.fontSize = s.size;
                        if (s.display) td.style.display = s.display;
                        if (s.colspan) td.setAttribute('colspan', s.colspan);
                        else td.removeAttribute('colspan');
                        if (s.rowspan) td.setAttribute('rowspan', s.rowspan);
                        else td.removeAttribute('rowspan');
                    }
                });
            }
            
            table.addEventListener('input', (e) => {
                const td = e.target;
                const r = parseInt(td.dataset.r);
                const c = parseInt(td.dataset.c);
                if (isNaN(r) || isNaN(c)) return;

                const val = td.textContent.trim();
                
                // Update in-memory XLSX workbook
                const cellRef = XLSX.utils.encode_cell({r, c});
                if (!ws[cellRef]) ws[cellRef] = {};
                ws[cellRef].v = isNaN(Number(val)) ? val : Number(val);
                ws[cellRef].w = val;

                triggerAutoSave();

                const parseVal = (str) => {
                    if (!str) return 0;
                    let clean = str.replace(/[^\d.,-]/g, '').trim();
                    if (clean.includes('.') && clean.includes(',')) {
                        clean = clean.replace(/\./g, '').replace(',', '.');
                    } else if (clean.includes(',') && !clean.includes('.')) {
                        clean = clean.replace(',', '.');
                    } else if (clean.includes('.') && !clean.includes(',')) {
                        const parts = clean.split('.');
                        if (parts[parts.length - 1].length > 2) {
                            clean = clean.replace(/\./g, '');
                        }
                    }
                    let num = parseFloat(clean);
                    return isNaN(num) ? 0 : num;
                };

                const customRound = (num) => {
                    const str = num.toFixed(4);
                    const dotIdx = str.indexOf('.');
                    if (dotIdx === -1) return num;
                    
                    const thirdDecimal = parseInt(str.charAt(dotIdx + 3));
                    if (isNaN(thirdDecimal)) return num;

                    if (thirdDecimal <= 5) {
                        return Math.floor(num * 100) / 100;
                    } else {
                        return Math.ceil(num * 100) / 100;
                    }
                };

                // Recalculate Formulas for Left Section (A-I)
                if (c === 3) { // VALOR DA NOTA changed (Col D)
                    const valorNota = parseVal(val);
                    if (valorNota > 0) {
                        const mo = customRound(valorNota * 0.8);
                        const me = customRound(valorNota * 0.2);
                        const inss = customRound(mo * 0.11);
                        const receber = customRound(valorNota - inss);

                        updateCell(r, 4, mo);
                        updateCell(r, 5, me);
                        updateCell(r, 6, inss);
                        updateCell(r, 7, receber);

                        // ── Replicate for next 11 months if this is the first row ──
                        if (r === colHeaderRow + 1) {
                            replicateRows(r, valorNota, val);
                        }
                    } else {
                        updateCell(r, 4, '');
                        updateCell(r, 5, '');
                        updateCell(r, 6, '');
                        updateCell(r, 7, '');
                    }
                }

                // Recalculate Formulas for Right Section (K-S)
                if (c === 13) { // VALOR DA NOTA changed (Col N)
                    const valorNota = parseVal(val);
                    if (valorNota > 0) {
                        const mo = customRound(valorNota * 0.8);
                        const me = customRound(valorNota * 0.2);
                        const inss = customRound(mo * 0.11);
                        const receber = customRound(valorNota - inss);

                        updateCell(r, 14, mo);
                        updateCell(r, 15, me);
                        updateCell(r, 16, inss);
                        updateCell(r, 17, receber);
                    } else {
                        updateCell(r, 14, '');
                        updateCell(r, 15, '');
                        updateCell(r, 16, '');
                        updateCell(r, 17, '');
                    }
                }

                function replicateRows(startRow, valorNota, origVal) {
                    // Get base values from startRow
                    const baseMes = cellVal(ws, startRow, 0).trim();
                    const baseAno = cellVal(ws, startRow, 1).trim();
                    const baseParcela = parseInt(cellVal(ws, startRow, 2).trim()) || 1;

                    const monthsArr = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
                    
                    let monthIdx = monthsArr.indexOf(baseMes.toLowerCase());
                    let isNumericMonth = !isNaN(parseInt(baseMes));
                    let numericMonth = parseInt(baseMes);

                    let maxFill = 12;
                    let startYear = parseInt(baseAno);
                    let hasNumericYear = !isNaN(startYear);

                    for (let i = 1; i <= maxFill; i++) {
                        const targetRow = startRow + i;
                        
                        // 1. Mês
                        let newMes = baseMes;
                        if (isNumericMonth) {
                            let totalMonths = numericMonth + i - 1;
                            let wrappedMonth = (totalMonths % 12) + 1;
                            newMes = String(wrappedMonth);
                        } else if (monthIdx !== -1) {
                            let currentMonthIdx = (monthIdx + i) % 12;
                            newMes = monthsArr[currentMonthIdx];
                            if (baseMes[0] === baseMes[0].toUpperCase()) {
                                newMes = newMes.charAt(0).toUpperCase() + newMes.slice(1);
                            }
                        }
                        
                        // 2. Ano
                        let newAno = baseAno;
                        if (hasNumericYear) {
                            let yearsToAdd = 0;
                            if (isNumericMonth) {
                                yearsToAdd = Math.floor((numericMonth + i - 1) / 12);
                            } else if (monthIdx !== -1) {
                                yearsToAdd = Math.floor((monthIdx + i) / 12);
                            }
                            newAno = String(startYear + yearsToAdd);
                        }
                        
                        // 3. Parcela
                        let newParcela = String(baseParcela + i);

                        // Update Cells in memory and DOM
                        updateDataCell(targetRow, 0, newMes);
                        updateDataCell(targetRow, 1, newAno);
                        updateDataCell(targetRow, 2, newParcela);
                        updateDataCell(targetRow, 3, origVal || String(valorNota));

                        // Calculate Formulas for target row
                        const mo = customRound(valorNota * 0.8);
                        const me = customRound(valorNota * 0.2);
                        const inss = customRound(mo * 0.11);
                        const receber = customRound(valorNota - inss);

                        updateCell(targetRow, 4, mo);
                        updateCell(targetRow, 5, me);
                        updateCell(targetRow, 6, inss);
                        updateCell(targetRow, 7, receber);
                    }
                }

                function updateDataCell(r, c, val) {
                    const cellRef = XLSX.utils.encode_cell({r, c});
                    if (!ws[cellRef]) ws[cellRef] = {};
                    ws[cellRef].v = isNaN(Number(val)) ? val : Number(val);
                    ws[cellRef].w = val;
                    
                    const td = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
                    if (td) {
                        td.textContent = val;
                    }
                }
            });

            // ── Attach Checkbox Listener ───────────────────────
            table.addEventListener('change', (e) => {
                if (e.target.classList.contains('pay-checkbox')) {
                    const r = parseInt(e.target.dataset.r);
                    const c = parseInt(e.target.dataset.c);
                    if (isNaN(r) || isNaN(c)) return;

                    const cellRef = XLSX.utils.encode_cell({r, c});
                    if (!ws[cellRef]) ws[cellRef] = {};
                    ws[cellRef].v = e.target.checked ? 'PG' : '';
                    ws[cellRef].w = e.target.checked ? 'PG' : '';
                    triggerAutoSave();
                }
            });

            // ── Attach Enter Key Navigation ─────────────────────
            table.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    const td = e.target;
                    const r = parseInt(td.dataset.r);
                    const c = parseInt(td.dataset.c);
                    if (isNaN(r) || isNaN(c)) return;

                    // Find next editable cell in the same row
                    let nextTd = null;
                    for (let nextC = c + 1; nextC <= 20; nextC++) {
                        const candidate = table.querySelector(`td[data-r="${r}"][data-c="${nextC}"][contenteditable="true"]`);
                        if (candidate) {
                            nextTd = candidate;
                            break;
                        }
                    }

                    // If not found, move to first editable cell of next row
                    if (!nextTd) {
                        for (let nextR = r + 1; nextR <= maxRow; nextR++) {
                            const candidate = table.querySelector(`td[data-r="${nextR}"][data-c="0"][contenteditable="true"]`);
                            if (candidate) {
                                nextTd = candidate;
                                break;
                            }
                        }
                    }

                    if (nextTd) {
                        nextTd.focus();
                        // Select all text in the focused cell
                        const range = document.createRange();
                        range.selectNodeContents(nextTd);
                        const sel = window.getSelection();
                        sel.removeAllRanges();
                        sel.addRange(range);
                    }
                }
            });

            // ── Attach Cell Selection & Context Menu ───────────
            let isSelecting = false;
            let startCell = null;
            let endCell = null;

            // Create or reuse context menu
            let contextMenu = document.getElementById('customContextMenu');
            if (!contextMenu) {
                contextMenu = document.createElement('div');
                contextMenu.id = 'customContextMenu';
                contextMenu.style.cssText = `
                    position: absolute;
                    display: none;
                    background: #1e293b;
                    color: #f1f5f9;
                    border: 1px solid #475569;
                    border-radius: 8px;
                    box-shadow: 0 10px 15px -3px rgba(0,0,0,0.3);
                    z-index: 9999;
                    width: 180px;
                    padding: 6px 0;
                    font-family: Arial, sans-serif;
                    font-size: 13px;
                `;
                document.body.appendChild(contextMenu);
                
                document.addEventListener('click', () => {
                    contextMenu.style.display = 'none';
                });
            }

            const menuOptions = [
                { label: '📝 Centralizar', action: 'center' },
                { label: '🧹 Limpar Conteúdo', action: 'clear' },
                { label: '<b>B</b> Negrito', action: 'bold' },
                { label: '➕ Aumentar Fonte', action: 'increase-font' },
                { label: '➖ Diminuir Fonte', action: 'decrease-font' },
                { label: '🔗 Mesclar Células', action: 'merge' }
            ];

            contextMenu.innerHTML = '';
            menuOptions.forEach(opt => {
                const item = document.createElement('div');
                item.innerHTML = opt.label;
                item.style.cssText = `
                    padding: 8px 16px;
                    cursor: pointer;
                    transition: background 0.2s;
                `;
                item.addEventListener('mouseover', () => item.style.background = '#334155');
                item.addEventListener('mouseout', () => item.style.background = 'transparent');
                item.addEventListener('click', (e) => {
                    e.stopPropagation();
                    executeAction(opt.action);
                    contextMenu.style.display = 'none';
                });
                contextMenu.appendChild(item);
            });

            // ── Undo System (Ctrl+Z) ──────────────────────────
            const undoStack = [];

            function pushToUndo() {
                const wsCopy = {};
                Object.keys(ws).forEach(key => {
                    if (key.startsWith('!')) {
                        wsCopy[key] = ws[key];
                    } else {
                        wsCopy[key] = { ...ws[key] };
                    }
                });

                const domStates = [];
                table.querySelectorAll('td').forEach(td => {
                    const r = parseInt(td.dataset.r);
                    const c = parseInt(td.dataset.c);
                    if (!isNaN(r) && !isNaN(c)) {
                        domStates.push({
                            r, c,
                            text: td.textContent,
                            align: td.style.textAlign,
                            weight: td.style.fontWeight,
                            size: td.style.fontSize,
                            display: td.style.display,
                            colspan: td.getAttribute('colspan'),
                            rowspan: td.getAttribute('rowspan')
                        });
                    }
                });

                undoStack.push({ ws: wsCopy, dom: domStates });
                if (undoStack.length > 50) undoStack.shift();
            }

            function applyUndo() {
                const state = undoStack.pop();
                if (!state) return;

                Object.keys(ws).forEach(key => {
                    if (!key.startsWith('!')) delete ws[key];
                });
                Object.assign(ws, state.ws);

                state.dom.forEach(s => {
                    const td = table.querySelector(`td[data-r="${s.r}"][data-c="${s.c}"]`);
                    if (td) {
                        if (!pagoCols.includes(s.c)) td.textContent = s.text;
                        td.style.textAlign = s.align;
                        td.style.fontWeight = s.weight;
                        td.style.fontSize = s.size;
                        td.style.display = s.display;
                        if (s.colspan) td.setAttribute('colspan', s.colspan);
                        else td.removeAttribute('colspan');
                        if (s.rowspan) td.setAttribute('rowspan', s.rowspan);
                        else td.removeAttribute('rowspan');
                        if (pagoCols.includes(s.c)) {
                            const chk = td.querySelector('input[type="checkbox"]');
                            if (chk) {
                                const cellRef = XLSX.utils.encode_cell({r: s.r, c: s.c});
                                const v = ws[cellRef] ? ws[cellRef].v : '';
                                chk.checked = (v === 'PG' || v === 'true' || v === true);
                            }
                        }
                    }
                });
            }

            table.addEventListener('focusin', (e) => {
                const td = e.target.closest('td');
                if (td && td.contentEditable === 'true') {
                    pushToUndo();
                }
            });

            window.addEventListener('keydown', (e) => {
                if (e.ctrlKey && (e.key === 'z' || e.key === 'Z')) {
                    e.preventDefault();
                    applyUndo();
                }
            });

            function executeAction(action) {
                if (!startCell || !endCell) return;
                pushToUndo(); // Save state before action
                const minR = Math.min(startCell.r, endCell.r);
                const maxR = Math.max(startCell.r, endCell.r);
                const minC = Math.min(startCell.c, endCell.c);
                const maxC = Math.max(startCell.c, endCell.c);

                if (action === 'merge') {
                    const topLeftTd = table.querySelector(`td[data-r="${minR}"][data-c="${minC}"]`);
                    if (!topLeftTd) return;

                    const colspan = maxC - minC + 1;
                    const rowspan = maxR - minR + 1;

                    if (colspan > 1) topLeftTd.setAttribute('colspan', colspan);
                    if (rowspan > 1) topLeftTd.setAttribute('rowspan', rowspan);

                    // Hide the rest
                    for (let r = minR; r <= maxR; r++) {
                        for (let c = minC; c <= maxC; c++) {
                            if (r === minR && c === minC) continue;
                            const td = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
                            if (td) td.style.display = 'none';
                        }
                    }
                    triggerAutoSave();
                    return;
                }

                for (let r = minR; r <= maxR; r++) {
                    for (let c = minC; c <= maxC; c++) {
                        const td = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
                        if (!td) continue;

                        const cellRef = XLSX.utils.encode_cell({r, c});

                        if (action === 'center') {
                            td.style.textAlign = td.style.textAlign === 'center' ? 'left' : 'center';
                        } else if (action === 'clear') {
                            if (td.contentEditable === 'true') {
                                td.textContent = '';
                                if (ws[cellRef]) {
                                    ws[cellRef].v = '';
                                    ws[cellRef].w = '';
                                }
                            }
                        } else if (action === 'bold') {
                            td.style.fontWeight = td.style.fontWeight === 'bold' ? 'normal' : 'bold';
                        } else if (action === 'increase-font') {
                            const currSize = parseInt(window.getComputedStyle(td).fontSize) || 13;
                            td.style.fontSize = (currSize + 2) + 'px';
                        } else if (action === 'decrease-font') {
                            const currSize = parseInt(window.getComputedStyle(td).fontSize) || 13;
                            td.style.fontSize = Math.max(9, currSize - 2) + 'px';
                        }
                    }
                }
                triggerAutoSave();
            }

            function isCellSelected(r, c) {
                if (!startCell || !endCell) return false;
                const minR = Math.min(startCell.r, endCell.r);
                const maxR = Math.max(startCell.r, endCell.r);
                const minC = Math.min(startCell.c, endCell.c);
                const maxC = Math.max(startCell.c, endCell.c);
                return r >= minR && r <= maxR && c >= minC && c <= maxC;
            }

            function updateSelectionHighlight() {
                table.querySelectorAll('td').forEach(td => {
                    const r = parseInt(td.dataset.r);
                    const c = parseInt(td.dataset.c);
                    if (!isNaN(r) && !isNaN(c) && isCellSelected(r, c)) {
                        td.classList.add('selected-cell');
                    } else {
                        td.classList.remove('selected-cell');
                    }
                });
            }

            // Click Col Header to select entire column
            table.querySelectorAll('.col-header').forEach(th => {
                th.addEventListener('click', (e) => {
                    const c = parseInt(th.dataset.c);
                    if (isNaN(c)) return;
                    startCell = { r: 0, c };
                    endCell = { r: maxRow, c };
                    updateSelectionHighlight();
                });
            });

            // Click Row Header to select entire row
            table.querySelectorAll('.row-header').forEach(td => {
                td.addEventListener('click', (e) => {
                    const r = parseInt(td.dataset.r);
                    if (isNaN(r)) return;
                    startCell = { r, c: 0 };
                    endCell = { r, c: maxCol };
                    updateSelectionHighlight();
                });
            });

            table.addEventListener('mousedown', (e) => {
                const td = e.target.closest('td');
                if (!td) return;
                const r = parseInt(td.dataset.r);
                const c = parseInt(td.dataset.c);
                if (isNaN(r) || isNaN(c)) return;

                if (e.button === 2) { // Right click
                    if (isCellSelected(r, c)) return; // Keep existing selection
                }

                isSelecting = true;
                startCell = { r, c };
                endCell = { r, c };
                updateSelectionHighlight();
            });

            table.addEventListener('mouseover', (e) => {
                if (!isSelecting) return;
                const td = e.target.closest('td');
                if (!td) return;
                const r = parseInt(td.dataset.r);
                const c = parseInt(td.dataset.c);
                if (isNaN(r) || isNaN(c)) return;

                endCell = { r, c };
                updateSelectionHighlight();
            });

            window.addEventListener('mouseup', () => {
                isSelecting = false;
            });

            table.addEventListener('contextmenu', (e) => {
                const td = e.target.closest('td');
                if (!td) return;
                const r = parseInt(td.dataset.r);
                const c = parseInt(td.dataset.c);
                if (isNaN(r) || isNaN(c)) return;

                e.preventDefault();

                if (!isCellSelected(r, c)) {
                    startCell = { r, c };
                    endCell = { r, c };
                    updateSelectionHighlight();
                }

                contextMenu.style.left = `${e.pageX}px`;
                contextMenu.style.top = `${e.pageY}px`;
                contextMenu.style.display = 'block';
            });
        }

        function updateCell(r, c, val) {
            const cellRef = XLSX.utils.encode_cell({r, c});
            if (!ws[cellRef]) ws[cellRef] = {};
            ws[cellRef].v = val === '' ? '' : Number(val);
            ws[cellRef].w = val === '' ? '' : 'R$ ' + parseFloat(val).toLocaleString('pt-BR', {minimumFractionDigits:2});
            
            const td = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            if (td && !td.querySelector('input')) {
                td.textContent = ws[cellRef].w;
            }
        }
    }


    // ─────────────────────────────────────────────────────────────
    // BREADCRUMB
    // ─────────────────────────────────────────────────────────────

    function updateBreadcrumb(container, letter, condName, onLetterClick, letterBar, workbook, letterToIndex, letterToLeaves) {
        container.innerHTML = '';

        const home = crumb('🏠 Início', true);
        home.addEventListener('click', () => buildNav(workbook));

        const sep1 = document.createTextNode(' › ');

        const lSpan = crumb(`Grupo ${letter}`, !!condName);
        if (condName) lSpan.addEventListener('click', onLetterClick);

        container.appendChild(home);
        container.appendChild(sep1);
        container.appendChild(lSpan);

        if (condName) {
            container.appendChild(document.createTextNode(' › '));
            container.appendChild(crumb(condName, false));
        }
    }

    function crumb(text, isLink) {
        const s = document.createElement('span');
        s.textContent = text;
        s.style.cssText = isLink
            ? 'color:#818cf8; cursor:pointer; text-decoration:underline; text-underline-offset:2px;'
            : 'color:#e2e8f0; font-weight:600;';
        return s;
    }

    // ─────────────────────────────────────────────────────────────
    // FALLBACK DROPDOWN
    // ─────────────────────────────────────────────────────────────

    function buildFallbackDropdown(workbook) {
        excelTabs.style.cssText = 'display:flex; align-items:center; padding:12px 20px; gap:12px;';
        excelTabs.innerHTML = '';

        const select = document.createElement('select');
        select.style.cssText = `
            padding:8px 14px; border-radius:8px;
            background:#4f46e5; color:white; border:none;
            font-size:14px; cursor:pointer;
        `;

        workbook.SheetNames.forEach(name => {
            if (name.toUpperCase().trim() === 'MENU') return;
            const opt = document.createElement('option');
            opt.value = name; opt.textContent = name;
            select.appendChild(opt);
        });

        select.addEventListener('change', e => openLeafSheet(workbook, e.target.value));
        excelTabs.appendChild(select);
        if (select.options.length > 0) openLeafSheet(workbook, select.options[0].value);
    }

    // ─────────────────────────────────────────────────────────────
    // SEARCH
    // ─────────────────────────────────────────────────────────────



    // ─────────────────────────────────────────────────────────────
    // IMPORT / EXPORT / LOGOUT
    // ─────────────────────────────────────────────────────────────

    const btnImport   = document.getElementById('btnImport');
    const importInput = document.getElementById('importInput');

    // btnImport is now a <label for="importInput"> in the sidebar — no extra click handler needed
    // but we still need the change event on importInput

    importInput.addEventListener('change', e => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = ev => {
            try {
                currentWorkbook = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
                buildNav(currentWorkbook);
                alert('Planilha importada com sucesso!');
            } catch (err) {
                console.error(err);
                alert('Erro ao processar a planilha.');
            }
            importInput.value = '';
        };
        reader.readAsArrayBuffer(file);
    });

    document.getElementById('btnExport').addEventListener('click', () => {
        if (!currentWorkbook) { alert('Nenhuma planilha carregada!'); return; }
        XLSX.writeFile(currentWorkbook, 'Planilha_Exportada.xlsx');
    });

    btnLogout.addEventListener('click', () => {
        localStorage.removeItem('isLoggedIn');
        localStorage.removeItem('username');
        window.location.href = 'index.html';
    });

    // ─────────────────────────────────────────────────────────────
    // INIT
    // ─────────────────────────────────────────────────────────────

    autoLoadExcel();
});
