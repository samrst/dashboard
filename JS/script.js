let rawData = [];
let filteredData = [];
let charts = {};

const fileInput = document.getElementById('file-input');
const filters = ['unidade', 'curso', 'modalidade'];

// Event Listeners
fileInput.addEventListener('change', handleFile);
filters.forEach(f => {
    const el = document.getElementById(`filter-${f}`);
    if (el) el.addEventListener('change', applyFilters);
});

// Helper: lê coluna com vários nomes possíveis (compatível com Python e Excel original)
function pick(d, keys, fallback = 0) {
    for (const k of keys) {
        if (d[k] !== undefined && d[k] !== null && d[k] !== '') {
            const n = parseInt(d[k]);
            return isNaN(n) ? fallback : n;
        }
    }
    return fallback;
}

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('loading-overlay').querySelector('p').textContent = 'Sincronizando dados...';

    const reader = new FileReader();
    reader.onload = function (ev) {
        try {
            const data = new Uint8Array(ev.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames.find(n => n.includes('Worksheet')) || workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(sheet);

            const cleanJson = json.filter(d => {
                const dr = String(d['Departamento Regional (DR)'] || '').toLowerCase();
                const unidade = d['Unidade operacional (Escola)'];
                return dr !== 'total' && unidade !== undefined && unidade !== null && unidade !== '';
            });

            rawData = cleanJson.map(d => {
                return {
                    ...d,

                    '_prova_objetiva': pick(
                        d,
                        ['Prova Objetiva', '_prova_objetiva']
                    ),

                    '_total_provas_geradas': pick(
                        d,
                        ['Total de Provas Geradas', '_total_provas_geradas']
                    ),

                    '_total_alunos': pick(
                        d,
                        ['Alunos Homologados', 'Total', '_total_alunos']
                    ),

                    '_provas_aplicadas': pick(
                        d,
                        ['Provas Aplicadas', 'Status de Geração', '_provas_aplicadas']
                    ),

                    '_provas_nao_aplicadas': pick(
                        d,
                        ['Provas Não Aplicadas', 'Provas Aptas', '_provas_nao_aplicadas']
                    ),

                    '_tabulacao_feita': pick(
                        d,
                        ['Tabulação Feita', '_tabulacao_feita']
                    ),

                    '_tabulacao_pendente': pick(
                        d,
                        ['Tabulação Pendente', '_tabulacao_pendente']
                    )
                }; });

            initDashboard();
        } catch (err) {
            console.error(err);
            alert("Erro ao ler a planilha: " + err.message);
            document.getElementById('loading-overlay').querySelector('p').textContent = 'Aguardando planilha...';
        }
    };
    reader.readAsArrayBuffer(file);
}

function initDashboard() {
    document.getElementById('loading-overlay').style.display = 'none';
    document.getElementById('main-content').style.display = 'block';
    populateFilters();
    applyFilters();
}

function populateFilters() {
    filters.forEach(f => {
        const key = f === 'unidade' ? 'Unidade operacional (Escola)'
                  : f === 'curso' ? 'Curso'
                  : f === 'modalidade' ? 'Modalidade'
                  : `_${f}`;
        const values = [...new Set(rawData.map(d => d[key]))].filter(Boolean).sort();
        const sel = document.getElementById(`filter-${f}`);
        if (sel) {
            const currentVal = sel.value;
            sel.innerHTML = `<option value="ALL">Todas</option>`;
            values.forEach(v => sel.add(new Option(v, v)));
            if (values.includes(currentVal)) sel.value = currentVal;
        }
    });
}

function applyFilters() {
    const vals = {};
    filters.forEach(f => {
        const el = document.getElementById(`filter-${f}`);
        vals[f] = el ? el.value : 'ALL';
    });

    filteredData = rawData.filter(d => {
        return (vals.unidade === 'ALL'    || d['Unidade operacional (Escola)'] === vals.unidade) &&
               (vals.curso === 'ALL'      || d['Curso'] === vals.curso) &&
               (vals.modalidade === 'ALL' || d['Modalidade'] === vals.modalidade) //&&
              // (vals.turno === 'ALL'      || d['_turno'] === vals.turno) &&
               //(vals.mes === 'ALL'        || d['_mes'] === vals.mes)
               ;
    });

    updateKPIs();
    updateCharts();
    updateTable();
}

function updateKPIs() {
    const sum = key => filteredData.reduce((a, b) => a + (b[key] || 0), 0);
    const total     = sum('_total_provas_geradas');
    const objetivo  = sum('_prova_objetiva');
    const homologados = sum('_total_alunos');
    const aplicadas = sum('_provas_aplicadas');
    const feitas    = sum('_tabulacao_feita');
    const pendentes = sum('_tabulacao_pendente');

    document.getElementById('kpi-total').textContent     = total.toLocaleString('pt-BR');
    document.getElementById('kpi-objetivo').textContent  = objetivo.toLocaleString('pt-BR');
    document.getElementById('kpi-homologados').textContent = homologados.toLocaleString('pt-BR');
    document.getElementById('kpi-aplicadas').textContent = aplicadas.toLocaleString('pt-BR');
    document.getElementById('kpi-feitas').textContent    = feitas.toLocaleString('pt-BR');
    document.getElementById('kpi-pendentes').textContent = pendentes.toLocaleString('pt-BR');

    document.getElementById('kpi-homologados-pct').textContent =
    (objetivo
        ? Math.round(homologados / objetivo * 100)
        : 0
    ) + '% da Prova Objetiva';
    document.getElementById('kpi-total-pct').textContent =
    (homologados
        ? Math.round(total / homologados * 100)
        : 0
    ) + '% dos Homologados';
    document.getElementById('kpi-aplicadas-pct').textContent = (total ? Math.round(aplicadas / total * 100) : 0) + '% das Agendadas';
    document.getElementById('kpi-feitas-pct').textContent    = (aplicadas ? Math.round(feitas / aplicadas * 100) : 0) + '% das Aplicadas';
    document.getElementById('kpi-pendentes-pct').textContent = (aplicadas ? Math.round(pendentes / aplicadas * 100) : 0) + '% das Aplicadas';
}

function updateCharts() {
    // 1. Alunos por Escola
    const dataAlunos = {};

    const mostrarPorCurso =
        document.getElementById('filter-unidade').value !== 'ALL';

    filteredData.forEach(d => {

        const k = mostrarPorCurso
            ? (d['Curso'] || 'S/I')
            : (d['Unidade operacional (Escola)'] || 'S/I');

        if (!dataAlunos[k]) {
                dataAlunos[k] = {
                    v1: 0,
                    v2: 0
                };
            }

            dataAlunos[k].v1 += d._total_alunos;
            dataAlunos[k].v2 += d._prova_objetiva;
        });

        renderBar(
            'chart-alunos-escola',
            Object.keys(dataAlunos),
            [
                {
                    label: 'Alunos Homologados',
                    data: Object.values(dataAlunos).map(v => v.v1),
                    color: '#005599'
                },
                {
                    label: 'Não Homologados',
                    data: Object.values(dataAlunos).map(v => v.v2),
                    color: '#94a3b8'
                }
            ],
            'x',
            true
        );

    // 2. Homologados vs Agendados — BARRAS agrupadas por Curso
    const groupKey = 'Curso';
    const groupAgg = {};
    filteredData.forEach(d => {
        const k = (d[groupKey] || 'S/I').toString();
        if (!groupAgg[k]) groupAgg[k] = { hom: 0, ag: 0 };
        groupAgg[k].hom += d._total_alunos;
        groupAgg[k].ag  += d._provas_aplicadas;
    });
    // Ordena desc por Homologados para leitura imediata
    const groupEntries = Object.entries(groupAgg).sort((a, b) => b[1].hom - a[1].hom);
    const labels       = groupEntries.map(([k]) => k);
    const homData      = groupEntries.map(([, v]) => v.hom);
    const agData       = groupEntries.map(([, v]) => v.ag);

    // Layout responsivo + scrollspy: poucos itens → vertical; muitos → horizontal com viewport fixo e rolagem interna
    const wrapper  = document.getElementById('chart-homologacao-wrapper');
    const viewport = document.getElementById('chart-homologacao-viewport');
    const hint     = document.getElementById('chart-homologacao-hint');
    const indicator= document.getElementById('chart-homologacao-indicator');
    const btnTop   = document.getElementById('chart-homologacao-top');
    const horizontal = labels.length > 8;
    const ROW_H = 30; // px por categoria — garante boa leitura

    if (horizontal) {
        viewport.classList.add('is-scrollable');
        wrapper.style.height = (labels.length * ROW_H + 40) + 'px';
        hint.textContent = `${labels.length} cursos — role para ver todos.`;
    } else {
        viewport.classList.remove('is-scrollable');
        wrapper.style.height = '350px';
        hint.textContent = '';
        indicator.style.display = 'none';
        btnTop.style.display = 'none';
    }
    renderBar('chart-homologacao', labels,
        [{ label: 'Homologados', data: homData, color: '#003366' },
         { label: 'Agendados',   data: agData,  color: '#00aaff' }],
        horizontal ? 'y' : 'x', false);

    if (horizontal) setupScrollspy(viewport, wrapper, indicator, btnTop, labels.length, ROW_H);


    // 3. Aplicação por Escola
    const dataApp = {};

        filteredData.forEach(d => {

            const k =
                document.getElementById('filter-unidade').value === 'ALL'
                    ? (d['Unidade operacional (Escola)'] || 'Outros')
                    : (d['Curso'] || 'Outros');

            if (!dataApp[k]) {
                dataApp[k] = {
                    v1: 0,
                    v2: 0
                };
            }

            dataApp[k].v1 += d._provas_aplicadas;
            dataApp[k].v2 += d._provas_nao_aplicadas;
        });

renderBar(
    'chart-aplicacao',
    Object.keys(dataApp),
    [
        {
            label: 'Agendadas',
            data: Object.values(dataApp).map(v => v.v1),
            color: '#005599'
        },
        {
            label: 'Não Agendadas',
            data: Object.values(dataApp).map(v => v.v2),
            color: '#94a3b8'
        }
    ],
    'x',
    true
);
    // 4. Tabulação por Curso
    const dataTab = {};
    filteredData.forEach(d => {
        const k = d['Curso'] || "Outros";
        if (!dataTab[k]) dataTab[k] = { v1: 0, v2: 0 };
        dataTab[k].v1 += d._tabulacao_feita;
        dataTab[k].v2 += d._tabulacao_pendente;
    });
    renderBar('chart-tabulacao', Object.keys(dataTab).map(k => k.substring(0, 28)),
        [{ label: 'Feita',    data: Object.values(dataTab).map(v => v.v1), color: '#003DA5' },
         { label: 'Pendente', data: Object.values(dataTab).map(v => v.v2), color: '#00aaff' }], 'y', true);

    // Pizzas
    renderPie('chart-aplicacao-pizza',
        ['Aplicadas', 'Não Aplicadas'],
        [filteredData.reduce((a, b) => a + b._provas_aplicadas, 0),
         filteredData.reduce((a, b) => a + b._provas_nao_aplicadas, 0)],
        ['#005599', '#94a3b8']);

    renderPie('chart-tabulacao-pizza',
        ['Feita', 'Pendente'],
        [filteredData.reduce((a, b) => a + b._tabulacao_feita, 0),
         filteredData.reduce((a, b) => a + b._tabulacao_pendente, 0)],
        ['#003DA5', '#00aaff']);
}

function renderBar(id, labels, datasets, axis = 'x', stacked = false) {
    const ctx = document.getElementById(id);
    if (!ctx) return;
    if (charts[id]) charts[id].destroy();

    charts[id] = new Chart(ctx, {
        type: 'bar',
        data: { labels, datasets: datasets.map(d => ({ label: d.label, data: d.data, backgroundColor: d.color })) },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: axis,
            maxBarThickness: 50,
            categoryPercentage: 0.8,
            barPercentage: 0.9,
            scales: {
                x: { stacked: stacked },
                y: { stacked: stacked, beginAtZero: true }
            },
            plugins: { legend: { position: 'bottom' } }
        }
    });
}

function renderPie(id, labels, data, colors) {
    const ctx = document.getElementById(id);
    if (!ctx) return;
    if (charts[id]) charts[id].destroy();

    const total = data.reduce((a, b) => a + b, 0);

    charts[id] = new Chart(ctx, {
        type: 'doughnut',
        data: { labels, datasets: [{ data, backgroundColor: colors, borderWidth: 2, borderColor: '#fff' }] },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '60%',
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: (ctx) => {
                            const v = ctx.parsed;
                            const pct = total ? ((v / total) * 100).toFixed(1) : 0;
                            return `${ctx.label}: ${v.toLocaleString('pt-BR')} (${pct}%)`;
                        }
                    }
                }
            }
        }
    });
}

function updateTable() {
    const body = document.getElementById('table-body');
    if (!body) return;
    body.innerHTML = filteredData.map(d => `
        <tr>
            <td>${d['Unidade operacional (Escola)'] || '-'}</td>
            <td>${(d['Curso'] || '-').substring(0, 25)}</td>
            <td>${d['_total_alunos']}</td>
            <td>${d['_provas_aplicadas']}</td>
            <td>${d['_provas_nao_aplicadas']}</td>
            <td>${d['_tabulacao_feita']}</td>
            <td>${d['_tabulacao_pendente']}</td>
        </tr>
    `).join('');

    const viewport  = document.getElementById('table-viewport');
    const indicator = document.getElementById('table-indicator');
    const btnTop    = document.getElementById('table-top');
    const hint      = document.getElementById('table-hint');
    const total     = filteredData.length;
    if (viewport && total > 0) {
        // altura média de linha (~42px) — recalibrada após o paint
        const firstRow = body.querySelector('tr');
        const rowH = firstRow ? firstRow.getBoundingClientRect().height || 42 : 42;
        hint.textContent = `${total} registro${total === 1 ? '' : 's'} — role para ver todos.`;
        setupScrollspy(viewport, body, indicator, btnTop, total, rowH);
    } else if (indicator) {
        indicator.style.display = 'none';
        btnTop.style.display = 'none';
        hint.textContent = '';
    }
}

/* ============ Scrollspy do gráfico de homologação ============ */
function setupScrollspy(viewport, inner, indicator, btnTop, total, rowH) {
    indicator.style.display = 'block';
    const update = () => {
        const top = viewport.scrollTop;
        const visible = viewport.clientHeight;
        const first = Math.max(1, Math.floor(top / rowH) + 1);
        const last  = Math.min(total, Math.ceil((top + visible) / rowH));
        indicator.textContent = `${first}–${last} de ${total}`;
        btnTop.style.display = top > 80 ? 'flex' : 'none';
    };
    // remove handlers antigos para não acumular ao re-renderizar
    if (viewport._scrollspyHandler) viewport.removeEventListener('scroll', viewport._scrollspyHandler);
    viewport._scrollspyHandler = update;
    viewport.addEventListener('scroll', update, { passive: true });
    btnTop.onclick = () => viewport.scrollTo({ top: 0, behavior: 'smooth' });
    // estado inicial
    viewport.scrollTop = 0;
    update();
}
