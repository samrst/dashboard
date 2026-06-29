let rawData = [], filteredData = [], charts = {};
const $ = id => document.getElementById(id);
const filters = ['unidade', 'curso', 'modalidade'];
const keysMap = { unidade: 'Unidade operacional (Escola)', curso: 'Curso', modalidade: 'Modalidade' };

// Event Listeners
if ($('file-input')) $('file-input').onchange = handleFile;
filters.forEach(f => { if ($(`filter-${f}`)) $(`filter-${f}`).onchange = applyFilters; });

const pick = (d, keys) => {
    for (const k of keys) if (d[k] != null && d[k] !== '') return parseInt(d[k]) || 0;
    return 0;
};

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    if ($('loading-overlay')) $('loading-overlay').querySelector('p').textContent = 'Sincronizando...';

    const reader = new FileReader();
    reader.onload = ev => {
        try {
            const workbook = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames.find(n => n.includes('Worksheet')) || workbook.SheetNames[0]];
            rawData = XLSX.utils.sheet_to_json(sheet).filter(d => 
                String(d['Departamento Regional (DR)'] || '').toLowerCase() !== 'total' && d['Unidade operacional (Escola)']
            ).map(d => ({
                ...d,
                _prova_objetiva: pick(d, ['Prova Objetiva', '_prova_objetiva']),
                _total_provas_geradas: pick(d, ['Total de Provas Geradas', '_total_provas_geradas']),
                _total_alunos: pick(d, ['Alunos Homologados', 'Total', '_total_alunos']),
                _provas_aplicadas: pick(d, ['Provas Aplicadas', 'Status de Geração', '_provas_aplicadas']),
                _provas_nao_aplicadas: pick(d, ['Provas Não Aplicadas', 'Provas Aptas', '_provas_nao_aplicadas']),
                _tabulacao_feita: pick(d, ['Tabulação Feita', '_tabulacao_feita']),
                _tabulacao_pendente: pick(d, ['Tabulação Pendente', '_tabulacao_pendente'])
            }));
            initDashboard();
        } catch (err) { alert("Erro: " + err.message); }
        finally { e.target.value = ''; } // Reset do input para permitir re-seleção do mesmo arquivo
    };
    reader.readAsArrayBuffer(file);
}

function initDashboard() {
    if ($('loading-overlay')) $('loading-overlay').style.display = 'none';
    if ($('main-content')) $('main-content').style.display = 'block';
    populateFilters();
    applyFilters();
}

function populateFilters() {
    filters.forEach(f => {
        const key = keysMap[f], sel = $(`filter-${f}`);
        if (!sel) return;
        const vals = [...new Set(rawData.map(d => d[key]))].filter(Boolean).sort();
        const current = sel.value;
        sel.innerHTML = '<option value="ALL">Todas</option>' + vals.map(v => `<option value="${v}">${v}</option>`).join('');
        if (vals.includes(current)) sel.value = current;
    });
}

function applyFilters() {
    const vals = Object.fromEntries(filters.map(f => [f, $(`filter-${f}`)?.value || 'ALL']));
    filteredData = rawData.filter(d => filters.every(f => vals[f] === 'ALL' || d[keysMap[f]] === vals[f]));
    updateKPIs(); updateCharts(); updateTable();
}

function updateKPIs() {
    const sum = k => filteredData.reduce((a, b) => a + (Number(b[k]) || 0), 0);
    const kpis = { total: '_total_provas_geradas', objetivo: '_prova_objetiva', homologados: '_total_alunos', aplicadas: '_provas_aplicadas', feitas: '_tabulacao_feita', pendentes: '_tabulacao_pendente' };
    const vals = Object.fromEntries(Object.entries(kpis).map(([k, v]) => [k, sum(v)]));
    
    Object.entries(vals).forEach(([k, v]) => { if ($(`kpi-${k}`)) $(`kpi-${k}`).textContent = v.toLocaleString('pt-BR'); });

    const pcts = {
        'homologados-pct': [vals.homologados, vals.objetivo, 'da Prova Objetiva'],
        'total-pct': [vals.total, vals.homologados, 'dos Homologados'],
        'aplicadas-pct': [vals.aplicadas, vals.total, 'das Agendadas'],
        'feitas-pct': [vals.feitas, vals.aplicadas, 'das Aplicadas'],
        'pendentes-pct': [vals.pendentes, vals.aplicadas, 'das Aplicadas']
    };
    Object.entries(pcts).forEach(([id, [v, total, msg]]) => {
        if ($(`kpi-${id}`)) $(`kpi-${id}`).textContent = (total ? Math.round(v / total * 100) : 0) + '% ' + msg;
    });
}

function aggregate(key, metrics) {
    return filteredData.reduce((acc, d) => {
        const k = d[key] || 'S/I';
        if (!acc[k]) acc[k] = Object.fromEntries(metrics.map(m => [m, 0]));
        metrics.forEach(m => acc[k][m] += Number(d[m]) || 0);
        return acc;
    }, {});
}

function updateCharts() {
    if (typeof Chart === 'undefined') return;
    const isAll = ($('filter-unidade')?.value === 'ALL');
    
    // 1. Alunos
    const dAlunos = aggregate(isAll ? keysMap.unidade : keysMap.curso, ['_total_alunos', '_prova_objetiva']);
    renderBar('chart-alunos-escola', Object.keys(dAlunos), [
        { label: 'Homologados', data: Object.values(dAlunos).map(v => v._total_alunos), color: '#005599' },
        { label: 'Prova Objetiva', data: Object.values(dAlunos).map(v => v._prova_objetiva), color: '#94a3b8' }
    ], 'x', true);

    // 2. Homologação vs Agendados
    const dHom = aggregate(keysMap.curso, ['_total_alunos', '_provas_aplicadas']);
    const sortedHom = Object.entries(dHom).sort((a, b) => b[1]._total_alunos - a[1]._total_alunos);
    const labels = sortedHom.map(e => e[0]), isH = labels.length > 8;
    
    const vp = $('chart-homologacao-viewport'), wr = $('chart-homologacao-wrapper');
    if (vp && wr) {
        vp.classList.toggle('is-scrollable', isH);
        wr.style.height = isH ? (labels.length * 30 + 40) + 'px' : '350px';
        if ($('chart-homologacao-hint')) $('chart-homologacao-hint').textContent = isH ? `${labels.length} cursos — role.` : '';
    }
    renderBar('chart-homologacao', labels, [
        { label: 'Homologados', data: sortedHom.map(e => e[1]._total_alunos), color: '#003366' },
        { label: 'Agendados', data: sortedHom.map(e => e[1]._provas_aplicadas), color: '#00aaff' }
    ], isH ? 'y' : 'x');
    if (isH) setupScrollspy(vp, wr, $('chart-homologacao-indicator'), $('chart-homologacao-top'), labels.length, 30);

    // 3 & 4 & Pizzas
    const dApp = aggregate(isAll ? keysMap.unidade : keysMap.curso, ['_provas_aplicadas', '_provas_nao_aplicadas']);
    renderBar('chart-aplicacao', Object.keys(dApp), [
        { label: 'Agendadas', data: Object.values(dApp).map(v => v._provas_aplicadas), color: '#005599' },
        { label: 'Não Agendadas', data: Object.values(dApp).map(v => v._provas_nao_aplicadas), color: '#94a3b8' }
    ], 'x', true);

    const dTab = aggregate(keysMap.curso, ['_tabulacao_feita', '_tabulacao_pendente']);
    renderBar('chart-tabulacao', Object.keys(dTab).map(k => k.slice(0, 28)), [
        { label: 'Feita', data: Object.values(dTab).map(v => v._tabulacao_feita), color: '#003DA5' },
        { label: 'Pendente', data: Object.values(dTab).map(v => v._tabulacao_pendente), color: '#00aaff' }
    ], 'y', true);

    const sumMetric = m => filteredData.reduce((a, b) => a + (Number(b[m]) || 0), 0);
    renderPie('chart-aplicacao-pizza', ['Aplicadas', 'Não'], [sumMetric('_provas_aplicadas'), sumMetric('_provas_nao_aplicadas')], ['#005599', '#94a3b8']);
    renderPie('chart-tabulacao-pizza', ['Feita', 'Pendente'], [sumMetric('_tabulacao_feita'), sumMetric('_tabulacao_pendente')], ['#003DA5', '#00aaff']);

    // 5. Percentual (Melhorado com cores dinâmicas e ordenação)
    const dPct = aggregate(isAll ? keysMap.unidade : keysMap.curso, ['_prova_objetiva', '_total_alunos']);
    const resPct = Object.entries(dPct).map(([n, v]) => {
        const p = v._prova_objetiva ? (v._total_alunos / v._prova_objetiva * 100) : 0;
        return { n, p: Number(p.toFixed(1)) };
    }).sort((a, b) => a.p - b.p);

    renderBar('chart-percentual-pratica', resPct.map(r => r.n), [{
        label: '% da Prova Objetiva',
        data: resPct.map(r => r.p),
        backgroundColor: resPct.map(r => r.p < 50 ? '#ef4444' : (r.p < 80 ? '#f59e0b' : '#003DA5'))
    }], 'y', false, true);
}

function renderBar(id, labels, datasets, axis = 'x', stacked = false, isPercent = false) {
    const ctx = $(id); if (!ctx || typeof Chart === 'undefined') return;
    if (charts[id]) charts[id].destroy();

    const options = { 
        responsive: true, maintainAspectRatio: false, indexAxis: axis, 
        scales: { 
            x: { stacked, beginAtZero: true }, 
            y: { stacked, beginAtZero: true } 
        }, 
        plugins: { 
            legend: { position: 'bottom' },
            tooltip: { 
                callbacks: { 
                    label: c => `${c.dataset.label}: ${isPercent ? c.raw + '%' : c.raw.toLocaleString('pt-BR')}` 
                } 
            }
        } 
    };

    if (isPercent) {
        const scale = axis === 'x' ? 'y' : 'x';
        options.scales[scale].max = 100;
        options.plugins.annotation = {
            annotations: {
                line1: {
                    type: 'line',
                    [axis === 'x' ? 'yMin' : 'xMin']: 50,
                    [axis === 'x' ? 'yMax' : 'xMax']: 50,
                    borderColor: 'rgba(0, 0, 0, 0.5)',
                    borderWidth: 2,
                    borderDash: [6, 6],
                    label: {
                        display: true,
                        content: 'Meta 50%',
                        position: 'end'
                    }
                }
            }
        };
    }

    charts[id] = new Chart(ctx, {
        type: 'bar', 
        data: { labels, datasets: datasets.map(d => ({ ...d, backgroundColor: d.backgroundColor || d.color })) },
        options: options
    });
}

function renderPie(id, labels, data, colors) {
    const ctx = $(id); if (!ctx || typeof Chart === 'undefined') return;
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart(ctx, {
        type: 'doughnut', data: { labels, datasets: [{ data, backgroundColor: colors }] },
        options: { 
            responsive: true, maintainAspectRatio: false, cutout: '60%', 
            plugins: { 
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: c => `${c.label}: ${c.raw.toLocaleString('pt-BR')}`
                    }
                }
            } 
        }
    });
}

function updateTable() {
    const body = $('table-body'); if (!body) return;
    body.innerHTML = filteredData.map(d => `<tr><td>${d[keysMap.unidade] || '-'}</td><td>${(d[keysMap.curso] || '-').slice(0, 25)}</td><td>${d._total_alunos}</td><td>${d._provas_aplicadas}</td><td>${d._provas_nao_aplicadas}</td><td>${d._tabulacao_feita}</td><td>${d._tabulacao_pendente}</td></tr>`).join('');
    const vp = $('table-viewport'), ind = $('table-indicator'), hint = $('table-hint');
    if (vp && filteredData.length) {
        if (hint) hint.textContent = `${filteredData.length} registros — role.`;
        setupScrollspy(vp, body, ind, $('table-top'), filteredData.length, 42);
    } else if (ind) ind.style.display = 'none';
}

function setupScrollspy(vp, inner, ind, btn, total, rowH) {
    if (!vp || !ind) return;
    ind.style.display = 'block';
    const up = () => {
        const t = vp.scrollTop, v = vp.clientHeight;
        ind.textContent = `${Math.floor(t / rowH) + 1}–${Math.min(total, Math.ceil((t + v) / rowH))} de ${total}`;
        if (btn) btn.style.display = t > 80 ? 'flex' : 'none';
    };
    if (vp._h) vp.removeEventListener('scroll', vp._h);
    vp._h = up; vp.addEventListener('scroll', up, { passive: true });
    if (btn) btn.onclick = () => vp.scrollTo({ top: 0, behavior: 'smooth' });
    up();
}
