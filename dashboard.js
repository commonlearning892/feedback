fetch('feedback_stats.json')
    .then(response => response.json())
    .then(data => {
        renderDashboard(data);
    })
    .catch(error => {
        console.error('Error:', error);
        document.getElementById('content').innerHTML = '<h2>Error loading data</h2>';
    });

function renderDashboard(data) {
    initTabs();
    populateFilters(data);
    renderExecutiveSummary(data);
    renderAcademicSection(data);
    renderEnvironmentSection(data);
    renderCommunicationSection(data);
    renderInfrastructureSection(data);
    renderStrengthsSection(data);
    renderBranchComparisonSection(data);
}

function renderBranchRankings(data) {
    const table = document.getElementById('branchRankings');
    const branches = data.rankings.branches.slice(0, 10);
    let html = '<thead><tr><th>Rank</th><th>Branch</th><th>Score</th><th>Responses</th></tr></thead><tbody>';
    branches.forEach((b, i) => {
        const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `#${i + 1}`;
        const scoreClass = b[1] >= 4.5 ? 'score-excellent' : b[1] >= 3.5 ? 'score-good' : 'score-average';
        html += `<tr><td>${medal}</td><td>${b[0]}</td><td><span class="score-badge ${scoreClass}">${b[1].toFixed(2)}</span></td><td>${b[2]}</td></tr>`;
    });
    table.innerHTML = html + '</tbody>';
}

function renderOrientationRankings(data) {
    const table = document.getElementById('orientationRankings');
    const rankings = data.rankings.orientations || [];
    let html = '<thead><tr><th>Rank</th><th>Orientation</th><th>Score</th><th>Responses</th></tr></thead><tbody>';
    rankings.forEach((item, idx) => {
        const medal = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : `#${idx + 1}`;
        const scoreClass = item[1] >= 4.5 ? 'score-excellent' : item[1] >= 3.5 ? 'score-good' : 'score-average';
        html += `<tr><td>${medal}</td><td>${item[0]}</td><td><span class="score-badge ${scoreClass}">${item[1].toFixed(2)}</span></td><td>${item[2]}</td></tr>`;
    });
    table.innerHTML = html + '</tbody>';
}

function renderClassRankings(data) {
    const table = document.getElementById('classRankings');
    const rankings = data.rankings.classes || [];
    let html = '<thead><tr><th>Rank</th><th>Class</th><th>Score</th><th>Responses</th></tr></thead><tbody>';
    rankings.forEach((item, idx) => {
        const medal = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : `#${idx + 1}`;
        const scoreClass = item[1] >= 4.5 ? 'score-excellent' : item[1] >= 3.5 ? 'score-good' : 'score-average';
        html += `<tr><td>${medal}</td><td>${item[0]}</td><td><span class="score-badge ${scoreClass}">${item[1].toFixed(2)}</span></td><td>${item[2]}</td></tr>`;
    });
    table.innerHTML = html + '</tbody>';
}

function renderSubjectRankings(data) {
    const table = document.getElementById('subjectRankings');
    const subjects = (data.rankings.subjects || []).slice();
    let html = '<thead><tr><th>Rank</th><th>Subject</th><th>Score</th></tr></thead><tbody>';
    subjects.forEach((item, idx) => {
        const medal = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : `#${idx + 1}`;
        const scoreClass = item[1] >= 4.5 ? 'score-excellent' : item[1] >= 3.5 ? 'score-good' : 'score-average';
        html += `<tr><td>${medal}</td><td>${item[0]}</td><td><span class="score-badge ${scoreClass}">${item[1].toFixed(2)}</span></td></tr>`;
    });
    table.innerHTML = html + '</tbody>';
}

function renderBranchChart(data) {
    const ctx = document.getElementById('branchChart').getContext('2d');
    const branches = data.rankings.branches.slice(0, 15);
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: branches.map(b => b[0].substring(0, 20)),
            datasets: [{
                label: 'Score',
                data: branches.map(b => b[1]),
                backgroundColor: 'rgba(102, 126, 234, 0.8)'
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: { x: { beginAtZero: true, max: 5 } }
        }
    });
}

function renderOrientationChart(data) {
    const ctx = document.getElementById('orientationChart').getContext('2d');
    const counts = data.summary.orientations || {};
    const labels = Object.keys(counts);
    const values = labels.map(l => counts[l]);
    new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: [
                    'rgba(255, 99, 132, 0.8)',
                    'rgba(54, 162, 235, 0.8)',
                    'rgba(255, 206, 86, 0.8)',
                    'rgba(75, 192, 192, 0.8)'
                ]
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
    });
}

function renderSubjectChart(data) {
    const ctx = document.getElementById('subjectChart').getContext('2d');
    const subjects = data.subject_performance || {};
    new Chart(ctx, {
        type: 'radar',
        data: {
            labels: Object.keys(subjects),
            datasets: [{
                label: 'Score',
                data: Object.values(subjects).map(s => s.average),
                backgroundColor: 'rgba(118, 75, 162, 0.2)',
                borderColor: 'rgba(118, 75, 162, 1)',
                borderWidth: 3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: { r: { beginAtZero: true, max: 5 } }
        }
    });
}

function renderClassChart(data) {
    const ctx = document.getElementById('classDistChart').getContext('2d');
    const counts = data.summary.classes || {};
    const labels = Object.keys(counts);
    const values = labels.map(l => counts[l]);
    new Chart(ctx, {
        type: 'pie',
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: [
                    'rgba(76, 175, 80, 0.8)',
                    'rgba(33, 150, 243, 0.8)',
                    'rgba(255, 152, 0, 0.8)',
                    'rgba(156, 39, 176, 0.8)'
                ]
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
    });
}

function renderEnvironmentChart(data) {
    const ctx = document.getElementById('environmentChart').getContext('2d');
    const env = (data.category_performance && data.category_performance['Environment Quality']) || {};
    const metrics = Object.entries(env).slice(0, 8);
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: metrics.map(m => m[0].substring(0, 30)),
            datasets: [{
                label: 'Score',
                data: metrics.map(m => m[1].average),
                backgroundColor: 'rgba(76, 175, 80, 0.8)'
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            scales: { x: { beginAtZero: true, max: 5 } }
        }
    });
}

function renderInfrastructureChart(data) {
    const ctx = document.getElementById('infrastructureChart').getContext('2d');
    const infra = (data.category_performance && data.category_performance['Infrastructure']) || {};
    const metrics = Object.entries(infra);
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: metrics.map(m => m[0].substring(0, 30)),
            datasets: [{
                label: 'Score',
                data: metrics.map(m => m[1].average),
                backgroundColor: 'rgba(255, 152, 0, 0.8)'
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            scales: { x: { beginAtZero: true, max: 5 } }
        }
    });
}

function renderParentTeacherChart(data) {
    const ctx = document.getElementById('parentTeacherChart').getContext('2d');
    const pt = (data.category_performance && data.category_performance['Parent-Teacher Interaction']) || {};
    const metrics = Object.entries(pt).slice(0, 5);
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: metrics.map(m => m[0].substring(0, 25)),
            datasets: [{
                label: 'Score',
                data: metrics.map(m => m[1].average),
                backgroundColor: 'rgba(156, 39, 176, 0.8)'
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            scales: { x: { beginAtZero: true, max: 5 } }
        }
    });
}

function renderAdminChart(data) {
    const ctx = document.getElementById('adminChart').getContext('2d');
    const adm = (data.category_performance && data.category_performance['Administrative Support']) || {};
    const metrics = Object.entries(adm).slice(0, 6);
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: metrics.map(m => m[0].substring(0, 25)),
            datasets: [{
                label: 'Score',
                data: metrics.map(m => m[1].average),
                backgroundColor: 'rgba(233, 30, 99, 0.8)'
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            scales: { x: { beginAtZero: true, max: 5 } }
        }
    });
}

// =============== New 7-section Dashboard Renderers ===============
function initTabs() {
    const btns = Array.from(document.querySelectorAll('.tab-btn'));
    const sections = {
        exec: document.getElementById('section-exec'),
        academic: document.getElementById('section-academic'),
        env: document.getElementById('section-env'),
        comm: document.getElementById('section-comm'),
        infra: document.getElementById('section-infra'),
        strengths: document.getElementById('section-strengths'),
        branch: document.getElementById('section-branch')
    };
    btns.forEach(b => b.addEventListener('click', () => {
        btns.forEach(x => x.classList.remove('active'));
        b.classList.add('active');
        Object.values(sections).forEach(s => s.classList.remove('active'));
        const key = b.dataset.tab;
        sections[key]?.classList.add('active');
    }));
}

function populateFilters(data) {
    const addOptions = (el, items) => {
        el.innerHTML = '<option value="all">All</option>' + items.map(v => `<option value="${v}">${v}</option>`).join('');
    };
    const branches = Object.keys(data.summary.branches || {});
    const classes = Object.keys(data.summary.classes || {});
    const orientations = Object.keys(data.summary.orientations || {});
    const fb = document.getElementById('filterBranch');
    const fc = document.getElementById('filterClass');
    const fo = document.getElementById('filterOrientation');
    if (fb && fc && fo) {
        addOptions(fb, branches);
        addOptions(fc, classes);
        addOptions(fo, orientations);
    }
}

function renderExecutiveSummary(data) {
    const kpi = document.getElementById('execKpiGrid');
    if (kpi) {
        const overallPct = data.summary.overall_avg ? (data.summary.overall_avg / 5 * 100) : null;
        const yesPct = data.recommendation && data.recommendation.yes_pct != null ? data.recommendation.yes_pct : null;
        const acad = data.summary.category_scores?.Academics ?? null;
        const infra = data.summary.category_scores?.Infrastructure ?? null;
        const fmt = (v, p=false) => v==null || isNaN(v) ? '-' : (p ? `${v.toFixed(1)}%` : v.toFixed(2));
        kpi.innerHTML = `
            <div class="kpi"><div class="label">Total Responses</div><div class="value">${data.summary.total_responses}</div></div>
            <div class="kpi"><div class="label">Overall Satisfaction</div><div class="value">${fmt(overallPct, true)}</div></div>
            <div class="kpi"><div class="label">% Recommend School</div><div class="value">${fmt(yesPct || 0, true)}</div></div>
            <div class="kpi"><div class="label">Average Academic Rating</div><div class="value">${fmt(acad)}</div></div>
            <div class="kpi"><div class="label">Average Infrastructure Rating</div><div class="value">${fmt(infra)}</div></div>
        `;
    }

    const cat = data.summary.category_scores || {};
    const catCtx = document.getElementById('summaryCategoryChart')?.getContext('2d');
    if (catCtx) {
        const labels = Object.keys(cat);
        const values = labels.map(l => cat[l]);
        new Chart(catCtx, {
            type: 'bar',
            data: { labels, datasets: [{ label: 'Avg Score', data: values, backgroundColor: 'rgba(33, 150, 243, 0.8)'}] },
            options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 5 } } }
        });
    }

    const rec = data.recommendation?.distribution || {};
    const recCtx = document.getElementById('recommendationChart')?.getContext('2d');
    if (recCtx) {
        const labels = ['Yes','No','Maybe'];
        const values = labels.map(l => rec[l] || 0);
        new Chart(recCtx, {
            type: 'doughnut',
            data: { labels, datasets: [{ data: values, backgroundColor: ['#43a047','#e53935','#fb8c00'] }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
        });
    }
    // Populate side KPI for % Recommend School next to the pie
    try {
        const el = document.getElementById('recYesPctKpi');
        if (el) {
            const pct = (data.recommendation && data.recommendation.yes_pct != null) ? data.recommendation.yes_pct : 0;
            el.textContent = `${pct.toFixed(1)}%`;
        }
    } catch (e) { }
    // Populate recommendation counts KPIs (counts + %)
    try {
        const total = (rec['Yes'] || 0) + (rec['No'] || 0) + (rec['Maybe'] || 0);
        const fmt = (n) => {
            const num = n || 0;
            return total > 0 ? `${num.toLocaleString()} (${(num/total*100).toFixed(1)}%)` : num.toLocaleString();
        };
        const setTxt = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = fmt(val); };
        setTxt('recYesCount', rec['Yes'] || 0);
        setTxt('recNoCount', rec['No'] || 0);
        setTxt('recMaybeCount', rec['Maybe'] || 0);
    } catch (e) { }
    const reasons = data.recommendation_reasons || {};
    const fillList = (id, items) => {
        const el = document.getElementById(id);
        if (!el) return;
        const arr = (items && items.top) ? items.top : [];
        el.innerHTML = arr.map(([label, pct]) => `<li><span>${label}</span><span>${pct}%</span></li>`).join('');
    };
    fillList('recYesList', reasons.Yes);
    fillList('recNoList', reasons.No);
    // Removed Factors Considered tiles and Maybe list per request

    const brCtx = document.getElementById('branchOverallChart')?.getContext('2d');
    if (brCtx && data.rankings?.branches) {
        const arr = (data.rankings.branches || []).slice();
        const parent = document.getElementById('branchOverallChart')?.parentElement;
        if (parent) parent.style.height = Math.max(400, arr.length * 24) + 'px';
        const colorFor = (v) => {
            const x = Math.max(0, Math.min(1, (v || 0) / 5));
            if (x < 0.5) {
                const t = x / 0.5;
                const r = Math.round(229 + (251-229)*t);
                const g = Math.round(57 + (140-57)*t);
                const b = Math.round(53 + (0-53)*t);
                return `rgb(${r},${g},${b})`;
            } else {
                const t = (x-0.5)/0.5;
                const r = Math.round(251 + (67-251)*t);
                const g = Math.round(140 + (160-140)*t);
                const b = Math.round(0 + (71-0)*t);
                return `rgb(${r},${g},${b})`;
            }
        };
        new Chart(brCtx, {
            type: 'bar',
            data: { labels: arr.map(x => x[0]), datasets: [{ label: 'Overall', data: arr.map(x => x[1]), backgroundColor: arr.map(x => colorFor(x[1])) }] },
            options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', scales: { x: { beginAtZero: true, max: 5 } } }
        });
    }
}

function renderAcademicSection(data) {
    // Subject-wise stacked distribution
    const subj = data.subject_performance || {};
    const subjects = Object.keys(subj);
    const groups = ['Excellent','Good','Average','Poor'];
    const colorMap = { Excellent: '#4caf50', Good: '#2196f3', Average: '#ff9800', Poor: '#e53935' };
    const stackData = groups.map(g => ({ label: g, backgroundColor: colorMap[g], data: subjects.map(s => {
        const dist = subj[s]?.rating_distribution || {};
        const total = Object.values(dist).reduce((a,b)=>a+(b||0),0) || 1;
        const sum = Object.entries(dist).reduce((acc, [k,v]) => {
            const low = String(k).toLowerCase();
            const isAvg = low.includes('average') || low.includes('satisfactory');
            const isNeed = low.includes('need') || low.includes('improve');
            if (g==='Excellent' && low.includes('excellent')) return acc + (v||0);
            if (g==='Good' && low.includes('good')) return acc + (v||0);
            if (g==='Average' && (isAvg)) return acc + (v||0);
            if (g==='Poor' && (low.includes('poor') || isNeed)) return acc + (v||0);
            return acc;
        }, 0);
        return (sum/total*100);
    }) }));
    const sc = document.getElementById('subjectStackedChart')?.getContext('2d');
    if (sc) {
        new Chart(sc, { type: 'bar', data: { labels: subjects, datasets: stackData }, options: { responsive: true, maintainAspectRatio: false, scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true, max: 100 } } } });
    }
    // Subject distribution counts table
    const table = document.getElementById('subjectCountsTable');
    if (table) {
        const header = '<thead><tr><th>Subject</th><th>Excellent</th><th>Good</th><th>Average</th><th>Poor</th><th>Total</th></tr></thead>';
        const rows = subjects.map(name => {
            const dist = subj[name]?.rating_distribution || {};
            let exc = 0, good = 0, avg = 0, poor = 0;
            for (const [k, v] of Object.entries(dist)) {
                const low = String(k).toLowerCase();
                const val = v || 0;
                const isAvg = low.includes('average') || low.includes('satisfactory');
                const isNeed = low.includes('need') || low.includes('improve');
                if (low.includes('excellent')) exc += val;
                else if (low.includes('good')) good += val;
                else if (isAvg) avg += val;
                else if (low.includes('poor') || isNeed) poor += val;
            }
            const total = exc + good + avg + poor;
            return `<tr><td>${name}</td><td>${exc.toLocaleString()}</td><td>${good.toLocaleString()}</td><td>${avg.toLocaleString()}</td><td>${poor.toLocaleString()}</td><td>${total.toLocaleString()}</td></tr>`;
        }).join('');
        table.innerHTML = header + `<tbody>${rows}</tbody>`;
    }

    // Teaching indicators
    const ti = data.teaching_indicators || {};
    const tic = document.getElementById('teachingIndicatorsChart')?.getContext('2d');
    if (tic) {
        const labels = Object.keys(ti);
        const values = labels.map(l => ti[l] || 0);
        new Chart(tic, { type: 'radar', data: { labels, datasets: [{ label: 'Avg', data: values, backgroundColor: 'rgba(118, 75, 162, 0.2)', borderColor: 'rgba(118, 75, 162, 1)', borderWidth: 2 }] }, options: { responsive: true, maintainAspectRatio: false, scales: { r: { beginAtZero: true, max: 5 } } } });
    }

    // PTM effectiveness
    const ptm = data.ptm_effectiveness ?? null;
    const ptmc = document.getElementById('ptmChart')?.getContext('2d');
    if (ptmc) {
        new Chart(ptmc, { type: 'bar', data: { labels: ['PTM'], datasets: [{ label: 'Avg', data: [ptm || 0], backgroundColor: '#ff7043' }] }, options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 5 } } } });
    }
}

function renderEnvironmentSection(data) {
    const env = data.environment_focus || {};
    const labels = Object.keys(env);
    const values = labels.map(l => env[l] || 0);
    const ec = document.getElementById('envRatingsChart')?.getContext('2d');
    if (ec) {
        new Chart(ec, { type: 'bar', data: { labels, datasets: [{ label: 'Avg', data: values, backgroundColor: '#66bb6a' }] }, options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 5 } } } });
    }

    // KPIs: Safety and Hygiene
    const safety = env['Campus safety'] ?? null;
    let hygiene = null;
    const infra = (data.category_performance && data.category_performance['Infrastructure']) || {};
    for (const [k,v] of Object.entries(infra)) {
        const low = String(k).toLowerCase();
        if (low.includes('hygiene') || low.includes('clean')) { hygiene = v?.average ?? hygiene; }
    }
    const setKpi = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = (val==null || isNaN(val)) ? '-' : val.toFixed(2); };
    setKpi('safetyKpi', safety);
    setKpi('hygieneKpi', hygiene);
}

function renderCommunicationSection(data) {
    const cm = data.communication_metrics || {};
    const cc = document.getElementById('communicationChart')?.getContext('2d');
    if (cc) {
        const labels = Object.keys(cm);
        const values = labels.map(l => cm[l] || 0);
        new Chart(cc, { type: 'bar', data: { labels, datasets: [{ label: 'Avg', data: values, backgroundColor: '#29b6f6' }] }, options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 5 } } } });
    }
    try { renderAdminChart(data); } catch(e) { }

    const roles = data.concern_roles || {};
    const rc = document.getElementById('concernRoleChart')?.getContext('2d');
    if (rc) {
        const labels = Object.keys(roles);
        const values = labels.map(l => roles[l] || 0);
        new Chart(rc, { type: 'bar', data: { labels, datasets: [{ label: 'Avg', data: values, backgroundColor: '#ab47bc' }] }, options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 5 } } } });
    }

    const cr = data.concern_resolution || {};
    const crc = document.getElementById('concernResolutionChart')?.getContext('2d');
    if (crc) {
        const labels = ['Yes','No','Not Applicable'];
        const values = labels.map(l => cr[l] || 0);
        new Chart(crc, { type: 'pie', data: { labels, datasets: [{ data: values, backgroundColor: ['#43a047','#e53935','#90a4ae'] }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } } });
    }
    // Populate concern resolution counts KPIs
    try {
        const setTxt = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = (val ?? 0).toLocaleString(); };
        const cr = data.concern_resolution || {};
        setTxt('concernYesCount', cr['Yes'] || 0);
        setTxt('concernNoCount', cr['No'] || 0);
        setTxt('concernNaCount', cr['Not Applicable'] || 0);
    } catch (e) { }
}

function renderInfrastructureSection(data) {
    const infra = (data.category_performance && data.category_performance['Infrastructure']) || {};
    const labels = Object.keys(infra);
    const values = labels.map(l => infra[l]?.average || 0);
    const ic = document.getElementById('infraCategoryChart')?.getContext('2d');
    if (ic) {
        new Chart(ic, { type: 'bar', data: { labels, datasets: [{ label: 'Avg', data: values, backgroundColor: '#ffa726' }] }, options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 5 } } } });
    }

    // Simple heatmap: Branch x {Academics, Infrastructure, Environment, Administration}
    const container = document.getElementById('infraHeatmap');
    if (container) {
        const branches = data.branch_performance || {};
        const rows = Object.entries(branches).map(([b, v]) => ({
            Branch: b,
            Academics: v.subject_avg,
            Infrastructure: v.infrastructure_avg,
            Environment: v.environment_avg,
            Administration: v.admin_avg
        }));
        rows.sort((a,b)=> (a.Infrastructure??0) - (b.Infrastructure??0));
        const makeCell = (val) => {
            const v = val==null || isNaN(val) ? 0 : val;
            const pct = v/5; const red = Math.round((1-pct)*255); const green = Math.round(pct*180);
            return `<td style="background: rgb(${red},${green},80); color:#fff; text-align:center; padding:8px;">${v? v.toFixed(2):'-'}</td>`;
        };
        container.innerHTML = `<div style="overflow:auto"><table class="ranking-table"><thead><tr><th>Branch</th><th>Academics</th><th>Infrastructure</th><th>Environment</th><th>Administration</th></tr></thead><tbody>`+
            rows.map(r => `<tr><td style="background:#001f3f;color:#fff;padding:8px;">${r.Branch}</td>${makeCell(r.Academics)}${makeCell(r.Infrastructure)}${makeCell(r.Environment)}${makeCell(r.Administration)}</tr>`).join('')+
            `</tbody></table></div>`;
    }
}

function renderStrengthsSection(data) {
    const cat = Object.assign({}, data.summary.category_scores || {});
    // Derive Communication aggregate
    const cm = data.communication_metrics || {};
    const commVals = Object.values(cm);
    if (commVals.length) cat['Communication'] = commVals.reduce((a,b)=>a+(b||0),0)/commVals.length;
    // Add Safety/Hygiene/PTM if present
    const safety = data.environment_focus?.['Campus safety'];
    if (safety!=null) cat['Safety'] = safety;
    let hygiene = null; const infra = data.category_performance?.['Infrastructure'] || {};
    for (const [k,v] of Object.entries(infra)) { const low = k.toLowerCase(); if (low.includes('hygiene')||low.includes('clean')) { hygiene = v?.average ?? hygiene; } }
    if (hygiene!=null) cat['Hygiene'] = hygiene;
    if (data.ptm_effectiveness!=null) cat['PTM'] = data.ptm_effectiveness;

    const pairs = Object.entries(cat).filter(([,v])=> v!=null && !isNaN(v));
    const top = pairs.slice().sort((a,b)=>b[1]-a[1]).slice(0,5);
    const low = pairs.slice().sort((a,b)=>a[1]-b[1]).slice(0,5);

    const strengthCtx = document.getElementById('topStrengthsChart')?.getContext('2d');
    if (strengthCtx) {
        new Chart(strengthCtx, { type: 'bar', data: { labels: top.map(x=>x[0]), datasets: [{ label: 'Avg', data: top.map(x=>x[1]), backgroundColor: '#26a69a' }] }, options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', scales: { x: { beginAtZero: true, max: 5 } } } });
    }

    const impCtx = document.getElementById('topImprovementsChart')?.getContext('2d');
    if (impCtx) {
        new Chart(impCtx, { type: 'bar', data: { labels: low.map(x=>x[0]), datasets: [{ label: 'Avg', data: low.map(x=>x[1]), backgroundColor: '#ef5350' }] }, options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', scales: { x: { beginAtZero: true, max: 5 } } } });
    }
}

function renderBranchComparisonSection(data) {
    const ranked = data.rankings?.branches || [];
    const rankCtx = document.getElementById('branchRankedChart')?.getContext('2d');
    if (rankCtx) {
        const arr = ranked.slice();
        const parent = document.getElementById('branchRankedChart')?.parentElement;
        if (parent) parent.style.height = Math.max(400, arr.length * 24) + 'px';
        const colorFor = (v) => {
            const x = Math.max(0, Math.min(1, (v || 0) / 5));
            if (x < 0.5) {
                const t = x / 0.5;
                const r = Math.round(229 + (251-229)*t);
                const g = Math.round(57 + (140-57)*t);
                const b = Math.round(53 + (0-53)*t);
                return `rgb(${r},${g},${b})`;
            } else {
                const t = (x-0.5)/0.5;
                const r = Math.round(251 + (67-251)*t);
                const g = Math.round(140 + (160-140)*t);
                const b = Math.round(0 + (71-0)*t);
                return `rgb(${r},${g},${b})`;
            }
        };
        new Chart(rankCtx, { type: 'bar', data: { labels: arr.map(x=>x[0]), datasets: [{ label: 'Overall', data: arr.map(x=>x[1]), backgroundColor: arr.map(x=>colorFor(x[1])) }] }, options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', scales: { x: { beginAtZero: true, max: 5 } } } });
    }

    let brPct = data.branch_recommendation_pct || {};
    // Fallback: if no pct available, derive from counts when possible
    try {
        const hasAny = Object.values(brPct || {}).some(v => v != null && !isNaN(v));
        if (!hasAny && data.branch_recommendation_counts) {
            const tmp = {};
            for (const [b, c] of Object.entries(data.branch_recommendation_counts)) {
                const yes = c?.Yes || 0, no = c?.No || 0, maybe = c?.Maybe || 0;
                const tot = yes + no + maybe;
                tmp[b] = tot > 0 ? (yes / tot * 100.0) : null;
            }
            brPct = tmp;
        }
    } catch (_e) {}
    const recCtx = document.getElementById('branchRecommendChart')?.getContext('2d');
    if (recCtx) {
        const entries = Object.entries(brPct).filter(([,v])=> v!=null);
        entries.sort((a,b)=>b[1]-a[1]);
        const labels = entries.slice(0,20).map(x=>x[0].slice(0,18));
        const values = entries.slice(0,20).map(x=>x[1]);
        new Chart(recCtx, { type: 'bar', data: { labels, datasets: [{ label: '% Recommend', data: values, backgroundColor: '#8d6e63' }] }, options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 100 } } } });
    }

    const scatterCtx = document.getElementById('branchScatterChart')?.getContext('2d');
    if (scatterCtx) {
        const branches = data.branch_performance || {};
        const points = Object.entries(branches).map(([name,val])=> ({ x: val.subject_avg || 0, y: val.infrastructure_avg || 0, r: Math.max(4, Math.min(10, (val.count||10)/50)), label: name }));
        new Chart(scatterCtx, { type: 'scatter', data: { datasets: [{ label: 'Branches', data: points, parsing: false, pointBackgroundColor: '#42a5f5' }] }, options: { responsive: true, maintainAspectRatio: false, scales: { x: { title: { display: true, text: 'Academics' }, min: 0, max: 5 }, y: { title: { display: true, text: 'Infrastructure' }, min: 0, max: 5 } }, plugins: { tooltip: { callbacks: { label: (ctx)=> `${ctx.raw.label}: (${ctx.raw.x.toFixed(2)}, ${ctx.raw.y.toFixed(2)})` } } } } });
    }

    // Helpers for filters and percentages
    const classSel = document.getElementById('branchClassFilter');
    const orientSel = document.getElementById('branchOrientationFilter');
    const pctStr = (num, den) => {
        if (!den) return '-';
        return `${((num/den)*100).toFixed(1)}%`;
    };
    const fillSelect = (sel, keys) => {
        if (!sel) return;
        if (sel.options.length <= 1) {
            keys.forEach(k => {
                const opt = document.createElement('option');
                opt.value = k; opt.textContent = k; sel.appendChild(opt);
            });
        }
    };
    // Populate filter options from summary
    fillSelect(classSel, Object.keys(data.summary?.classes || {}));
    fillSelect(orientSel, Object.keys(data.summary?.orientations || {}));

    // Resolve current recommendation counts source based on filters
    const currentRecCounts = () => {
        const cls = classSel?.value || '';
        const ori = orientSel?.value || '';
        const by = data.branch_recommendation_counts_by || {};
        if (cls && ori) return (by.pair?.[cls]?.[ori]) || {};
        if (cls) return (by.class?.[cls]) || {};
        if (ori) return (by.orientation?.[ori]) || {};
        return data.branch_recommendation_counts || {};
    };
    // Resolve current rating counts source based on filters
    const currentRatingCounts = () => {
        const cls = classSel?.value || '';
        const ori = orientSel?.value || '';
        const by = data.branch_rating_counts_by || {};
        if (cls && ori) return (by.pair?.[cls]?.[ori]) || {};
        if (cls) return (by.class?.[cls]) || {};
        if (ori) return (by.orientation?.[ori]) || {};
        return data.branch_rating_counts || {};
    };

    // Update recommendation tables (counts + %)
    const updateRecTables = () => {
        const counts = currentRecCounts();
        const rows = Object.entries(counts).map(([b, c])=> {
            const yes = c?.Yes || 0, no = c?.No || 0, maybe = c?.Maybe || 0, na = c?.['Not Applicable'] || 0;
            const totalRec = yes + no + maybe; // denominator for %
            const totalAll = totalRec + na;
            return { branch: b, yes, no, maybe, na, totalRec, totalAll };
        });
        const topYes = rows.slice().sort((a,b)=> b.yes - a.yes).slice(0, 15);
        const topNo = rows.slice().sort((a,b)=> b.no - a.no).slice(0, 15);
        const render = (id, rws) => {
            const el = document.getElementById(id); if (!el) return;
            el.innerHTML = '<thead><tr><th>Branch</th><th>Yes</th><th>No</th><th>Maybe</th><th>NA</th><th>Total</th></tr></thead>' +
                '<tbody>' + rws.map(r=> `<tr><td>${r.branch}</td>`+
                `<td>${r.yes.toLocaleString()} (${pctStr(r.yes, r.totalRec)})</td>`+
                `<td>${r.no.toLocaleString()} (${pctStr(r.no, r.totalRec)})</td>`+
                `<td>${r.maybe.toLocaleString()} (${pctStr(r.maybe, r.totalRec)})</td>`+
                `<td>${r.na.toLocaleString()}</td>`+
                `<td>${r.totalAll.toLocaleString()}</td>`+
                `</tr>`).join('') + '</tbody>';
        };
        render('branchYesRecsTable', topYes);
        render('branchNoRecsTable', topNo);
    };

    // Update rating tables (Poor/Excellent: counts + %)
    const updateRatingTables = () => {
        const ratingCounts = currentRatingCounts();
        const select = document.getElementById('ratingGroupSelect');
        const group = select?.value || 'Subjects';
        const rows = Object.entries(ratingCounts).map(([b, groups]) => {
            const g = groups?.[group] || {};
            const ex = g.Excellent || 0, gd = g.Good || 0, av = g.Average || 0, pr = g.Poor || 0;
            const tot = ex + gd + av + pr;
            return { branch: b, Excellent: ex, Poor: pr, Total: tot };
        });
        const topPoor = rows.slice().sort((a,b)=> b.Poor - a.Poor).slice(0, 15);
        const topExcellent = rows.slice().sort((a,b)=> b.Excellent - a.Excellent).slice(0, 15);
        const render = (id, rws, key, label) => {
            const el = document.getElementById(id); if (!el) return;
            el.innerHTML = `<thead><tr><th>Branch</th><th>${label} Count</th><th>${label} %</th></tr></thead>` +
                '<tbody>' + rws.map(r=> `<tr><td>${r.branch}</td><td>${r[key].toLocaleString()}</td><td>${pctStr(r[key], r.Total)}</td></tr>`).join('') + '</tbody>';
        };
        render('branchPoorTable', topPoor, 'Poor', 'Poor');
        render('branchExcellentTable', topExcellent, 'Excellent', 'Excellent');
    };

    // Wire up listeners
    const ratingGroupSel = document.getElementById('ratingGroupSelect');
    if (ratingGroupSel) ratingGroupSel.addEventListener('change', updateRatingTables);
    if (classSel) classSel.addEventListener('change', () => { updateRecTables(); updateRatingTables(); });
    if (orientSel) orientSel.addEventListener('change', () => { updateRecTables(); updateRatingTables(); });

    // Initial render
    updateRecTables();
    updateRatingTables();
}
