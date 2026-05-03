document.addEventListener('DOMContentLoaded', () => {
    if (!window.Auth.check()) return;
    setupModalActions();

    fetch('data.json')
        .then(response => response.json())
        .then(data => initDashboard(data))
        .catch(error => console.error('Error loading data:', error));
});

const chartColors = ['#3b82f6', '#8b5cf6', '#10b981', '#f59e0b', '#ef4444'];
const SUBJECT_ORDER = ['语文', '数学', '英语', '道德与法治', '历史', '地理', '生物'];
let appData = null;
const chartInstances = {};
let modalChart = null;

function initDashboard(data) {
    appData = data;
    renderCurrentDate();
    renderOverviewCharts(data);
    setupSearch(data.students);
    updateRankListTitles(data.comparison_window);
}

function renderCurrentDate() {
    const dateEl = document.getElementById('currentDate');
    if (!dateEl) return;
    const now = new Date();
    dateEl.textContent = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
}

function renderOverviewCharts(data) {
    renderScoreDistributionChart(data);
    renderSubjectRadarChart(data);
    renderClassChart(data);
    renderRankList('topImproversList', data.top_improvers);
    renderRankList('bottomImproversList', data.bottom_improvers);
}

function getBins(scoreMap) {
    const allScores = Object.values(scoreMap).flatMap(item => item.scores || []);
    const maxScore = Math.max(...allScores, 0);
    const upper = Math.max(600, Math.ceil(maxScore / 50) * 50 + 50);
    const bins = [];
    for (let i = 0; i <= upper; i += 50) bins.push(i);
    return bins;
}

function renderScoreDistributionChart(data) {
    const chart = echarts.init(document.getElementById('scoreDistChart'));
    chartInstances.scoreDistChart = chart;
    const distributions = data.global_stats.score_distribution || {};
    const bins = getBins(distributions);
    const labels = bins.slice(0, -1).map((b, i) => `${b}-${bins[i + 1]}`);
    const series = data.exams.map((exam, index) => ({
        name: exam.exam_name,
        type: 'line',
        smooth: true,
        areaStyle: { opacity: 0.1 },
        lineStyle: { width: 2 },
        itemStyle: { color: chartColors[index % chartColors.length] },
        data: histogram((distributions[exam.exam_key] || {}).scores || [], bins)
    }));

    chart.setOption({
        tooltip: { trigger: 'axis', backgroundColor: 'rgba(255,255,255,0.95)', textStyle: { color: '#333' } },
        legend: { data: data.exams.map(item => item.exam_name), textStyle: { color: '#333' } },
        xAxis: { type: 'category', data: labels, axisLabel: { color: '#666' } },
        yAxis: { type: 'value', axisLabel: { color: '#666' }, splitLine: { lineStyle: { color: '#eee' } } },
        series
    });

    window.addEventListener('resize', () => chart.resize());
}

function renderSubjectRadarChart(data) {
    const chart = echarts.init(document.getElementById('subjectAvgChart'));
    chartInstances.subjectAvgChart = chart;
    const grouped = {};
    data.subject_stats.forEach(item => {
        if (!grouped[item.subject]) grouped[item.subject] = {};
        grouped[item.subject][item.exam_key] = item.avg_score;
    });

    const subjects = getOrderedSubjectNames(Object.keys(grouped));
    const maxScore = Math.max(
        ...data.subject_stats.map(item => Number(item.avg_score) || 0),
        100
    );

    chart.setOption({
        tooltip: {
            backgroundColor: 'rgba(255,255,255,0.95)',
            textStyle: { color: '#333' },
            formatter: params => {
                const values = Array.isArray(params.value) ? params.value : [];
                const details = subjects
                    .map((subject, index) => `${subject}: ${formatFixed(values[index], 2)}`)
                    .join('<br/>');
                return `${params.marker}${params.seriesName}<br/>${details}`;
            }
        },
        legend: { data: data.exams.map(item => item.exam_name), textStyle: { color: '#333' } },
        radar: {
            indicator: subjects.map(subject => ({ name: subject, max: Math.ceil(maxScore / 10) * 10 })),
            axisName: { color: '#333' },
            splitArea: { areaStyle: { color: ['#fff', '#f8fafc'] } },
            splitLine: { lineStyle: { color: '#cbd5e1' } }
        },
        series: [{
            type: 'radar',
            data: data.exams.map((exam, index) => ({
                name: exam.exam_name,
                value: subjects.map(subject => grouped[subject][exam.exam_key] || 0),
                lineStyle: { color: chartColors[index % chartColors.length] },
                itemStyle: { color: chartColors[index % chartColors.length] },
                areaStyle: { opacity: 0.05 }
            }))
        }]
    });

    window.addEventListener('resize', () => chart.resize());
}

function renderClassChart(data) {
    const chart = echarts.init(document.getElementById('classAvgChart'));
    chartInstances.classAvgChart = chart;
    const classes = [...new Set(data.class_stats.map(item => item.class_name))].sort(compareClassNames);
    const series = data.exams.map((exam, index) => ({
        name: exam.exam_name,
        type: 'bar',
        itemStyle: { color: chartColors[index % chartColors.length] },
        label: {
            show: true,
            position: 'top',
            color: '#333',
            formatter: params => params.value == null ? '' : formatFixed(params.value, 2)
        },
        data: classes.map(className => {
            const found = data.class_stats.find(item => item.exam_key === exam.exam_key && item.class_name === className);
            return found ? found.avg_total_score : null;
        })
    }));

    chart.setOption({
        tooltip: {
            trigger: 'axis',
            backgroundColor: 'rgba(255,255,255,0.95)',
            textStyle: { color: '#333' },
            formatter: params => {
                let html = `${params[0].axisValue}<br/>`;
                params.forEach(param => {
                    html += `${param.marker}${param.seriesName}: ${formatFixed(param.value, 2)}<br/>`;
                });
                return html;
            }
        },
        legend: { data: data.exams.map(item => item.exam_name), textStyle: { color: '#333' } },
        grid: { left: '3%', right: '4%', bottom: '15%', containLabel: true },
        dataZoom: [
            { type: 'slider', xAxisIndex: 0, start: 0, end: classes.length > 12 ? 50 : 100, bottom: '2%' },
            { type: 'inside', xAxisIndex: 0, start: 0, end: classes.length > 12 ? 50 : 100 }
        ],
        xAxis: { type: 'category', data: classes, axisLabel: { color: '#666', rotate: 45, interval: 0 } },
        yAxis: { type: 'value', axisLabel: { color: '#666' }, splitLine: { lineStyle: { color: '#eee' } } },
        series
    });

    window.addEventListener('resize', () => chart.resize());
}

function histogram(data, bins) {
    const hist = new Array(bins.length - 1).fill(0);
    (data || []).forEach(val => {
        for (let i = 0; i < bins.length - 1; i++) {
            if (val >= bins[i] && val < bins[i + 1]) {
                hist[i] += 1;
                break;
            }
        }
    });
    return hist;
}

function updateRankListTitles(windowInfo) {
    const cards = document.querySelectorAll('.right-panel .card h2');
    if (cards.length < 3 || !windowInfo) return;
    cards[1].textContent = `进步榜 (${windowInfo.from_exam_name} -> ${windowInfo.to_exam_name})`;
    cards[2].textContent = `退步榜 (${windowInfo.from_exam_name} -> ${windowInfo.to_exam_name})`;
}

function renderRankList(elementId, list) {
    const ul = document.getElementById(elementId);
    ul.dataset.rankMode = elementId === 'topImproversList' ? 'positive' : 'negative';
    ul.onclick = () => openRankModal(ul.dataset.rankMode);
    ul.innerHTML = '';
    if (!list || list.length === 0) {
        ul.innerHTML = '<li class="rank-item">暂无数据</li>';
        return;
    }

    list.forEach(item => {
        const li = document.createElement('li');
        const change = Number(item.rank_change || 0);
        const sign = change > 0 ? '+' : '';
        const colorClass = change > 0 ? 'positive' : (change < 0 ? 'negative' : '');
        li.className = 'rank-item';
        li.innerHTML = `
            <span>${item.name} <small>(${formatClass(item.class)})</small></span>
            <span class="${colorClass}">${sign}${change}</span>
        `;
        ul.appendChild(li);
    });
}

function setupSearch(students) {
    const searchInput = document.getElementById('studentSearch');
    const resultsDiv = document.getElementById('searchResults');

    searchInput.addEventListener('input', e => {
        const query = e.target.value.trim();
        resultsDiv.innerHTML = '';
        if (!query) return;

        const matches = students
            .filter(student => (student.name || '').includes(query))
            .slice(0, 10);

        matches.forEach(student => {
            const div = document.createElement('div');
            div.className = 'search-item';
            div.textContent = `${student.name} (${formatClass(student.class)}) - 学号:${student.student_id}`;
            div.onclick = () => {
                showStudentDetail(student);
                resultsDiv.innerHTML = '';
                searchInput.value = student.name;
            };
            resultsDiv.appendChild(div);
        });
    });
}

function showStudentDetail(student) {
    document.getElementById('welcomeMsg').style.display = 'none';
    document.getElementById('studentDetail').style.display = 'flex';

    const windowInfo = appData.comparison_window;
    const prevExam = student.exam_stats.find(item => item.exam_key === windowInfo.from_exam_key);
    const latestExam = student.exam_stats.find(item => item.exam_key === windowInfo.to_exam_key);

    document.getElementById('studentName').textContent = student.name;
    document.getElementById('studentClass').textContent = formatClass(student.class);
    document.getElementById('scoreLabel1').textContent = `${windowInfo.from_exam_name}总分`;
    document.getElementById('scoreLabel2').textContent = `${windowInfo.to_exam_name}总分`;
    document.getElementById('rankLabel1').textContent = `${windowInfo.from_exam_name}联考排名`;
    document.getElementById('rankLabel2').textContent = `${windowInfo.to_exam_name}联考排名`;
    document.getElementById('changeLabel').textContent = `排名变化`;

    document.getElementById('monthlyScore').textContent = formatValue(prevExam?.total_score);
    document.getElementById('midtermScore').textContent = formatValue(latestExam?.total_score);
    document.getElementById('monthlyRank').textContent = formatValue(prevExam?.total_joint_rank);
    document.getElementById('midtermRank').textContent = formatValue(latestExam?.total_joint_rank);

    const rankChange = Number(student.latest_rank_change || 0);
    const rankEl = document.getElementById('rankChange');
    rankEl.textContent = `${rankChange > 0 ? '+' : ''}${rankChange}`;
    rankEl.className = rankChange > 0 ? 'positive' : (rankChange < 0 ? 'negative' : '');

    renderStudentSubjectChart(student, windowInfo);
}

function renderStudentSubjectChart(student, windowInfo) {
    const chartDom = document.getElementById('studentSubjectChart');
    let chart = echarts.getInstanceByDom(chartDom);
    if (chart) chart.dispose();
    chart = echarts.init(chartDom);
    const orderedSubjects = getOrderedStudentSubjects(student);
    const categories = ['总分', ...orderedSubjects.map(item => item.name)];
    const series = appData.exams.map((exam, examIndex) => {
        const data = buildStudentBarData(student, orderedSubjects, exam.exam_key, windowInfo, examIndex);
        const isLatest = exam.exam_key === windowInfo.to_exam_key;
        return {
            name: isLatest ? `${exam.exam_name} (红升绿降)` : exam.exam_name,
            type: 'bar',
            data,
            itemStyle: { color: chartColors[examIndex % chartColors.length] },
            label: {
                show: true,
                position: 'top',
                color: '#333',
                formatter: params => {
                    const rank = typeof params.value === 'object' ? params.value.value : params.value;
                    return rank == null ? '' : `${formatValue(rank)}名`;
                }
            }
        };
    });

    chart.setOption({
        tooltip: {
            trigger: 'axis',
            backgroundColor: 'rgba(255,255,255,0.95)',
            textStyle: { color: '#333' },
            axisPointer: { type: 'shadow' },
            formatter: params => {
                const idx = params[0].dataIndex;
                let html = `${params[0].axisValue}<br/>`;
                params.forEach(param => {
                    const rank = typeof param.value === 'object' ? param.value.value : param.value;
                    const score = param.data && typeof param.data === 'object' ? param.data.score : null;
                    html += `${param.marker}${param.seriesName}: ${formatValue(rank)}名`;
                    if (score != null) {
                        html += `，分数 ${formatValue(score)}`;
                    }
                    html += '<br/>';
                });

                const latestValue = getRankValueByIndex(student, orderedSubjects, windowInfo.to_exam_key, idx);
                const prevValue = getRankValueByIndex(student, orderedSubjects, windowInfo.from_exam_key, idx);
                if (latestValue != null && prevValue != null) {
                    const change = prevValue - latestValue;
                    const label = change > 0 ? '进步' : (change < 0 ? '退步' : '持平');
                    const color = change > 0 ? 'red' : (change < 0 ? 'green' : 'gray');
                    html += `<span style="color:${color};font-weight:bold">对比${windowInfo.from_exam_name}: ${label} ${Math.abs(change)} 名</span>`;
                }
                return html;
            }
        },
        legend: {
            data: series.map(item => item.name),
            textStyle: { color: '#333' }
        },
        grid: {
            left: '3%',
            right: '4%',
            bottom: '15%',
            containLabel: true
        },
        dataZoom: [
            { type: 'slider', show: true, xAxisIndex: [0], start: 0, end: 100, bottom: '2%' },
            { type: 'inside', xAxisIndex: [0], start: 0, end: 100 }
        ],
        xAxis: {
            type: 'category',
            data: categories,
            axisLabel: { color: '#333', interval: 0, fontWeight: 'bold' }
        },
        yAxis: {
            type: 'value',
            inverse: false,
            name: '联考排名 (数值越小越好)',
            nameTextStyle: { color: '#666' },
            axisLabel: { color: '#666' },
            splitLine: { lineStyle: { color: '#eee' } }
        },
        series
    });

    window.addEventListener('resize', () => chart.resize());
}

function formatValue(value) {
    if (value == null || value === '') return '-';
    const num = Number(value);
    return Number.isFinite(num) ? (Number.isInteger(num) ? num : num.toFixed(1)) : value;
}

function formatClass(value) {
    if (value == null || value === '') return '-';
    const text = String(value).trim();
    return text.endsWith('班') ? text : `${text}班`;
}

function formatFixed(value, digits = 2) {
    if (value == null || value === '') return '-';
    const num = Number(value);
    return Number.isFinite(num) ? num.toFixed(digits) : value;
}

function getOrderedSubjectNames(names) {
    const nameSet = new Set((names || []).filter(Boolean));
    const ordered = SUBJECT_ORDER.filter(subject => nameSet.has(subject));
    const remaining = [...nameSet]
        .filter(subject => !SUBJECT_ORDER.includes(subject))
        .sort((a, b) => String(a).localeCompare(String(b), 'zh-CN'));
    return [...ordered, ...remaining];
}

function getOrderedStudentSubjects(student) {
    const subjectMap = new Map((student.subjects || []).map(subject => [subject.name, subject]));
    return getOrderedSubjectNames((student.subjects || []).map(subject => subject.name))
        .map(name => subjectMap.get(name))
        .filter(Boolean);
}

function compareClassNames(a, b) {
    const aText = formatClass(a);
    const bText = formatClass(b);
    const aMatch = aText.match(/\d+/);
    const bMatch = bText.match(/\d+/);
    if (aMatch && bMatch) {
        return Number(aMatch[0]) - Number(bMatch[0]);
    }
    return aText.localeCompare(bText, 'zh-CN');
}

function buildStudentBarData(student, orderedSubjects, examKey, windowInfo, examIndex) {
    const entries = [
        {
            name: '总分',
            rank: student.exam_stats.find(item => item.exam_key === examKey)?.total_joint_rank ?? null,
            score: student.exam_stats.find(item => item.exam_key === examKey)?.total_score ?? null
        },
        ...orderedSubjects.map(subject => {
            const stat = subject.exam_stats.find(item => item.exam_key === examKey);
            return {
                name: subject.name,
                rank: stat?.joint_rank ?? null,
                score: stat?.score ?? null
            };
        })
    ];

    return entries.map((entry, index) => {
        const dataItem = {
            value: entry.rank,
            score: entry.score
        };

        if (examKey === windowInfo.to_exam_key) {
            const previousRank = getRankValueByIndex(student, orderedSubjects, windowInfo.from_exam_key, index);
            const change = previousRank != null && entry.rank != null ? previousRank - entry.rank : 0;
            dataItem.itemStyle = {
                color: change > 0 ? '#ef4444' : (change < 0 ? '#10b981' : chartColors[examIndex % chartColors.length])
            };
        }

        return dataItem;
    });
}

function getRankValueByIndex(student, orderedSubjects, examKey, index) {
    if (index === 0) {
        return student.exam_stats.find(item => item.exam_key === examKey)?.total_joint_rank ?? null;
    }
    const subject = orderedSubjects[index - 1];
    return subject?.exam_stats.find(item => item.exam_key === examKey)?.joint_rank ?? null;
}

function setupModalActions() {
    document.querySelectorAll('[data-modal-chart]').forEach(button => {
        button.addEventListener('click', () => {
            openChartModal(button.dataset.modalChart, button.dataset.modalTitle || '图表详情');
        });
    });

    document.querySelectorAll('[data-rank-mode]').forEach(button => {
        button.addEventListener('click', () => {
            openRankModal(button.dataset.rankMode);
        });
    });

    const overlay = document.getElementById('modalOverlay');
    const closeButton = document.getElementById('modalCloseBtn');
    closeButton.addEventListener('click', closeModal);
    overlay.addEventListener('click', event => {
        if (event.target === overlay) {
            closeModal();
        }
    });
    document.addEventListener('keydown', event => {
        if (event.key === 'Escape') {
            closeModal();
        }
    });
}

function openChartModal(chartId, title) {
    const sourceChart = chartInstances[chartId];
    if (!sourceChart) return;

    const overlay = document.getElementById('modalOverlay');
    const chartContainer = document.getElementById('modalChartContainer');
    const detailContainer = document.getElementById('modalDetailContainer');
    document.getElementById('modalTitle').textContent = title;
    detailContainer.classList.add('hidden');
    chartContainer.classList.remove('hidden');
    detailContainer.innerHTML = '';
    overlay.classList.remove('hidden');

    if (modalChart) {
        modalChart.dispose();
    }
    modalChart = echarts.init(chartContainer);
    modalChart.setOption(sourceChart.getOption(), true);
    requestAnimationFrame(() => modalChart.resize());
}

function openRankModal(mode) {
    if (!appData) return;

    const overlay = document.getElementById('modalOverlay');
    const chartContainer = document.getElementById('modalChartContainer');
    const detailContainer = document.getElementById('modalDetailContainer');
    const modeLabel = mode === 'positive' ? '各班进步前五' : '各班退步前五';
    const windowInfo = appData.comparison_window || {};

    document.getElementById('modalTitle').textContent = `${modeLabel} (${windowInfo.from_exam_name} -> ${windowInfo.to_exam_name})`;
    chartContainer.classList.add('hidden');
    detailContainer.classList.remove('hidden');
    detailContainer.innerHTML = buildRankModalHtml(mode);
    overlay.classList.remove('hidden');

    if (modalChart) {
        modalChart.dispose();
        modalChart = null;
    }
}

function buildRankModalHtml(mode) {
    const grouped = new Map();
    (appData.students || []).forEach(student => {
        const className = formatClass(student.class);
        if (!grouped.has(className)) {
            grouped.set(className, []);
        }
        grouped.get(className).push(student);
    });

    const classCards = [...grouped.entries()]
        .sort((a, b) => compareClassNames(a[0], b[0]))
        .map(([className, students]) => {
            const filtered = students
                .filter(student => {
                    const change = Number(student.latest_rank_change || 0);
                    return mode === 'positive' ? change > 0 : change < 0;
                })
                .sort((a, b) => mode === 'positive'
                    ? Number(b.latest_rank_change || 0) - Number(a.latest_rank_change || 0)
                    : Number(a.latest_rank_change || 0) - Number(b.latest_rank_change || 0))
                .slice(0, 5);

            if (filtered.length === 0) {
                return `
                    <div class="rank-class-card">
                        <h3>${className}</h3>
                        <div class="rank-class-empty">暂无符合条件的数据</div>
                    </div>
                `;
            }

            const listHtml = filtered.map(student => {
                const change = Number(student.latest_rank_change || 0);
                const sign = change > 0 ? '+' : '';
                const classNameText = mode === 'positive' ? 'positive' : 'negative';
                return `
                    <li>
                        <span>${student.name}</span>
                        <span class="${classNameText}">${sign}${change}</span>
                    </li>
                `;
            }).join('');

            return `
                <div class="rank-class-card">
                    <h3>${className}</h3>
                    <ul>${listHtml}</ul>
                </div>
            `;
        });

    return `<div class="rank-class-grid">${classCards.join('')}</div>`;
}

function closeModal() {
    const overlay = document.getElementById('modalOverlay');
    overlay.classList.add('hidden');
    document.getElementById('modalDetailContainer').innerHTML = '';
    if (modalChart) {
        modalChart.dispose();
        modalChart = null;
    }
}
