document.addEventListener('DOMContentLoaded', () => {
    if (!window.Auth.check()) return;

    fetch('data.json')
        .then(response => response.json())
        .then(data => initDashboard(data))
        .catch(error => console.error('Error loading data:', error));
});

const chartColors = ['#3b82f6', '#8b5cf6', '#10b981', '#f59e0b', '#ef4444'];
let appData = null;

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
    const grouped = {};
    data.subject_stats.forEach(item => {
        if (!grouped[item.subject]) grouped[item.subject] = {};
        grouped[item.subject][item.exam_key] = item.avg_score;
    });

    const subjects = Object.keys(grouped);
    const maxScore = Math.max(
        ...data.subject_stats.map(item => Number(item.avg_score) || 0),
        100
    );

    chart.setOption({
        tooltip: { backgroundColor: 'rgba(255,255,255,0.95)', textStyle: { color: '#333' } },
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
    const classes = [...new Set(data.class_stats.map(item => item.class_name))].sort((a, b) => String(a).localeCompare(String(b), 'zh-CN'));
    const series = data.exams.map((exam, index) => ({
        name: exam.exam_name,
        type: 'bar',
        itemStyle: { color: chartColors[index % chartColors.length] },
        label: {
            show: classes.length <= 12,
            position: 'top',
            color: '#333',
            formatter: params => params.value == null ? '' : Number(params.value).toFixed(1)
        },
        data: classes.map(className => {
            const found = data.class_stats.find(item => item.exam_key === exam.exam_key && item.class_name === className);
            return found ? found.avg_total_score : null;
        })
    }));

    chart.setOption({
        tooltip: { trigger: 'axis', backgroundColor: 'rgba(255,255,255,0.95)', textStyle: { color: '#333' } },
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

    const categories = ['总排名', ...student.subjects.map(item => item.name)];
    const series = appData.exams.map((exam, examIndex) => {
        const data = [];
        const totalStat = student.exam_stats.find(item => item.exam_key === exam.exam_key);
        data.push(totalStat?.total_joint_rank ?? null);

        student.subjects.forEach(subject => {
            const stat = subject.exam_stats.find(item => item.exam_key === exam.exam_key);
            data.push(stat?.joint_rank ?? null);
        });

        const isLatest = exam.exam_key === windowInfo.to_exam_key;
        const prevKey = windowInfo.from_exam_key;

        if (isLatest) {
            const coloredData = data.map((value, idx) => {
                const prevValue = idx === 0
                    ? (student.exam_stats.find(item => item.exam_key === prevKey)?.total_joint_rank ?? null)
                    : (student.subjects[idx - 1].exam_stats.find(item => item.exam_key === prevKey)?.joint_rank ?? null);
                const change = prevValue != null && value != null ? prevValue - value : 0;
                return {
                    value,
                    itemStyle: {
                        color: change > 0 ? '#ef4444' : (change < 0 ? '#10b981' : chartColors[examIndex % chartColors.length])
                    }
                };
            });
            return {
                name: `${exam.exam_name} (红升绿降)`,
                type: 'bar',
                data: coloredData,
                label: { show: true, position: 'top', color: '#333' }
            };
        }

        return {
            name: exam.exam_name,
            type: 'bar',
            itemStyle: { color: chartColors[examIndex % chartColors.length] },
            data,
            label: { show: false }
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
                    const value = typeof param.value === 'object' ? param.value.value : param.value;
                    html += `${param.marker}${param.seriesName}: ${formatValue(value)}名<br/>`;
                });

                const latestValue = idx === 0
                    ? (student.exam_stats.find(item => item.exam_key === windowInfo.to_exam_key)?.total_joint_rank ?? null)
                    : (student.subjects[idx - 1].exam_stats.find(item => item.exam_key === windowInfo.to_exam_key)?.joint_rank ?? null);
                const prevValue = idx === 0
                    ? (student.exam_stats.find(item => item.exam_key === windowInfo.from_exam_key)?.total_joint_rank ?? null)
                    : (student.subjects[idx - 1].exam_stats.find(item => item.exam_key === windowInfo.from_exam_key)?.joint_rank ?? null);
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
