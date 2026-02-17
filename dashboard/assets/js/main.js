document.addEventListener('DOMContentLoaded', () => {
    // 1. Load Data
    fetch('data.json')
        .then(response => response.json())
        .then(data => {
            initDashboard(data);
        })
        .catch(error => console.error('Error loading data:', error));
});

let charts = {};

function initDashboard(data) {
    // Render Overview Charts
    renderOverviewCharts(data);
    
    // Setup Search
    setupSearch(data.students);
}

function renderOverviewCharts(data) {
    // 1. Score Distribution (Left Panel)
    const scoreChart = echarts.init(document.getElementById('scoreDistChart'));
    // Need to process data for histogram - simplified for demo using line/bar
    // Assuming data.global_stats.score_distribution contains lists of scores
    
    // Create bins for histogram
    const bins = [0, 300, 350, 400, 450, 500, 550, 600];
    const monthlyHist = histogram(data.global_stats.score_distribution.monthly, bins);
    const midtermHist = histogram(data.global_stats.score_distribution.midterm, bins);
    
    scoreChart.setOption({
        tooltip: { trigger: 'axis', backgroundColor: 'rgba(255,255,255,0.9)', textStyle: {color: '#333'} },
        legend: { data: ['第一次月考', '期中考试'], textStyle: { color: '#333' } },
        xAxis: { 
            type: 'category', 
            data: bins.slice(0, -1).map((b, i) => `${b}-${bins[i+1]}`),
            axisLabel: { color: '#666' }
        },
        yAxis: { type: 'value', axisLabel: { color: '#666' }, splitLine: { lineStyle: { color: '#eee' } } },
        series: [
            { name: '第一次月考', type: 'line', data: monthlyHist, smooth: true, areaStyle: { opacity: 0.3 } },
            { name: '期中考试', type: 'line', data: midtermHist, smooth: true, areaStyle: { opacity: 0.3 } }
        ]
    });

    // 2. Subject Averages (Left Panel)
    const subjectChart = echarts.init(document.getElementById('subjectAvgChart'));
    const subjects = data.subject_stats.map(s => s.Subject);
    const monthlyAvgs = data.subject_stats.map(s => s.Avg_Score_Monthly);
    const midtermAvgs = data.subject_stats.map(s => s.Avg_Score_Midterm);

    subjectChart.setOption({
        tooltip: { trigger: 'axis', backgroundColor: 'rgba(255,255,255,0.9)', textStyle: {color: '#333'} },
        radar: {
            indicator: subjects.map(s => ({ name: s, max: 120 })), 
            axisName: { color: '#333' },
            splitArea: { areaStyle: { color: ['#fff', '#f8fafc'] } },
            splitLine: { lineStyle: { color: '#cbd5e1' } }
        },
        series: [{
            type: 'radar',
            data: [
                { value: monthlyAvgs, name: '第一次月考' },
                { value: midtermAvgs, name: '期中考试' }
            ]
        }]
    });

    // 3. Class Averages (Right Panel)
    const classChart = echarts.init(document.getElementById('classAvgChart'));
    const classes = data.class_stats.map(c => c.Class_Midterm);
    const classScores = data.class_stats.map(c => c.Avg_Score_Midterm);

    classChart.setOption({
        tooltip: { trigger: 'axis', backgroundColor: 'rgba(255,255,255,0.9)', textStyle: {color: '#333'} },
        xAxis: { 
            type: 'category', 
            data: classes, 
            axisLabel: { color: '#666', rotate: 45, interval: 0 } 
        },
        yAxis: { 
            type: 'value', 
            axisLabel: { color: '#666' }, 
            scale: true,
            splitLine: { lineStyle: { color: '#eee' } }
        },
        series: [{
            type: 'bar',
            data: classScores,
            itemStyle: { color: '#3b82f6' },
            label: { 
                show: true, 
                position: 'top', 
                color: '#333', 
                formatter: (params) => params.value.toFixed(1)
            }
        }]
    });

    // 4. Top/Bottom Improvers (Right Panel)
    renderRankList('topImproversList', data.top_improvers, true);
    renderRankList('bottomImproversList', data.bottom_improvers, false);

    // Resize handler
    window.addEventListener('resize', () => {
        scoreChart.resize();
        subjectChart.resize();
        classChart.resize();
    });
}

function histogram(data, bins) {
    const hist = new Array(bins.length - 1).fill(0);
    if (!data) return hist;
    
    data.forEach(val => {
        for (let i = 0; i < bins.length - 1; i++) {
            if (val >= bins[i] && val < bins[i+1]) {
                hist[i]++;
                break;
            }
        }
    });
    return hist;
}

function renderRankList(elementId, list, isPositive) {
    const ul = document.getElementById(elementId);
    ul.innerHTML = '';
    
    // Check if list is valid
    if (!list || list.length === 0) {
        ul.innerHTML = '<li class="rank-item">暂无数据</li>';
        return;
    }
    
    list.forEach(item => {
        const li = document.createElement('li');
        li.className = 'rank-item';
        // Improvement_School_Rank
        // If isPositive (Top Improvers), we expect positive numbers.
        // If isNegative (Bottom Improvers), we expect negative numbers.
        const change = item.Improvement_School_Rank;
        
        const sign = change > 0 ? '+' : '';
        const colorClass = change > 0 ? 'positive' : (change < 0 ? 'negative' : '');
        
        li.innerHTML = `
            <span>${item.Name_Midterm} <small>(${item.Class_Midterm}班)</small></span>
            <span class="${colorClass}">${sign}${change}</span>
        `;
        ul.appendChild(li);
    });
}

function setupSearch(students) {
    const searchInput = document.getElementById('studentSearch');
    const resultsDiv = document.getElementById('searchResults');

    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.trim();
        if (query.length < 1) {
            resultsDiv.innerHTML = '';
            return;
        }

        const matches = students.filter(s => s.name.includes(query)).slice(0, 10);
        resultsDiv.innerHTML = '';
        
        matches.forEach(student => {
            const div = document.createElement('div');
            div.className = 'search-item';
            div.textContent = `${student.name} (${student.class}班) - 学号:${student.student_id}`;
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

    // Update Text Info
    document.getElementById('studentName').textContent = student.name;
    document.getElementById('studentClass').textContent = `${student.class}班`;
    document.getElementById('monthlyScore').textContent = student.total_score_monthly;
    document.getElementById('midtermScore').textContent = student.total_score_midterm;
    document.getElementById('monthlyRank').textContent = student.total_rank_monthly;
    document.getElementById('midtermRank').textContent = student.total_rank_midterm;
    
    const rankChange = student.rank_change; // Monthly - Midterm. Positive is Improvement.
    const rankEl = document.getElementById('rankChange');
    rankEl.textContent = (rankChange > 0 ? '+' : '') + rankChange;
    rankEl.className = rankChange > 0 ? 'positive' : (rankChange < 0 ? 'negative' : '');

    // Render Student Subject Chart (Combined with Total Rank)
    const chartDom = document.getElementById('studentSubjectChart');
    let myChart = echarts.getInstanceByDom(chartDom);
    if (myChart) myChart.dispose();
    myChart = echarts.init(chartDom);

    // Prepend "Total Rank" to the subjects list
    // Use "总排名" instead of "总分" because mixing Score (e.g. 400) and Rank (e.g. 10) on same Y-axis is confusing.
    // User said "Total Score Bar Chart" but context is "Rank Comparison".
    // If we mix Score (400) and Rank (10), we need dual Y-axis.
    // However, user said "Add Total Score Bar Chart to Single Subject...".
    // Let's assume user wants to compare Ranks (including Total Rank) OR truly mix Score and Rank.
    // Given the previous chart was "Rank Comparison", adding "Total Rank" makes more sense visually.
    // But user explicitly said "Total Score" (总分).
    // Let's try to add "Total Rank" first as it fits the "Rank Comparison" theme.
    // Or if user insists on Score, we need dual axis.
    // Let's use Total Rank as it's comparable with Subject Ranks.
    
    // Wait, user said "Add Total Score Bar Chart...". 
    // If I add Total Score (e.g. 300) next to Ranks (e.g. 50), the scale is fine if ranks are also large.
    // But if Ranks are small (1-10), Score (300) will dwarf them.
    // Let's use "Total Rank" (联考总排名) which is consistent with "Single Subject Rank".
    
    const subjects = ['总排名', ...student.subjects.map(s => s.name)];
    const monthlyData = [student.total_rank_monthly, ...student.subjects.map(s => s.rank_monthly)];
    const midtermData = [student.total_rank_midterm, ...student.subjects.map(s => s.rank_midterm)];
    const changesData = [student.rank_change, ...student.subjects.map(s => s.change)];

    // Prepare series data for bars
    const barSeriesMonthly = {
        name: '第一次月考',
        type: 'bar',
        data: monthlyData,
        itemStyle: { color: '#3b82f6' },
        label: { show: true, position: 'top', color: '#333' }
    };

    const midtermColored = midtermData.map((rank, i) => {
        const change = changesData[i]; // Monthly - Midterm. Pos = Improved.
        const color = change > 0 ? '#ef4444' : (change < 0 ? '#10b981' : '#ccc'); // Red for Rise (Improve), Green for Fall
        return {
            value: rank,
            itemStyle: { color: color }
        };
    });

    const barSeriesMidterm = {
        name: '期中考试 (红升绿降)',
        type: 'bar',
        data: midtermColored,
        // itemStyle: { color: '#10b981' }, // Overridden by data itemStyle
        label: { show: true, position: 'top', color: '#333' }
    };

    const option = {
        tooltip: {
            trigger: 'axis',
            backgroundColor: 'rgba(255,255,255,0.9)',
            textStyle: {color: '#333'},
            axisPointer: { type: 'shadow' },
            formatter: function (params) {
                let res = params[0].name + '<br/>';
                params.forEach(param => {
                    if (param.seriesType === 'bar') {
                        res += param.marker + param.seriesName.split(' ')[0] + ': ' + param.value + '名<br/>';
                    }
                });
                const idx = params[0].dataIndex;
                const change = changesData[idx];
                const label = change > 0 ? '进步' : (change < 0 ? '退步' : '持平');
                const color = change > 0 ? 'red' : (change < 0 ? 'green' : 'gray');
                res += `<span style="color:${color};font-weight:bold">变化: ${label} ${Math.abs(change)} 名</span>`;
                return res;
            }
        },
        legend: {
            data: ['第一次月考', '期中考试 (红升绿降)'],
            textStyle: { color: '#333' }
        },
        grid: {
            left: '3%',
            right: '4%',
            bottom: '15%', // Increased bottom margin for dataZoom
            containLabel: true
        },
        // Add dataZoom for future expandability
        dataZoom: [
            {
                type: 'slider',
                show: true,
                xAxisIndex: [0],
                start: 0,
                end: 100,
                bottom: '2%'
            },
            {
                type: 'inside',
                xAxisIndex: [0],
                start: 0,
                end: 100
            }
        ],
        xAxis: {
            type: 'category',
            data: subjects,
            axisLabel: { color: '#333', interval: 0, fontWeight: 'bold' }
        },
        yAxis: {
            type: 'value',
            inverse: false, // Rank 1 is at bottom
            name: '联考排名 (数值越小越好)',
            nameTextStyle: { color: '#666' },
            axisLabel: { color: '#666' },
            splitLine: { lineStyle: { color: '#eee' } }
        },
        series: [
            barSeriesMonthly,
            barSeriesMidterm
        ]
    };

    myChart.setOption(option);
}
