/* ECharts factory functions for Mannings Dashboard */

const PASTEL_PALETTE = [
  '#5B73E8', '#674EAA', '#43A047', '#E53935', '#F6CF71',
  '#7CB5EC', '#F89C74', '#87C55F', '#E8A44D', '#9B59B6',
];

const SENTIMENT_COLORS = {
  Positive: '#43A047',
  Neutral: '#674EAA',
  Negative: '#E53935',
  'N/A': '#B3B3B3',
};

const chartInstances = {};

function getChart(id) {
  if (chartInstances[id]) {
    chartInstances[id].dispose();
  }
  var el = document.getElementById(id);
  if (!el) return null;
  var chart = echarts.init(el, null, { renderer: 'svg' });
  chartInstances[id] = chart;
  return chart;
}

/* ── Donut with centered total + outside labels ── */

function renderDonut(containerId, data, title) {
  var chart = getChart(containerId);
  if (!chart || !data || data.length === 0) return;

  var total = data.reduce(function(s, d) { return s + d.value; }, 0);

  chart.setOption({
    title: {
      text: total.toLocaleString(),
      left: '50%',
      top: 'center',
      textAlign: 'center',
      textVerticalAlign: 'middle',
      textStyle: {
        fontSize: 22,
        fontWeight: 'bold',
        color: '#18181B',
      },
    },
    tooltip: {
      trigger: 'item',
      formatter: '<b>{b}</b><br/>Count: {c} ({d}%)',
    },
    legend: {
      orient: 'horizontal',
      bottom: 5,
      left: 'center',
      textStyle: { fontSize: 11 },
      itemWidth: 10,
      itemHeight: 10,
      itemGap: 12,
    },
    series: [{
      type: 'pie',
      radius: ['42%', '62%'],
      center: ['50%', '46%'],
      avoidLabelOverlap: true,
      label: {
        show: true,
        position: 'outside',
        formatter: '{b}\n{d}%',
        fontSize: 11,
        color: '#52525B',
      },
      labelLine: {
        show: true,
        length: 8,
        length2: 8,
      },
      data: data.map(function(d, i) {
        return {
          name: d.name,
          value: d.value,
          itemStyle: { color: SENTIMENT_COLORS[d.name] || PASTEL_PALETTE[i % PASTEL_PALETTE.length] },
        };
      }),
    }],
  });
}

function renderDonutGeneric(containerId, data, title) {
  var chart = getChart(containerId);
  if (!chart || !data || data.length === 0) return;

  var total = data.reduce(function(s, d) { return s + d.value; }, 0);

  chart.setOption({
    title: {
      text: total.toLocaleString(),
      left: '50%',
      top: 'center',
      textAlign: 'center',
      textVerticalAlign: 'middle',
      textStyle: {
        fontSize: 20,
        fontWeight: 'bold',
        color: '#18181B',
      },
    },
    tooltip: {
      trigger: 'item',
      formatter: '<b>{b}</b><br/>Count: {c} ({d}%)',
    },
    legend: {
      orient: 'horizontal',
      bottom: 5,
      left: 'center',
      textStyle: { fontSize: 11 },
      itemWidth: 10,
      itemHeight: 10,
      itemGap: 12,
    },
    series: [{
      type: 'pie',
      radius: ['42%', '62%'],
      center: ['50%', '46%'],
      avoidLabelOverlap: true,
      label: {
        show: true,
        position: 'outside',
        formatter: '{b}\n{d}%',
        fontSize: 11,
        color: '#52525B',
      },
      labelLine: {
        show: true,
        length: 8,
        length2: 8,
      },
      data: data.map(function(d, i) {
        return {
          name: d.name,
          value: d.value,
          itemStyle: { color: PASTEL_PALETTE[i % PASTEL_PALETTE.length] },
        };
      }),
    }],
  });
}

/* ── Bar Chart ── */

function renderBarChart(containerId, data, yLabel) {
  var chart = getChart(containerId);
  if (!chart || !data || data.length === 0) return;

  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      formatter: function(params) {
        var p = params[0];
        return '<b>' + p.name + '</b><br/>' + (yLabel || 'Value') + ': ' + p.value.toLocaleString();
      },
    },
    grid: { left: 60, right: 30, top: 20, bottom: 80 },
    xAxis: {
      type: 'category',
      data: data.map(function(d) { return d.name; }),
      axisLabel: {
        rotate: data.length > 5 ? 30 : 0,
        fontSize: 10,
        color: '#6E6F75',
        fontWeight: 450,
        interval: 0,
      },
      axisLine: { lineStyle: { color: '#E4E4E7' } },
    },
    yAxis: {
      type: 'value',
      axisLabel: {
        fontSize: 10,
        color: '#6E6F75',
        fontWeight: 450,
        formatter: function(val) {
          if (val >= 1000) return (val / 1000).toFixed(0) + 'k';
          return val;
        },
      },
      splitLine: { lineStyle: { color: '#F0F0F3' } },
    },
    series: [{
      type: 'bar',
      data: data.map(function(d, i) {
        return {
          value: d.value,
          itemStyle: { color: PASTEL_PALETTE[i % PASTEL_PALETTE.length] },
        };
      }),
      barWidth: '50%',
      label: {
        show: true,
        position: 'top',
        formatter: function(p) {
          return p.value.toLocaleString();
        },
        fontSize: 10,
        fontWeight: 600,
        color: '#18181B',
      },
    }],
  });
}

/* ── Sentiment Dashboard: aligned stacked bar + table combo ──
 *
 * Chart grid boundaries match table column boundaries so each bar
 * center-aligns with its corresponding table column.
 *
 * Table uses .sentiment-combo-table class (matches dash-table style).
 */
function renderSentimentDashboard(containerId, data, groupKey) {
  var chartEl = document.getElementById(containerId);
  if (!chartEl || !data || data.length === 0) return;

  var sentiments = ['Positive', 'Neutral', 'Negative'];
  var categories = data.map(function(d) { return d[groupKey]; });
  var n = categories.length;
  var totalCols = n + 2;  // +Sentiment label + Grand Total
  var colPercent = 100 / totalCols;

  // Remove old table if re-rendering
  var oldTable = document.getElementById(containerId + '-table');
  if (oldTable) oldTable.remove();

  // Totals
  var totalRow = {};
  sentiments.forEach(function(s) {
    totalRow[s] = data.reduce(function(sum, d) { return sum + (d[s] || 0); }, 0);
  });
  var grandTotal = sentiments.reduce(function(sum, s) { return sum + totalRow[s]; }, 0);

  // Chart — grid left/right match first and last column widths
  var chart = getChart(containerId);
  chart.setOption({
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      formatter: function(params) {
        var cat = params[0].name;
        var html = '<b>' + cat + '</b>';
        params.forEach(function(p) {
          html += '<br/><span style="color:' + p.color + '">&#9679;</span> ' + p.seriesName + ': ' + p.value;
        });
        return html;
      },
    },
    legend: {
      top: 0,
      right: 0,
      textStyle: { fontSize: 11 },
      data: sentiments.slice().reverse(),
    },
    grid: {
      left: colPercent + '%',
      right: colPercent + '%',
      top: 40,
      bottom: 10,
      containLabel: false,
    },
    xAxis: {
      type: 'category',
      data: categories,
      show: false,
    },
    yAxis: {
      type: 'value',
      axisLabel: { fontSize: 10, color: '#6E6F75', fontWeight: 450 },
      splitLine: { lineStyle: { color: '#F0F0F3', type: 'dashed' } },
      axisLine: { show: false },
      axisTick: { show: false },
    },
    series: sentiments.slice().reverse().map(function(s) {
      return {
        name: s,
        type: 'bar',
        stack: 'total',
        data: data.map(function(d) { return d[s] || 0; }),
        itemStyle: { color: SENTIMENT_COLORS[s] },
        emphasis: { focus: 'series' },
        barWidth: '60%',
      };
    }),
  });

  // Build aligned table with .sentiment-combo-table class
  var wrap = document.createElement('div');
  wrap.className = 'chart-table-combo';
  wrap.id = containerId + '-table';

  var table = document.createElement('table');
  table.className = 'sentiment-combo-table';
  table.style.tableLayout = 'fixed';

  // Header
  var thead = document.createElement('thead');
  var headRow = document.createElement('tr');
  var thFirst = document.createElement('th');
  thFirst.textContent = 'Sentiment';
  headRow.appendChild(thFirst);
  categories.forEach(function(cat) {
    var th = document.createElement('th');
    th.textContent = cat;
    headRow.appendChild(th);
  });
  var thTotal = document.createElement('th');
  thTotal.textContent = 'Grand Total';
  headRow.appendChild(thTotal);
  thead.appendChild(headRow);
  table.appendChild(thead);

  // Body — one row per sentiment
  var tbody = document.createElement('tbody');
  sentiments.forEach(function(s) {
    var tr = document.createElement('tr');
    var thLabel = document.createElement('th');
    thLabel.textContent = s;
    tr.appendChild(thLabel);
    categories.forEach(function(cat) {
      var td = document.createElement('td');
      var item = data.find(function(d) { return d[groupKey] === cat; });
      td.textContent = item && item[s] ? item[s] : '0';
      tr.appendChild(td);
    });
    var tdTotal = document.createElement('td');
    tdTotal.textContent = totalRow[s];
    tr.appendChild(tdTotal);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  // Footer — Grand Total row
  var tfoot = document.createElement('tfoot');
  var footRow = document.createElement('tr');
  var footLabel = document.createElement('td');
  footLabel.textContent = 'Grand Total';
  footRow.appendChild(footLabel);
  categories.forEach(function(cat) {
    var td = document.createElement('td');
    var item = data.find(function(d) { return d[groupKey] === cat; });
    var colTotal = sentiments.reduce(function(sum, s) { return sum + (item && item[s] ? item[s] : 0); }, 0);
    td.textContent = colTotal;
    footRow.appendChild(td);
  });
  var footTotal = document.createElement('td');
  footTotal.textContent = grandTotal;
  footRow.appendChild(footTotal);
  tfoot.appendChild(footRow);
  table.appendChild(tfoot);

  wrap.appendChild(table);
  chartEl.parentNode.insertBefore(wrap, chartEl.nextSibling);
}

/* ── Date formatter: converts various date formats to "d Mmm yy" ── */
function fmtDate(d) {
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var dt;
  if (d instanceof Date) {
    dt = d;
  } else {
    var s = String(d).trim();
    // dd/mm/yyyy
    var m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    // yyyy-mm-dd
    var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m1) {
      dt = new Date(parseInt(m1[3]), parseInt(m1[2]) - 1, parseInt(m1[1]));
    } else if (m2) {
      dt = new Date(parseInt(m2[1]), parseInt(m2[2]) - 1, parseInt(m2[3]));
    } else {
      dt = new Date(s);
    }
  }
  if (isNaN(dt)) return s;
  return dt.getDate() + ' ' + months[dt.getMonth()] + ' ' + String(dt.getFullYear()).slice(2);
}

/* ── Followers Growth — Combined Bar + Line ──
 * Gain (faded green bars), Loss (faded red bars) side by side
 * Net (purple line) centered on bar pair
 */
function renderFollowersGrowth(containerId, data) {
  var chart = getChart(containerId);
  if (!chart || !data || !data.dates || data.dates.length === 0) return;

  var xLabels = data.dates.map(function(d) { return fmtDate(d); });

  chart.setOption({
    title: {
      text: data.monthly_net.toLocaleString(),
      right: 10,
      top: 4,
      textStyle: {
        fontSize: 13,
        fontWeight: 'bold',
        color: data.monthly_net >= 0 ? '#5B73E8' : '#E53935',
      },
    },
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'cross' },
      formatter: function(params) {
        var html = '<b>' + params[0].axisValue + '</b>';
        params.forEach(function(p) {
          html += '<br/><span style="color:' + p.color + '">&#9679;</span> ' + p.seriesName + ': ' + p.value.toLocaleString();
        });
        return html;
      },
    },
    legend: {
      bottom: 0,
      textStyle: { fontSize: 10 },
      itemWidth: 12,
      itemHeight: 8,
    },
    grid: { left: 60, right: 30, top: 40, bottom: 60 },
    xAxis: {
      type: 'category',
      data: xLabels,
      axisLabel: {
        rotate: 45,
        fontSize: 9,
        color: '#6E6F75',
        fontWeight: 450,
        interval: Math.floor(xLabels.length / 12),
      },
    },
    yAxis: {
      type: 'value',
      axisLabel: { fontSize: 10, color: '#6E6F75', fontWeight: 450 },
      splitLine: { lineStyle: { color: '#F0F0F3' } },
    },
    series: [
      {
        name: 'Gain',
        type: 'bar',
        data: data.gain,
        itemStyle: { color: 'rgba(76, 175, 80, 0.55)' },
        barGap: '-100%',
        barCategoryGap: '60%',
      },
      {
        name: 'Loss',
        type: 'bar',
        data: data.loss,
        itemStyle: { color: 'rgba(244, 67, 54, 0.55)' },
      },
      {
        name: 'Net',
        type: 'line',
        data: data.net,
        itemStyle: { color: '#5B73E8' },
        lineStyle: { color: '#5B73E8', width: 2 },
        symbol: 'circle',
        symbolSize: 3,
        smooth: true,
      },
    ],
  });
}

/* ── Total Reach — Line Chart ──
 * Orange line (Mannings color), values in millions.
 */
function renderTotalReach(containerId, data) {
  var chart = getChart(containerId);
  if (!chart || !data || !data.dates || data.dates.length === 0) return;

  function fmtM(v) {
    if (v >= 1e6) return (v / 1e6).toFixed(2) + ' M';
    if (v >= 1e3) return (v / 1e3).toFixed(0) + 'k';
    return v;
  }

  var xLabels = data.dates.map(function(d) { return fmtDate(d); });

  chart.setOption({
    title: {
      text: data.monthly_total.toLocaleString(),
      right: 10,
      top: 4,
      textStyle: {
        fontSize: 13,
        fontWeight: 'bold',
        color: '#FE8301',
      },
    },
    tooltip: {
      trigger: 'axis',
      formatter: function(params) {
        var p = params[0];
        return '<b>' + p.axisValue + '</b><br/>Total Reach: ' + p.value.toLocaleString();
      },
    },
    grid: { left: 60, right: 30, top: 40, bottom: 60 },
    xAxis: {
      type: 'category',
      boundaryGap: false,
      data: xLabels,
      axisLabel: {
        rotate: 45,
        fontSize: 9,
        color: '#6E6F75',
        fontWeight: 450,
        interval: Math.floor(xLabels.length / 12),
      },
    },
    yAxis: {
      type: 'value',
      axisLabel: {
        fontSize: 10,
        color: '#6E6F75',
        fontWeight: 450,
        formatter: function(v) { return fmtM(v); },
      },
      splitLine: { lineStyle: { color: '#F0F0F3' } },
    },
    series: [{
      name: 'Total Reach',
      type: 'line',
      data: data.reach,
      itemStyle: { color: '#FE8301' },
      lineStyle: { color: '#FE8301', width: 2.5 },
      symbol: 'circle',
      symbolSize: 3,
      smooth: true,
      areaStyle: {
        color: {
          type: 'linear',
          x: 0, y: 0, x2: 0, y2: 1,
          colorStops: [
            { offset: 0, color: 'rgba(254, 131, 1, 0.2)' },
            { offset: 1, color: 'rgba(254, 131, 1, 0.02)' },
          ],
        },
      },
    }],
  });
}

/* ── Organic Reach Funnel ──
 * Organic (orange, top), Paid (green, bottom)
 */
function renderReachFunnel(containerId, data) {
  var chart = getChart(containerId);
  if (!chart || !data || !data.total) return;

  function fmtM(v) {
    if (v >= 1e6) return (v / 1e6).toFixed(1) + ' M';
    if (v >= 1e3) return (v / 1e3).toFixed(0) + 'k';
    return v;
  }

  var organicPct = data.total > 0 ? (data.organic / data.total * 100) : 0;
  var paidPct = 100 - organicPct;

  chart.setOption({
    tooltip: {
      trigger: 'item',
      formatter: '<b>{b}</b><br/>Reach: {c}',
    },
    series: [{
      type: 'funnel',
      left: '10%',
      right: '10%',
      top: 10,
      bottom: 10,
      minSize: '15%',
      maxSize: '100%',
      gap: 2,
      label: {
        show: true,
        position: 'inside',
        formatter: function(p) {
          return p.name + '\n' + fmtM(p.value);
        },
        fontSize: 13,
        fontWeight: 'bold',
        color: '#fff',
      },
      data: [
        { value: data.organic, name: 'Organic Reach', itemStyle: { color: '#FE8301' } },
        { value: data.paid, name: 'Paid Reach', itemStyle: { color: '#43A047' } },
      ],
    }],
  });
}

/* ── IG Followers Growth — Dual Line (Total + Net) ── */
function renderIgFollowers(containerId, data) {
  var chart = getChart(containerId);
  if (!chart || !data || !data.dates || data.dates.length === 0) return;

  var xLabels = data.dates.map(function(d) { return fmtDate(d); });
  var hasTotal = data.total && data.total.length > 0;

  var yAxis = [{
    type: 'value',
    name: 'Total',
    nameTextStyle: { fontSize: 10, color: '#6E6F75' },
    position: 'left',
    axisLabel: { fontSize: 10, color: '#6E6F75', fontWeight: 450 },
    splitLine: { lineStyle: { color: '#F0F0F3' } },
  }, {
    type: 'value',
    name: 'Net',
    nameTextStyle: { fontSize: 10, color: '#6E6F75' },
    position: 'right',
    max: 1000,
    axisLabel: { fontSize: 10, color: '#6E6F75', fontWeight: 450 },
    splitLine: { show: false },
  }];

  var series = [];
  if (hasTotal) {
    series.push({
      name: 'Total Followers',
      type: 'line',
      yAxisIndex: 0,
      data: data.total,
      itemStyle: { color: '#FE8301' },
      lineStyle: { color: '#FE8301', width: 2.5 },
      symbol: 'circle',
      symbolSize: 3,
      smooth: true,
      areaStyle: {
        color: {
          type: 'linear', x: 0, y: 0, x2: 0, y2: 1,
          colorStops: [
            { offset: 0, color: 'rgba(254, 131, 1, 0.2)' },
            { offset: 1, color: 'rgba(254, 131, 1, 0.02)' },
          ],
        },
      },
    });
  }
  series.push({
    name: 'Net',
    type: 'line',
    yAxisIndex: 1,
    data: data.net,
    itemStyle: { color: '#5B73E8' },
    lineStyle: { color: '#5B73E8', width: 2 },
    symbol: 'circle',
    symbolSize: 3,
    smooth: true,
  });

  chart.setOption({
    title: {
      text: data.monthly_net.toLocaleString(),
      right: 10,
      top: 4,
      textStyle: { fontSize: 13, fontWeight: 'bold', color: '#5B73E8' },
    },
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'cross' },
      formatter: function(params) {
        var html = '<b>' + params[0].axisValue + '</b>';
        params.forEach(function(p) {
          html += '<br/><span style="color:' + p.color + '">&#9679;</span> ' + p.seriesName + ': ' + p.value.toLocaleString();
        });
        return html;
      },
    },
    legend: {
      bottom: 0,
      textStyle: { fontSize: 10 },
      itemWidth: 12,
      itemHeight: 8,
    },
    grid: { left: 60, right: 60, top: 40, bottom: 60 },
    xAxis: {
      type: 'category',
      boundaryGap: false,
      data: xLabels,
      axisLabel: {
        rotate: 45,
        fontSize: 9,
        color: '#6E6F75',
        fontWeight: 450,
        interval: Math.floor(xLabels.length / 12),
      },
    },
    yAxis: yAxis,
    series: series,
  });
}

/* ── Competitor Fans Growth Trend — Line Chart ── */
function renderGrowthTrendChart(containerId, data) {
  var chart = getChart(containerId);
  if (!chart || !data || !data.series || data.series.length === 0) return;

  var xLabels = data.dates.map(function(d) { return fmtDate(d); });

  chart.setOption({
    tooltip: { trigger: 'axis' },
    legend: {
      bottom: 0,
      textStyle: { fontSize: 10 },
      itemWidth: 12,
      itemHeight: 8,
    },
    grid: { left: 70, right: 30, top: 20, bottom: 70 },
    xAxis: {
      type: 'category',
      boundaryGap: false,
      data: xLabels,
      axisLabel: { rotate: 45, fontSize: 9, color: '#6E6F75', fontWeight: 450, interval: Math.floor(xLabels.length / 10) },
    },
    yAxis: {
      type: 'value',
      axisLabel: {
        fontSize: 10,
        color: '#6E6F75',
        fontWeight: 450,
        formatter: function(v) { return v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v; },
      },
      splitLine: { lineStyle: { color: '#F0F0F3' } },
    },
    series: data.series.map(function(s) {
      return {
        name: s.name,
        type: 'line',
        data: s.data,
        itemStyle: { color: s.color },
        lineStyle: { color: s.color, width: 2 },
        symbol: 'circle',
        symbolSize: 4,
      };
    }),
  });
}

/* ── KPI Comparison — Bubble/Scatter with logos ── */
function renderKpiBubbleChart(containerId, data) {
  var chart = getChart(containerId);
  if (!chart || !data || !data.items || data.items.length === 0) return;

  chart.setOption({
    tooltip: {
      formatter: function(p) {
        return '<b>' + p.data.name + '</b><br/>Posts: ' + p.data.value[0] + '<br/>Reactions, Comments & Shares: ' + p.data.value[1];
      },
    },
    grid: { left: 100, right: 50, top: 20, bottom: 60 },
    xAxis: {
      type: 'value',
      name: 'Number of Posts',
      nameLocation: 'middle',
      nameGap: 35,
      nameTextStyle: { fontSize: 11, color: '#6E6F75', fontWeight: 450 },
      axisLabel: { color: '#6E6F75', fontWeight: 450 },
    },
    yAxis: {
      type: 'value',
      name: 'Reactions, Comments & Shares',
      nameLocation: 'middle',
      nameGap: 60,
      nameTextStyle: { fontSize: 11, color: '#6E6F75', fontWeight: 450 },
      axisLabel: {
        color: '#6E6F75',
        fontWeight: 450,
        formatter: function(v) { return v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v; },
      },
    },
    series: [{
      type: 'scatter',
      data: data.items.map(function(item) {
        return {
          value: [item.posts, item.reactions],
          name: item.name,
          symbol: 'image://' + item.logo,
          symbolSize: 36,
        };
      }),
      markLine: {
        silent: true,
        symbol: 'none',
        data: [
          { xAxis: data.avg_posts, lineStyle: { type: 'dashed', color: '#999' }, label: { show: true, formatter: 'Avg Posts: ' + data.avg_posts, fontSize: 10, position: 'insideEndTop' } },
          { yAxis: data.avg_reactions, lineStyle: { type: 'solid', color: '#999' }, label: { show: true, formatter: 'Avg Reactions: ' + data.avg_reactions, fontSize: 10, position: 'insideEndTop' } },
        ],
      },
    }],
  });
}

/* ── Top 50 Words — Tag Cloud ── */
function renderTagCloud(containerId, data) {
  var chart = getChart(containerId);
  if (!chart || !data || data.length === 0) return;

  var times = data.map(function(d) { return d.times_above_avg; });
  var minT = Math.min.apply(null, times);
  var maxT = Math.max.apply(null, times);

  function tagColor(t) {
    var ratio = maxT === minT ? 0.5 : (t - minT) / (maxT - minT);
    if (ratio < 0.5) {
      var k = ratio * 2;
      return 'rgb(' + Math.round(244 + (255 - 244) * k) + ',' + Math.round(67 + (193 - 67) * k) + ',' + Math.round(36 + (7 - 36) * k) + ')';
    }
    var k2 = (ratio - 0.5) * 2;
    return 'rgb(' + Math.round(255 + (76 - 255) * k2) + ',' + Math.round(193 + (175 - 193) * k2) + ',' + Math.round(7 + (80 - 7) * k2) + ')';
  }

  var values = data.map(function(d) { return d.value; });
  var minV = Math.min.apply(null, values);
  var maxV = Math.max.apply(null, values);

  chart.setOption({
    tooltip: {
      formatter: function(p) {
        return p.name + ': ' + p.value + ' (×' + p.data.times + ')';
      },
    },
    series: [{
      type: 'wordCloud',
      shape: 'circle',
      left: 'center',
      top: 'center',
      width: '95%',
      height: '95%',
      sizeRange: [14, 52],
      rotationRange: [0, 0],
      rotationStep: 0,
      gridSize: 6,
      drawOutOfBound: false,
      layoutAnimation: true,
      textStyle: {
        fontFamily: 'Inter, "Noto Sans TC", sans-serif',
        fontWeight: 'bold',
      },
      emphasis: { textStyle: { shadowBlur: 10, shadowColor: 'rgba(0,0,0,0.3)' } },
      data: data.map(function(d) {
        return {
          name: d.word,
          value: d.value,
          times: d.times_above_avg,
          textStyle: { color: tagColor(d.times_above_avg) },
        };
      }).sort(function(a, b) {
        return b.value - a.value;
      }),
    }],
  });
}

window.chartInstances = chartInstances;

window.addEventListener('resize', function() {
  Object.keys(chartInstances).forEach(function(id) {
    var c = chartInstances[id];
    if (c && !c.isDisposed()) c.resize();
  });
});
