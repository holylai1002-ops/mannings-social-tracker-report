/* Mannings Dashboard v2 — Alpine.js controller (Calendar Picker) */

function dashboard(initial) {
  return {
    // ── State ──
    page: initial.page,
    startDate: initial.startDate,
    endDate: initial.endDate,
    tempStart: initial.startDate,
    tempEnd: initial.endDate,
    dateError: '',

    loading: false,

    // ── Data ──
    kpis: [],
    sentimentDist: [],
    sentimentByCategory: [],
    sentimentByType: [],
    wallPosts: [],
    wallPostCols: [],
    sortState: {},

    pillarDonut: [],
    typeDonut: [],
    pillarInteractions: [],
    catInteractions: [],
    breakdowns: {},

    fbKeyMetrics: [],
    igKeyMetrics: [],
    followersGrowth: null,
    totalReach: null,
    reachFunnel: null,
    igFollowers: null,
    igPillarDonut: [],
    igPillarInteractions: [],
    igEngagementDonut: [],
    igStoryCatDonut: [],
    igStoryClicks: [],
    igWallPosts: [],
    igStoryPosts: [],

    fpkGrowthTrend: null,
    fpkKpiComparison: null,
    fpkCompetitorsOverview: [],
    fpkTopWords: [],

    // ── Chat ──
    chatOpen: false,
    chatInput: '',
    chatHistory: [],
    chatStreaming: false,
    suggestedPrompts: [
      'What are the key highlights for this period?',
      'Which pillar performed best and why?',
      'How is the sentiment distribution?',
      'What should we improve next period?',
    ],

    // ── AI Insights ──
    aiInsightsBusy: false,
    insightsText: {},
    insightsHtml: {},
    insightsLoading: {},

    // ── Computed ──
    get periodLabel() {
      return `${this.startDate} ~ ${this.endDate}`;
    },

    get pageTitles() {
      return {
        fb_page: 'FB Page',
        fb_posts: 'FB Posts',
        instagram: 'Instagram',
      };
    },

    // ── Generic Table Sorting ──
    sortTable(tableKey, col, event) {
      if (!this.sortState[tableKey]) this.sortState[tableKey] = {};
      if (this.sortState[tableKey].col === col) {
        this.sortState[tableKey].asc = !this.sortState[tableKey].asc;
      } else {
        this.sortState[tableKey] = { col, asc: true };
      }
    },

    getSortedData(arr, tableKey) {
      if (!arr || arr.length === 0) return arr;
      const st = this.sortState[tableKey];
      if (!st || !st.col) return arr;
      const col = st.col;
      const asc = st.asc;
      return [...arr].sort((a, b) => {
        let av = a[col], bv = b[col];
        if (typeof av === 'number' && typeof bv === 'number') {
          return asc ? av - bv : bv - av;
        }
        av = String(av || ''); bv = String(bv || '');
        return asc ? av.localeCompare(bv) : bv.localeCompare(av);
      });
    },

    sortIcon(tableKey, col) {
      const st = this.sortState[tableKey];
      if (!st || st.col !== col) return '';
      return st.asc ? ' \u25B2' : ' \u25BC';
    },

    initDefaultSort(tableKey, data) {
      if (this.sortState[tableKey]) return;
      if (!data || data.length === 0) return;
      const keys = Object.keys(data[0]);
      const dateCol = keys.find(k => ['Publish time', 'Post Date', 'Comment Date', 'Publish Time'].includes(k));
      if (dateCol) {
        this.sortState[tableKey] = { col: dateCol, asc: true };
      }
    },

    sortByValue(arr) {
      return [...arr].sort((a, b) => (b.value || 0) - (a.value || 0));
    },

    sortSentimentByTotal(arr) {
      return [...arr].sort((a, b) => {
        const ta = (a.Positive || 0) + (a.Neutral || 0) + (a.Negative || 0);
        const tb = (b.Positive || 0) + (b.Neutral || 0) + (b.Negative || 0);
        return tb - ta;
      });
    },

    // ── Init ──
    init() {
      this.loadAll();
    },

    // ── Preset Buttons ──
    setPreset(preset) {
      const today = new Date();
      let s, e;
      if (preset === 'lastMonth') {
        const d = new Date(today.getFullYear(), today.getMonth() - 1, 1);
        s = new Date(d.getFullYear(), d.getMonth(), 1);
        e = new Date(d.getFullYear(), d.getMonth() + 1, 0);
      } else if (preset === 'last7Days') {
        e = new Date(today);
        s = new Date(today);
        s.setDate(s.getDate() - 6);
      } else if (preset === 'thisYear') {
        s = new Date(today.getFullYear(), 0, 1);
        e = new Date(today);
      }
      this.startDate = this.fmtDate(s);
      this.endDate = this.fmtDate(e);
      this.tempStart = this.startDate;
      this.tempEnd = this.endDate;
      this.dateError = '';
      this.onRangeChange();
    },

    fmtDate(d) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    },

    get maxDate() {
      return this.fmtDate(new Date());
    },

    applyDateRange() {
      if (!this.tempStart || !this.tempEnd) {
        this.dateError = 'Please select both dates.';
        return;
      }
      if (this.tempStart > this.tempEnd) {
        this.dateError = 'Start date cannot be after end date.';
        return;
      }
      this.dateError = '';
      this.startDate = this.tempStart;
      this.endDate = this.tempEnd;
      this.onRangeChange();
    },

    // ── Range Change ──
    onRangeChange() {
      this.insightsText = {};
      this.insightsHtml = {};
      this.insightsLoading = {};
      this.sortState = {};
      this.loadAll();
      const url = new URL(window.location);
      url.searchParams.set('start', this.startDate);
      url.searchParams.set('end', this.endDate);
      window.history.replaceState({}, '', url);
    },

    // ── Load All Data ──
    async loadAll() {
      this.loading = true;
      try {
        await Promise.all([
          this.loadKPIs(),
          this.loadPageData(),
        ]);
        if (this.page === 'fb_page' || this.page === 'instagram') {
          await this.loadFPK();
        }
      } catch (e) {
        console.error('Load error:', e);
      }
      this.loading = false;
      this.$nextTick(() => {
        this.renderCharts();
        this.initColumnResize();
        setTimeout(() => {
          this.renderCharts();
          this.initColumnResize();
          Object.values(window.chartInstances || {}).forEach(function(c) {
            if (c && !c.isDisposed()) c.resize();
          });
        }, 150);
      });
    },

    async loadKPIs() {
      try {
        const res = await fetch(`/api/data/kpis?start=${this.startDate}&end=${this.endDate}`);
        const data = await res.json();
        let kpis = [];
        if (this.page === 'fb_page' || this.page === 'fb_posts') {
          kpis = [
            { label: 'Total Page Follows', value: this.fmt(data.fb_followers) },
            { label: 'FB Page Follows Growth', value: this.fmt(data.fb_growth) },
            { label: 'No. of Wall Post', value: this.fmt(data.fb_wall_posts) },
            { label: 'Total Interactions', value: this.fmt(data.fb_interactions) },
          ];
        } else {
          kpis = [
            { label: 'Instagram Followers', value: this.fmt(data.ig_followers) },
            { label: 'Instagram Growth', value: this.fmt(data.ig_growth) },
            { label: 'Total Reach', value: this.fmt(data.ig_reach) },
            { label: 'Total Interactions', value: this.fmt(data.ig_interactions) },
          ];
        }
        this.kpis = kpis;
      } catch (e) { console.error('KPI load error:', e); }
    },

    async loadPageData() {
      const pageMap = {
        fb_page: 'fb_page',
        fb_posts: 'fb_posts',
        instagram: 'instagram',
      };
      const endpoint = pageMap[this.page];
      try {
        const res = await fetch(`/api/data/${endpoint}?start=${this.startDate}&end=${this.endDate}`);
        const data = await res.json();

        if (this.page === 'fb_page') {
          this.sentimentDist = data.sentiment_distribution || [];
          this.sentimentByCategory = this.sortSentimentByTotal(data.sentiment_by_category || []);
          this.sentimentByType = this.sortSentimentByTotal(data.sentiment_by_type || []);
          this.fbKeyMetrics = data.fb_key_metrics || [];
          this.followersGrowth = data.followers_growth || null;
          this.totalReach = data.total_reach || null;
          this.reachFunnel = data.reach_funnel || null;
        } else if (this.page === 'fb_posts') {
          this.pillarDonut = data.pillar_donut || [];
          this.typeDonut = data.type_donut || [];
          this.pillarInteractions = this.sortByValue(data.pillar_interactions || []);
          this.catInteractions = this.sortByValue(data.cat_interactions || []);
          this.breakdowns = data.breakdowns || {};
          this.wallPosts = data.wall_posts || [];
          this.wallPostCols = this.wallPosts.length > 0 ? Object.keys(this.wallPosts[0]).filter(k => !k.startsWith('_')) : [];
          this.initDefaultSort('fbWall', this.wallPosts);
        } else if (this.page === 'instagram') {
          this.igKeyMetrics = data.ig_key_metrics || [];
          this.igFollowers = data.ig_followers || null;
          this.igPillarDonut = data.pillar_donut || [];
          this.igPillarInteractions = this.sortByValue(data.pillar_interactions || []);
          this.igEngagementDonut = data.engagement_donut || [];
          this.igStoryCatDonut = data.story_cat_donut || [];
          this.igStoryClicks = this.sortByValue(data.story_clicks || []);
          this.igWallPosts = data.wall_posts || [];
          this.igStoryPosts = data.ig_story_posts || [];
          this.initDefaultSort('igWall', this.igWallPosts);
          this.initDefaultSort('igStory', this.igStoryPosts);
        }
      } catch (e) { console.error('Page data load error:', e); }
    },

    async loadFPK() {
      const platform = this.page === 'instagram' ? 'ig' : 'fb';
      try {
        const res = await fetch(`/api/data/fpk?start=${this.startDate}&end=${this.endDate}&platform=${platform}`);
        const data = await res.json();
        this.fpkGrowthTrend = data.growth_trend || null;
        this.fpkKpiComparison = data.kpi_comparison || null;
        this.fpkCompetitorsOverview = data.competitors_overview || [];
        this.fpkTopWords = data.top_words || [];
      } catch (e) { console.error('FPK load error:', e); }
    },

    // ── Chart Rendering ──
    renderCharts() {
      if (this.page === 'fb_page') {
        if (this.sentimentDist.length > 0) {
          renderDonut('chart-sentiment-donut', this.sentimentDist, 'Sentiment');
        }
        if (this.sentimentByCategory.length > 0) {
          renderSentimentDashboard('chart-sentiment-category', this.sentimentByCategory, 'category');
        }
        if (this.sentimentByType.length > 0) {
          renderSentimentDashboard('chart-sentiment-type', this.sentimentByType, 'type');
        }
        if (this.followersGrowth) {
          renderFollowersGrowth('chart-followers-growth', this.followersGrowth);
        }
        if (this.totalReach) {
          renderTotalReach('chart-total-reach', this.totalReach);
        }
        if (this.reachFunnel) {
          renderReachFunnel('chart-reach-funnel', this.reachFunnel);
        }
      } else if (this.page === 'fb_posts') {
        if (this.pillarDonut.length > 0) {
          renderDonutGeneric('chart-pillar-donut', this.pillarDonut, 'Pillar');
        }
        if (this.typeDonut.length > 0) {
          renderDonutGeneric('chart-type-donut', this.typeDonut, 'Type');
        }
        if (this.pillarInteractions.length > 0) {
          renderBarChart('chart-pillar-bar', this.pillarInteractions, 'Interactions');
        }
        if (this.catInteractions.length > 0) {
          renderBarChart('chart-category-bar', this.catInteractions, 'Interactions');
        }
      } else if (this.page === 'instagram') {
        if (this.igFollowers) {
          renderIgFollowers('chart-ig-followers', this.igFollowers);
        }
        if (this.igPillarDonut.length > 0) {
          renderDonutGeneric('chart-ig-pillar-donut', this.igPillarDonut, 'Pillar');
        }
        if (this.igEngagementDonut.length > 0) {
          renderDonutGeneric('chart-ig-engagement', this.igEngagementDonut, 'Engagement');
        }
        if (this.igPillarInteractions.length > 0) {
          renderBarChart('chart-ig-pillar-bar', this.igPillarInteractions, 'Interactions');
        }
        if (this.igStoryCatDonut.length > 0) {
          renderDonutGeneric('chart-ig-story-cat', this.igStoryCatDonut, 'Category');
        }
        if (this.igStoryClicks.length > 0) {
          renderBarChart('chart-ig-story-clicks', this.igStoryClicks, 'Link Clicks');
        }
      }

      if (this.fpkGrowthTrend) {
        renderGrowthTrendChart('chart-fpk-growth', this.fpkGrowthTrend);
      }
      if (this.fpkKpiComparison) {
        renderKpiBubbleChart('chart-fpk-bubble', this.fpkKpiComparison);
      }
      if (this.fpkTopWords.length > 0) {
        this.$nextTick(function() { renderTagCloud('chart-fpk-words', this.fpkTopWords); }.bind(this));
      }
    },

    // ── Table Helpers ──

    formatCell(val) {
      if (val === null || val === undefined || val === '') return '';
      if (typeof val === 'number') return val.toLocaleString();
      return String(val);
    },

    cellClass(col) {
      var c = String(col).toLowerCase();
      if (c === 'permalink' || c === 'link') return 'cell-permalink';
      if (c === 'description') return 'cell-description';
      return '';
    },

    initColumnResize() {
      var tables = document.querySelectorAll('.dash-table');
      tables.forEach(function(table) {
        var ths = table.querySelectorAll('thead th');
        if (!ths || ths.length === 0) return;
        if (table.dataset.colResizeInit === '1') return;
        table.dataset.colResizeInit = '1';
        table.style.tableLayout = 'fixed';

        var measurer = document.createElement('span');
        measurer.style.cssText = 'position:absolute;visibility:hidden;white-space:nowrap;font-size:12px;font-weight:600;font-family:Inter,"Noto Sans TC",sans-serif;';
        document.body.appendChild(measurer);

        ths.forEach(function(th, idx) {
          measurer.textContent = th.textContent || th.innerText || '';
          var headerWidth = measurer.offsetWidth + 32;
          th.style.width = headerWidth + 'px';

          var handle = document.createElement('div');
          handle.className = 'col-resizer';
          th.appendChild(handle);

          handle.addEventListener('mousedown', function(e) {
            e.preventDefault();
            e.stopPropagation();
            var startX = e.pageX;
            var startW = th.offsetWidth;
            var body = document.body;
            body.style.cursor = 'col-resize';
            body.style.userSelect = 'none';

            function onMove(ev) {
              var newW = Math.max(40, startW + (ev.pageX - startX));
              th.style.width = newW + 'px';
            }
            function onUp() {
              body.style.cursor = '';
              body.style.userSelect = '';
              document.removeEventListener('mousemove', onMove);
              document.removeEventListener('mouseup', onUp);
            }
            document.addEventListener('mousemove', onMove);
            document.addEventListener('mouseup', onUp);
          });
        });

        document.body.removeChild(measurer);
      });
    },

    fmt(val) {
      if (val === null || val === undefined || 0 === val) return '0';
      if (typeof val === 'number') return val.toLocaleString();
      return val;
    },

    // ── Print PDF ──
    printPDF() {
      const origTitle = document.title;
      document.title = `Mannings_${this.pageTitles[this.page]}_${this.startDate}_${this.endDate}`;
      window.print();
      setTimeout(() => { document.title = origTitle; }, 2000);
    },

    // ── AI Insights ──
    chartMeta(id) {
      const meta = {
        'chart-sentiment-donut': ['Total Comments by Sentiment', 'donut', () => this.sentimentDist],
        'chart-sentiment-category': ['Sentiment by Category', 'bar-table', () => this.sentimentByCategory],
        'chart-sentiment-type': ['Sentiment by Type', 'bar-table', () => this.sentimentByType],
        'chart-followers-growth': ['Followers Growth', 'bar-line', () => this.followersGrowth],
        'chart-total-reach': ['Total Reach', 'line', () => this.totalReach],
        'chart-reach-funnel': ['Organic Reach Funnel', 'funnel', () => this.reachFunnel],
        'chart-ig-followers': ['IG Followers Growth', 'line', () => this.igFollowers],
        'chart-pillar-donut': ['Posts by Pillar', 'donut', () => this.pillarDonut],
        'chart-type-donut': ['Posts by Type', 'donut', () => this.typeDonut],
        'chart-pillar-bar': ['Interactions by Pillar', 'bar', () => this.pillarInteractions],
        'chart-category-bar': ['Interactions by Category', 'bar', () => this.catInteractions],
        'chart-ig-pillar-donut': ['Instagram Posts by Pillar', 'donut', () => this.igPillarDonut],
        'chart-ig-engagement': ['Wall Post Engagement', 'donut', () => this.igEngagementDonut],
        'chart-ig-pillar-bar': ['Interactions by Pillar', 'bar', () => this.igPillarInteractions],
        'chart-ig-story-cat': ['Stories by Category', 'donut', () => this.igStoryCatDonut],
        'chart-ig-story-clicks': ['Story Link Clicks', 'bar', () => this.igStoryClicks],
        'chart-fpk-growth': ['Competitor Fans Growth Trend', 'line', () => this.fpkGrowthTrend],
        'chart-fpk-bubble': ['KPI Comparison', 'bubble', () => this.fpkKpiComparison],
        'chart-fpk-words': ['Top 50 Words: Post Interaction Rate', 'tagcloud', () => this.fpkTopWords],
      };
      const tableMeta = {
        'tbl-fb-key-metrics': ['FB Key Metrics', 'table', () => this.fbKeyMetrics],
        'tbl-ig-key-metrics': ['IG Key Metrics', 'table', () => this.igKeyMetrics],
        'tbl-fb-wall-posts': ['FB Wall Post Performance', 'table', () => this.wallPosts],
        'tbl-ig-wall-posts': ['IG Wall Posts', 'table', () => this.igWallPosts],
        'tbl-fpk-overview': ['Competitors Overview', 'table', () => this.fpkCompetitorsOverview],
      };
      return meta[id] || tableMeta[id];
    },

    buildSummary(id) {
      const m = this.chartMeta(id);
      if (!m) return '';
      const data = m[2]();
      if (!data) return '';
      if (Array.isArray(data)) {
        return JSON.stringify(data.slice(0, 30), null, 2);
      }
      return JSON.stringify(data, null, 2);
    },

    async getInsights(id) {
      if (this.insightsLoading[id]) return;
      if (this.aiInsightsBusy) {
        this.insightsText[id] = '⏳ 請等候其他分析完成後再點擊。';
        this.insightsHtml[id] = '<p>⏳ 請等候其他分析完成後再點擊。</p>';
        return;
      }
      const m = this.chartMeta(id);
      if (!m) return;
      const title = m[0], ctype = m[1];
      const summary = this.buildSummary(id);
      if (!summary) return;

      this.aiInsightsBusy = true;
      this.insightsLoading[id] = true;
      this.insightsText[id] = '';
      this.insightsHtml[id] = '';

      try {
        const res = await fetch('/api/ai/insights', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            chart_title: title,
            chart_type: ctype,
            data_summary: summary,
            start: this.startDate,
            end: this.endDate,
          }),
        });

        if (!res.ok) throw new Error('HTTP ' + res.status);

        const reader = res.body.getReader();
        const decoder = new TextDecoder();
        let fullText = '';

        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          fullText += decoder.decode(value, { stream: true });
          this.insightsText[id] = fullText;
          this.insightsHtml[id] = marked.parse(fullText);
        }
      } catch (e) {
        this.insightsText[id] = '抱歉，生成分析時發生錯誤，請稍後再試。';
        this.insightsHtml[id] = marked.parse('抱歉，生成分析時發生錯誤，請稍後再試。');
      }

      this.insightsLoading[id] = false;
      this.aiInsightsBusy = false;
    },

    // ── Chat ──
    openChat() {
      this.chatOpen = true;
      this.$nextTick(() => {
        if (this.$refs.chatMessages) {
          this.$refs.chatMessages.scrollTop = this.$refs.chatMessages.scrollHeight;
        }
      });
    },

    sendPrompt(prompt) {
      this.chatInput = prompt;
      this.sendMessage();
    },

    async sendMessage() {
      const msg = this.chatInput.trim();
      if (!msg || this.chatStreaming) return;

      this.chatInput = '';
      this.chatHistory.push({ role: 'user', content: msg, html: marked.parse(msg) });
      this.chatStreaming = true;

      this.$nextTick(() => {
        if (this.$refs.chatMessages) {
          this.$refs.chatMessages.scrollTop = this.$refs.chatMessages.scrollHeight;
        }
      });

      const assistantIdx = this.chatHistory.length;
      this.chatHistory.push({ role: 'assistant', content: '', html: '' });

      try {
        const res = await fetch('/api/ai/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            message: msg,
            history: this.chatHistory
              .slice(0, assistantIdx)
              .map(h => ({ role: h.role, content: h.content })),
            start: this.startDate,
            end: this.endDate,
          }),
        });

        const reader = res.body.getReader();
        const decoder = new TextDecoder();
        let fullText = '';

        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          const chunk = decoder.decode(value, { stream: true });
          fullText += chunk;
          this.chatHistory[assistantIdx].content = fullText;
          this.chatHistory[assistantIdx].html = marked.parse(fullText);

          this.$nextTick(() => {
            if (this.$refs.chatMessages) {
              this.$refs.chatMessages.scrollTop = this.$refs.chatMessages.scrollHeight;
            }
          });
        }
      } catch (e) {
        this.chatHistory[assistantIdx].content = '抱歉，發生錯誤。請檢查 OpenRouter API 設定。';
        this.chatHistory[assistantIdx].html = marked.parse('抱歉，發生錯誤。請檢查 OpenRouter API 設定。');
      }

      this.chatStreaming = false;
    },

    // ── Excel Export (SheetJS) ──
    exportToExcel(data, filename, sheetName) {
      if (!data || data.length === 0) return;
      if (typeof XLSX === 'undefined') {
        alert('SheetJS library not loaded. Cannot export.');
        return;
      }
      var ws = XLSX.utils.json_to_sheet(data);
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Sheet1');
      XLSX.writeFile(wb, filename);
    },

    exportGrowthTrend() {
      if (!this.fpkGrowthTrend) return;
      var d = this.fpkGrowthTrend;
      var self = this;
      var rows = d.dates.map(function(date, i) {
        var row = { Date: date };
        d.series.forEach(function(s) { row[s.name] = s.data[i]; });
        return row;
      });
      this.exportToExcel(rows, 'competitor_growth_trend_' + this.startDate + '_' + this.endDate + '.xlsx', 'Growth Trend');
    },

    exportFollowersGrowth() {
      if (!this.followersGrowth) return;
      var d = this.followersGrowth;
      var rows = d.dates.map(function(date, i) {
        return { Date: date, Gain: d.gain[i], Loss: d.loss[i], Net: d.net[i] };
      });
      this.exportToExcel(rows, 'followers_growth_' + this.startDate + '_' + this.endDate + '.xlsx', 'Followers Growth');
    },

    exportTotalReach() {
      if (!this.totalReach) return;
      var d = this.totalReach;
      var rows = d.dates.map(function(date, i) {
        return { Date: date, 'Total Reach': d.reach[i] };
      });
      this.exportToExcel(rows, 'total_reach_' + this.startDate + '_' + this.endDate + '.xlsx', 'Total Reach');
    },

    exportReachFunnel() {
      if (!this.reachFunnel) return;
      var d = this.reachFunnel;
      this.exportToExcel([
        { Type: 'Organic Reach', Reach: d.organic },
        { Type: 'Paid Reach', Reach: d.paid },
        { Type: 'Total', Reach: d.total },
      ], 'reach_funnel_' + this.startDate + '_' + this.endDate + '.xlsx', 'Reach Funnel');
    },

    exportIgFollowers() {
      if (!this.igFollowers) return;
      var d = this.igFollowers;
      var rows = d.dates.map(function(date, i) {
        return { Date: date, 'Followers Net': d.net[i] };
      });
      this.exportToExcel(rows, 'ig_followers_' + this.startDate + '_' + this.endDate + '.xlsx', 'IG Followers');
    },

    exportKpiComparison() {
      if (!this.fpkKpiComparison) return;
      var rows = this.fpkKpiComparison.items.map(function(item) {
        return { Company: item.name, 'Number of Posts': item.posts, 'Reactions, Comments & Shares': item.reactions };
      });
      this.exportToExcel(rows, 'kpi_comparison_' + this.startDate + '_' + this.endDate + '.xlsx', 'KPI Comparison');
    },

    exportTopWords() {
      if (!this.fpkTopWords || this.fpkTopWords.length === 0) return;
      var rows = this.fpkTopWords.map(function(d) {
        return { Word: d.word, Value: d.value, 'Times Above Average': d.times_above_avg };
      });
      this.exportToExcel(rows, 'top50_words_' + this.startDate + '_' + this.endDate + '.xlsx', 'Top 50 Words');
    },
  };
}
