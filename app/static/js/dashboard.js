/* Mannings Dashboard — Alpine.js controller */

function dashboard(initial) {
  const monthNames = {
    1: 'January', 2: 'February', 3: 'March', 4: 'April',
    5: 'May', 6: 'June', 7: 'July', 8: 'August',
    9: 'September', 10: 'October', 11: 'November', 12: 'December',
  };

  return {
    // ── State ──
    page: initial.page,
    year: initial.year,
    month: initial.month,
    selectedYear: initial.year,
    selectedMonth: initial.month,

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

    // ── LinkedIn ──
    liData: {},

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
      'What are the key highlights this month?',
      'Which pillar performed best and why?',
      'How is the sentiment distribution?',
      'What should we improve next month?',
    ],

    // ── AI Insights ──
    aiInsightsBusy: false,
    insightsText: {},
    insightsHtml: {},
    insightsLoading: {},

    // ── Periods ──
    allPeriods: [[initial.year, initial.month]],
    get years() {
      return [...new Set(this.allPeriods.map(p => p[0]))].sort((a, b) => b - a);
    },
    get monthsForYear() {
      return this.allPeriods
        .filter(p => p[0] === this.selectedYear)
        .map(p => [p[1], monthNames[p[1]] || p[1]])
        .sort((a, b) => a[0] - b[0]);
    },

    get periodLabel() {
      return `${monthNames[this.month] || ''} ${this.year}`;
    },

    get pageTitles() {
      return {
        fb_page: 'FB Page',
        fb_posts: 'FB Posts',
        instagram: 'Instagram',
        linkedin: 'LinkedIn',
      };
    },

    // ── Generic Table Sorting ──
    sortTable(tableKey, col) {
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
      const dateCol = keys.find(k => ['Publish time', 'Post Date', 'Comment Date'].includes(k));
      if (dateCol) {
        this.sortState[tableKey] = { col: dateCol, asc: true };
      }
    },

    // ── Sort data by value descending ──
    sortByValue(arr) {
      return [...arr].sort((a, b) => (b.value || 0) - (a.value || 0));
    },

    // ── Sort sentiment combo by total desc ──
    sortSentimentByTotal(arr) {
      return [...arr].sort((a, b) => {
        const ta = (a.Positive || 0) + (a.Neutral || 0) + (a.Negative || 0);
        const tb = (b.Positive || 0) + (b.Neutral || 0) + (b.Negative || 0);
        return tb - ta;
      });
    },

    // ── Init ──
    init() {
      this.loadPeriods().then(() => {
        this.$nextTick(() => {
          this.selectedYear = this.year;
          this.selectedMonth = this.month;
          this.loadAll(this.year, this.month);
        });
      });
    },

    async loadPeriods() {
      try {
        const res = await fetch('/api/periods');
        const data = await res.json();
        const fetched = (data.periods || []).map(p => [p.year, p.month]);
        if (fetched.length > 0) {
          this.allPeriods = fetched;
        }
      } catch (e) {
        console.error('Failed to load periods:', e);
      }
    },

    // ── Period Change ──
    onPeriodChange() {
      this.year = this.selectedYear;
      this.month = this.selectedMonth;
      this.insightsText = {};
      this.insightsHtml = {};
      this.insightsLoading = {};
      this.sortState = {};
      this.loadAll(this.year, this.month);
      const url = new URL(window.location);
      url.searchParams.set('year', this.year);
      url.searchParams.set('month', this.month);
      window.history.replaceState({}, '', url);
    },

    // ── Load All Data ──
    async loadAll(year, month) {
      this.loading = true;
      try {
        await Promise.all([
          this.loadKPIs(year, month),
          this.loadPageData(year, month),
        ]);
        if (this.page === 'fb_page' || this.page === 'instagram') {
          await this.loadFPK(year, month);
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

    async loadKPIs(year, month) {
      try {
        const res = await fetch(`/api/data/kpis?year=${year}&month=${month}`);
        const data = await res.json();
        let kpis = [];
        if (this.page === 'linkedin') {
          return;
        } else if (this.page === 'fb_page' || this.page === 'fb_posts') {
          kpis = [
            { label: 'Total Followers', value: this.fmt(data.fb_followers) },
            { label: 'FB Page Follows Growth', value: this.fmt(data.fb_growth) },
            { label: 'No. of Wall Post', value: this.fmt(data.fb_wall_posts) },
            { label: 'Total Interactions', value: this.fmt(data.fb_interactions) },
          ];
        } else {
          kpis = [
            { label: 'Total Followers', value: this.fmt(data.ig_followers) },
            { label: 'Instagram Growth', value: this.fmt(data.ig_growth) },
            { label: 'Total Reach', value: this.fmt(data.ig_reach) },
            { label: 'Total Interactions', value: this.fmt(data.ig_interactions) },
          ];
        }
        this.kpis = kpis;
      } catch (e) { console.error('KPI load error:', e); }
    },

    async loadPageData(year, month) {
      const pageMap = {
        fb_page: 'fb_page',
        fb_posts: 'fb_posts',
        instagram: 'instagram',
        linkedin: 'linkedin',
      };
      const endpoint = pageMap[this.page];
      try {
        const res = await fetch(`/api/data/${endpoint}?year=${year}&month=${month}`);
        const data = await res.json();

        if (this.page === 'fb_page') {
          this.sentimentDist = data.sentiment_distribution || [];
          this.sentimentByCategory = this.sortSentimentByTotal(data.sentiment_by_category || []);
          this.sentimentByType = this.sortSentimentByTotal(data.sentiment_by_type || []);
          this.fbKeyMetrics = data.fb_key_metrics || [];
          this.initDefaultSort('fbKeyMetrics', this.fbKeyMetrics);
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
        } else if (this.page === 'linkedin') {
          this.liData = data;
          this.initDefaultSort('liPosts', data.posts || []);
          this.kpis = [
            { label: 'Total Followers', value: this.fmt(data.total_followers || 0) },
            { label: 'New Followers', value: this.fmt(data.new_followers || 0) },
            { label: 'Posts', value: this.fmt(data.post_count || 0) },
            { label: 'Total Impressions', value: this.fmt(data.total_impressions || 0) },
          ];
        }
      } catch (e) { console.error('Page data load error:', e); }
    },

    async loadFPK(year, month) {
      const platform = this.page === 'instagram' ? 'ig' : 'fb';
      try {
        const res = await fetch(`/api/data/fpk?year=${year}&month=${month}&platform=${platform}`);
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
      } else if (this.page === 'linkedin') {
        this.renderLinkedInCharts();
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

    renderLinkedInCharts() {
      var d = this.liData;
      if (!d) return;

      if (d.org_paid_donut && d.org_paid_donut.length > 0) {
        renderLiDonut('chart-li-org-paid', d.org_paid_donut, ['Organic', 'Paid'], ['#43A047', '#FE8301']);
      }
      if (d.net_followers) {
        renderLiNetFollowers('chart-li-net-followers', d.net_followers);
      }
      if (d.clicks) {
        renderLiLine('chart-li-clicks', d.clicks, '#FE8301', 'Clicks');
      }
      if (d.impressions) {
        renderLiLine('chart-li-impressions', d.impressions, '#FE8301', 'Impressions');
      }
      if (d.social_actions) {
        renderLiSocialActions('chart-li-social', d.social_actions);
      }
      if (d.top_countries && d.top_countries.length > 0) {
        renderLiHBar('chart-li-country', d.top_countries, '#FE8301');
      }
      if (d.top_company_size && d.top_company_size.length > 0) {
        renderLiHBar('chart-li-company-size', d.top_company_size, '#FE8301');
      }
      if (d.top_seniority && d.top_seniority.length > 0) {
        renderLiHBar('chart-li-seniority', d.top_seniority, '#FE8301');
      }
      if (d.top_job_function && d.top_job_function.length > 0) {
        renderLiVBarMulti('chart-li-job-function', d.top_job_function);
      }
      if (d.top_industry && d.top_industry.length > 0) {
        renderLiVBarMulti('chart-li-industry', d.top_industry);
      }
      if (d.visitor_metrics) {
        renderLiVisitorMetrics('chart-li-visitors', d.visitor_metrics);
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
      document.title = `Mannings_${this.pageTitles[this.page]}_${this.year}_${this.month}`;
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
        'chart-li-org-paid': ['LinkedIn: Organic vs. Paid New Followers', 'donut', () => (this.liData || {}).org_paid_donut],
        'chart-li-net-followers': ['LinkedIn: Net Followers', 'bar', () => (this.liData || {}).net_followers],
        'chart-li-clicks': ['LinkedIn: Clicks', 'line', () => (this.liData || {}).clicks],
        'chart-li-impressions': ['LinkedIn: Impressions', 'line', () => (this.liData || {}).impressions],
        'chart-li-social': ['LinkedIn: Social Actions', 'bar', () => (this.liData || {}).social_actions],
        'chart-li-country': ['LinkedIn: Top Followers by Country', 'bar', () => (this.liData || {}).top_countries],
        'chart-li-company-size': ['LinkedIn: Top Followers by Company Size', 'bar', () => (this.liData || {}).top_company_size],
        'chart-li-seniority': ['LinkedIn: Top Followers by Seniority', 'bar', () => (this.liData || {}).top_seniority],
        'chart-li-job-function': ['LinkedIn: Followers by Job Function', 'bar', () => (this.liData || {}).top_job_function],
        'chart-li-industry': ['LinkedIn: Followers by Industry', 'bar', () => (this.liData || {}).top_industry],
        'chart-li-visitors': ['LinkedIn: Visitor Metrics', 'line', () => (this.liData || {}).visitor_metrics],
        'chart-li-followers-card': ['LinkedIn: Followers Summary', 'card', () => ({
          total_followers: (this.liData || {}).total_followers,
          new_followers: (this.liData || {}).new_followers,
          organic: (this.liData || {}).organic_followers,
          paid: (this.liData || {}).paid_followers,
        })],
      };
      const tableMeta = {
        'tbl-fb-key-metrics': ['FB Key Metrics', 'table', () => this.getSortedData(this.fbKeyMetrics, 'fbKeyMetrics')],
        'tbl-ig-key-metrics': ['IG Key Metrics', 'table', () => this.getSortedData(this.igKeyMetrics, 'igKeyMetrics')],
        'tbl-fb-wall-posts': ['FB Wall Post Performance', 'table', () => this.getSortedData(this.wallPosts, 'fbWall')],
        'tbl-ig-wall-posts': ['IG Wall Posts', 'table', () => this.getSortedData(this.igWallPosts, 'igWall')],
        'tbl-fpk-overview': ['Competitors Overview', 'table', () => this.getSortedData(this.fpkCompetitorsOverview, 'fpkOverview')],
        'tbl-li-posts': ['LinkedIn Post Content', 'table', () => this.getSortedData((this.liData || {}).posts || [], 'liPosts')],
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
            year: this.year,
            month: this.month,
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
            year: this.year,
            month: this.month,
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
      var rows = d.dates.map(function(date, i) {
        var row = { Date: date };
        d.series.forEach(function(s) { row[s.name] = s.data[i]; });
        return row;
      });
      this.exportToExcel(rows, 'competitor_growth_trend_' + this.year + '_' + this.month + '.xlsx', 'Growth Trend');
    },

    exportFollowersGrowth() {
      if (!this.followersGrowth) return;
      var d = this.followersGrowth;
      var rows = d.dates.map(function(date, i) {
        return { Date: date, Gain: d.gain[i], Loss: d.loss[i], Net: d.net[i] };
      });
      this.exportToExcel(rows, 'followers_growth_' + this.year + '_' + this.month + '.xlsx', 'Followers Growth');
    },

    exportTotalReach() {
      if (!this.totalReach) return;
      var d = this.totalReach;
      var rows = d.dates.map(function(date, i) {
        return { Date: date, 'Total Reach': d.reach[i] };
      });
      this.exportToExcel(rows, 'total_reach_' + this.year + '_' + this.month + '.xlsx', 'Total Reach');
    },

    exportReachFunnel() {
      if (!this.reachFunnel) return;
      var d = this.reachFunnel;
      this.exportToExcel([
        { Type: 'Organic Reach', Reach: d.organic },
        { Type: 'Paid Reach', Reach: d.paid },
        { Type: 'Total', Reach: d.total },
      ], 'reach_funnel_' + this.year + '_' + this.month + '.xlsx', 'Reach Funnel');
    },

    exportIgFollowers() {
      if (!this.igFollowers) return;
      var d = this.igFollowers;
      var rows = d.dates.map(function(date, i) {
        return { Date: date, 'Followers Net': d.net[i] };
      });
      this.exportToExcel(rows, 'ig_followers_' + this.year + '_' + this.month + '.xlsx', 'IG Followers');
    },

    exportKpiComparison() {
      if (!this.fpkKpiComparison) return;
      var rows = this.fpkKpiComparison.items.map(function(item) {
        return { Company: item.name, 'Number of Posts': item.posts, 'Reactions, Comments & Shares': item.reactions };
      });
      this.exportToExcel(rows, 'kpi_comparison_' + this.year + '_' + this.month + '.xlsx', 'KPI Comparison');
    },

    exportTopWords() {
      if (!this.fpkTopWords || this.fpkTopWords.length === 0) return;
      var rows = this.fpkTopWords.map(function(d) {
        return { Word: d.word, Value: d.value, 'Times Above Average': d.times_above_avg };
      });
      this.exportToExcel(rows, 'top50_words_' + this.year + '_' + this.month + '.xlsx', 'Top 50 Words');
    },

    exportLiNetFollowers() {
      if (!this.liData || !this.liData.net_followers) return;
      var nf = this.liData.net_followers;
      var rows = nf.dates.map(function(d, i) {
        return { Date: d, Organic: nf.organic[i] || 0, Paid: nf.paid[i] || 0 };
      });
      this.exportToExcel(rows, 'li_net_followers_' + this.year + '_' + this.month + '.xlsx', 'Net Followers');
    },

    exportLiDailyData(data, label) {
      if (!data || !data.dates) return;
      var rows = data.dates.map(function(d, i) {
        var r = { Date: d };
        r[label.charAt(0).toUpperCase() + label.slice(1)] = data.values[i] || 0;
        return r;
      });
      this.exportToExcel(rows, 'li_' + label + '_' + this.year + '_' + this.month + '.xlsx', label);
    },

    exportLiSocialActions() {
      if (!this.liData || !this.liData.social_actions) return;
      var sa = this.liData.social_actions;
      var rows = sa.dates.map(function(d, i) {
        return { Date: d, Comments: sa.comments[i] || 0, Likes: sa.likes[i] || 0, Shares: sa.shares[i] || 0 };
      });
      this.exportToExcel(rows, 'li_social_actions_' + this.year + '_' + this.month + '.xlsx', 'Social Actions');
    },

    exportLiVisitorMetrics() {
      if (!this.liData || !this.liData.visitor_metrics) return;
      var vm = this.liData.visitor_metrics;
      var rows = vm.dates.map(function(d, i) {
        return { Date: d, Desktop: vm.desktop[i] || 0, Mobile: vm.mobile[i] || 0 };
      });
      this.exportToExcel(rows, 'li_visitor_metrics_' + this.year + '_' + this.month + '.xlsx', 'Visitor Metrics');
    },
  };
}
