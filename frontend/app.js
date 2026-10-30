/* global Vue, VueRouter, axios */

const api = axios.create({ timeout: 240000 });

const toastState = Vue.observable({
  visible: false,
  message: "",
  type: "info",
  timer: null
});

function notify(message, type) {
  toastState.message = message || "操作失败";
  toastState.type = type || "info";
  toastState.visible = true;
  clearTimeout(toastState.timer);
  toastState.timer = setTimeout(function () {
    toastState.visible = false;
  }, 4200);
}

Vue.prototype.$notify = notify;

api.interceptors.response.use(
  function (response) {
    return response;
  },
  function (error) {
    let message = error.message;
    if (error.response && error.response.data && error.response.data.error) {
      message = error.response.data.error;
    }
    notify(message, "error");
    return Promise.reject(error);
  }
);

const STATUS_LABELS = {
  pending: "排队中",
  running: "执行中",
  success: "成功",
  failed: "失败"
};

const FIELD_LABELS = {
  ticket: "需求/Ticket编号",
  description: "任务描述/完成事项",
  author: "执行人/姓名",
  repo: "代码路径",
  branch: "分支",
  date: "日期",
  commits: "提交数",
  hours: "投入工时(h)",
  progress: "进度(%)",
  remark: "备注/状态"
};

const WEEKDAY_LABELS = ["一", "二", "三", "四", "五", "六", "日"];

function pad(value) {
  return value < 10 ? "0" + value : String(value);
}

function parseDate(value) {
  const parts = String(value).slice(0, 10).split("-");
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function formatDate(date) {
  return (
    date.getFullYear() +
    "-" +
    pad(date.getMonth() + 1) +
    "-" +
    pad(date.getDate())
  );
}

function todayStr() {
  return formatDate(new Date());
}

function addDays(value, days) {
  const date = parseDate(value);
  date.setDate(date.getDate() + days);
  return formatDate(date);
}

function weekStartOf(value, startDay) {
  const date = parseDate(value);
  const pythonWeekday = (date.getDay() + 6) % 7;
  const offset = (pythonWeekday - (Number(startDay) || 0) + 7) % 7;
  date.setDate(date.getDate() - offset);
  return formatDate(date);
}

function isoWeek(value) {
  const date = parseDate(value);
  const target = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const dayNr = (date.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  const firstThursday = new Date(target.getFullYear(), 0, 4);
  const firstDayNr = (firstThursday.getDay() + 6) % 7;
  firstThursday.setDate(firstThursday.getDate() - firstDayNr + 3);
  return 1 + Math.round((target - firstThursday) / 604800000);
}

function splitList(text) {
  return String(text || "")
    .split(/[,，;；\n]/)
    .map(function (item) {
      return item.trim();
    })
    .filter(Boolean);
}

Vue.mixin({
  methods: {
    statusText(status) {
      return STATUS_LABELS[status] || status || "—";
    },
    fieldLabel(field) {
      return FIELD_LABELS[field] || field;
    },
    shortTime(value) {
      return value ? String(value).slice(0, 16) : "—";
    }
  }
});

Vue.component("inspect-view", {
  props: { result: Object },
  data() {
    return { active: "" };
  },
  computed: {
    current() {
      const sheets = (this.result && this.result.sheets) || [];
      for (let index = 0; index < sheets.length; index += 1) {
        if (sheets[index].name === this.active) {
          return sheets[index];
        }
      }
      return null;
    },
    detected() {
      return (this.result && this.result.detected) || null;
    }
  },
  watch: {
    result: {
      immediate: true,
      handler(value) {
        if (!value) {
          this.active = "";
          return;
        }
        const sheets = value.sheets || [];
        this.active = value.detected
          ? value.detected.name
          : sheets.length
            ? sheets[0].name
            : "";
      }
    }
  },
  template: `
    <div class="inspect-result">
      <div class="inspect-summary">
        <span v-if="detected">
          识别到工作表「{{ detected.name }}」第 {{ detected.header_row }} 行表头，匹配 {{ Object.keys(detected.columns).length }} 个字段。
        </span>
        <span v-else>
          未识别到表头，请手动指定表头行，或在字段映射里补充别名。
        </span>
      </div>
      <div class="sheet-tabs">
        <button
          v-for="sheet in result.sheets"
          :key="sheet.name"
          :class="{ active: sheet.name === active }"
          @click="active = sheet.name"
        >{{ sheet.name }}</button>
      </div>
      <div v-if="current">
        <div class="column-map">
          <span v-for="(column, field) in current.columns" :key="field" class="map-chip">
            {{ fieldLabel(field) }} → 第 {{ column }} 列
          </span>
          <span v-if="!Object.keys(current.columns).length" class="map-chip">未匹配到字段</span>
        </div>
        <div class="table-wrap">
          <table class="data-table compact">
            <tbody>
              <tr
                v-for="(row, index) in current.preview"
                :key="index"
                :class="{ header: current.header_row === index + 1 }"
              >
                <td class="rownum">{{ index + 1 }}</td>
                <td v-for="(cell, cellIndex) in row" :key="cellIndex">{{ cell }}</td>
              </tr>
            </tbody>
          </table>
        </div>
        <button
          v-if="current.header_row"
          class="mini-button primary"
          style="margin-top: 12px"
          @click="$emit('apply', current)"
        >应用此工作表与表头</button>
      </div>
    </div>
  `
});

const HomePage = {
  data() {
    return {
      config: null,
      weekAnchor: todayStr(),
      rangeMode: "week",
      customStart: weekStartOf(todayStr(), 0),
      customEnd: addDays(weekStartOf(todayStr(), 0), 6),
      author: "",
      repos: [],
      collecting: false,
      rewritingAI: false,
      preview: null,
      rows: [],
      rawOpen: false,
      fillOpen: false,
      filling: false,
      dryRun: true,
      activeRun: null,
      existingRuns: [],
      workbookStatus: {
        checking: false,
        excel_running: false,
        found: false,
        workbooks: [],
        workbook: null,
        target: "",
        message: ""
      },
      pollTimer: null
    };
  },
  computed: {
    rangeStart() {
      if (this.rangeMode === "custom") {
        return this.customStart;
      }
      return weekStartOf(
        this.weekAnchor,
        this.config ? this.config.week_start_day : 0
      );
    },
    rangeEnd() {
      if (this.rangeMode === "custom") {
        return this.customEnd;
      }
      return addDays(this.rangeStart, 6);
    },
    weekStart() {
      return this.rangeStart;
    },
    weekEnd() {
      return this.rangeEnd;
    },
    rangeDays() {
      const start = parseDate(this.rangeStart);
      const end = parseDate(this.rangeEnd);
      return Math.max(1, Math.round((end - start) / 86400000) + 1);
    },
    rangeLabel() {
      return this.rangeStart + " ~ " + this.rangeEnd;
    },
    weekRangeLabel() {
      return (
        this.weekStart.slice(5).replace("-", "/") +
        " — " +
        this.weekEnd.slice(5).replace("-", "/")
      );
    },
    weekAnchorLabel() {
      return this.weekStart.slice(0, 4) + " · 第 " + isoWeek(this.weekStart) + " 周";
    },
    dailyMap() {
      const map = {};
      if (!this.preview) {
        return map;
      }
      (this.preview.daily || []).forEach(function (item) {
        map[item.date] = item.count;
      });
      return map;
    },
    weekDays() {
      const counts = this.dailyMap;
      const days = [];
      const today = todayStr();
      for (let index = 0; index < 7; index += 1) {
        const dateStr = addDays(this.weekStart, index);
        const dateObj = parseDate(dateStr);
        const pythonWeekday = (dateObj.getDay() + 6) % 7;
        days.push({
          date: dateStr,
          label: WEEKDAY_LABELS[pythonWeekday],
          num: dateObj.getDate(),
          count: counts[dateStr] || 0,
          isToday: dateStr === today,
          isWeekend: pythonWeekday >= 5
        });
      }
      return days;
    },
    repoNames() {
      return this.repos.map(function (path) {
        const parts = path.split(/[\\/]/).filter(Boolean);
        return parts.length ? parts[parts.length - 1] : path;
      });
    },
    ticketCount() {
      return this.rows.length;
    },
    uniqueTicketCount() {
      const seen = {};
      this.rows.forEach(function (row) {
        seen[row.ticket] = true;
      });
      return Object.keys(seen).length;
    },
    aggregateLabel() {
      if (this.config && this.config.target_mode === "weekly") {
        return "按 Ticket（周报汇总）";
      }
      if (this.config && this.config.target_mode === "both") {
        return "日报：按 Ticket + 日期 · 周报汇总：按 Ticket";
      }
      const mode = (this.config && this.config.aggregate_mode) || "daily";
      return mode === "daily" ? "按 Ticket + 日期" : "按 Ticket";
    },
    weeklySheetName() {
      const prefix = (
        (this.config && this.config.weekly_sheet_prefix) || "周报汇总"
      ).trim();
      const separator = /[ _-]$/.test(prefix) ? "" : " ";
      const start = parseDate(this.rangeStart);
      return prefix + separator + (start.getMonth() + 1) + "." + start.getDate();
    },
    totalCommits() {
      return this.preview ? this.preview.commits_count : 0;
    },
    dateModeLabel() {
      if (this.config && this.config.target_mode === "weekly") {
        return "无日期列";
      }
      if (this.config && this.config.target_mode === "both") {
        return "日报：提交当天";
      }
      if (((this.config && this.config.aggregate_mode) || "daily") === "daily") {
        return "提交当天";
      }
      const mode = (this.config && this.config.date_mode) || "last";
      const labels = {
        last: "最后提交日",
        first: "首次提交日",
        week_start: "范围起始日",
        range: "日期范围（文本）"
      };
      return labels[mode] || mode;
    },
    resultTitle() {
      if (!this.activeRun) {
        return "";
      }
      if (this.activeRun.status === "success") {
        return this.activeRun.mode === "dry_run"
          ? "演练完成"
          : "已写入 SharePoint";
      }
      if (this.activeRun.status === "failed") {
        return "填充失败";
      }
      return "正在处理…";
    },
    workbookNames() {
      return (this.workbookStatus.workbooks || [])
        .map(function (item) {
          return item.name;
        })
        .join("、");
    },
    fillReady() {
      if (this.filling || !this.config) {
        return false;
      }
      if ((this.config.write_mode || "attach") !== "attach") {
        return true;
      }
      return !!this.workbookStatus.found;
    }
  },
  watch: {
    weekAnchor() {
      if (this.rangeMode !== "week") {
        return;
      }
      this.resetPreview();
      this.checkExistingRuns();
    }
  },
  created() {
    this.loadConfig();
  },
  beforeDestroy() {
    clearTimeout(this.pollTimer);
  },
  methods: {
    loadConfig() {
      api
        .get("/api/config")
        .then((res) => {
          this.config = res.data.config;
          this.author = this.config.author;
          this.repos = this.config.repos.slice();
          this.dryRun = this.config.dry_run;
          if (!this.restorePreview() && this.repos.length) {
            this.collect(true);
          }
          this.checkExistingRuns();
          this.checkWorkbook();
        })
        .catch(function () {});
    },
    restorePreview() {
      try {
        const raw = sessionStorage.getItem("weeklog.preview");
        if (!raw) {
          return false;
        }
        const cache = JSON.parse(raw);
        if (
          cache.week_start !== this.weekStart ||
          cache.week_end !== this.weekEnd ||
          cache.range_mode !== this.rangeMode
        ) {
          return false;
        }
        if (Date.now() - cache.saved_at > 10 * 60 * 1000) {
          return false;
        }
        this.preview = cache.preview;
        this.rows = cache.rows;
        return true;
      } catch (error) {
        return false;
      }
    },
    cachePreview() {
      try {
        sessionStorage.setItem(
          "weeklog.preview",
          JSON.stringify({
            saved_at: Date.now(),
            week_start: this.weekStart,
            week_end: this.weekEnd,
            range_mode: this.rangeMode,
            preview: this.preview,
            rows: this.rows
          })
        );
      } catch (error) {}
    },
    resetPreview() {
      this.preview = null;
      this.rows = [];
      this.rawOpen = false;
      this.activeRun = null;
      try {
        sessionStorage.removeItem("weeklog.preview");
      } catch (error) {}
    },
    prevWeek() {
      this.weekAnchor = addDays(this.weekStart, -7);
    },
    nextWeek() {
      this.weekAnchor = addDays(this.weekStart, 7);
    },
    goToday() {
      this.weekAnchor = todayStr();
      this.rangeMode = "week";
      this.resetPreview();
      this.checkExistingRuns();
    },
    setRangeMode(mode) {
      if (mode === this.rangeMode) {
        return;
      }
      if (mode === "custom") {
        this.customStart = this.weekStart;
        this.customEnd = this.weekEnd;
      }
      this.rangeMode = mode;
      this.resetPreview();
      this.checkExistingRuns();
    },
    onCustomRangeChange() {
      if (this.customStart > this.customEnd) {
        this.$notify("开始日期不能晚于结束日期", "error");
        return;
      }
      this.resetPreview();
      this.checkExistingRuns();
    },
    onAuthorChange() {
      this.resetPreview();
      this.collect(true);
    },
    collect(silent) {
      if (this.collecting) {
        return;
      }
      if (this.rangeMode === "custom" && this.customStart > this.customEnd) {
        this.$notify("开始日期不能晚于结束日期", "error");
        return;
      }
      this.collecting = true;
      this.preview = null;
      this.rows = [];
      this.activeRun = null;
      api
        .post("/api/collect", {
          range_start: this.rangeStart,
          range_end: this.rangeEnd,
          author: this.author,
          repos: this.repos
        })
        .then((res) => {
          this.preview = res.data;
          this.rows = res.data.rows.map(function (row) {
            return Object.assign({}, row);
          });
          this.cachePreview();
          if (!silent) {
            this.$notify(
              res.data.commits_count
                ? "已采集 " + res.data.commits_count + " 条提交"
                : "所选范围没有匹配的提交",
              res.data.commits_count ? "success" : "info"
            );
          }
        })
        .catch(function () {})
        .finally(() => {
          this.collecting = false;
        });
    },
    rewriteDescriptions() {
      if (!this.rows.length || this.rewritingAI) {
        return;
      }
      this.rewritingAI = true;
      api
        .post("/api/ai/rewrite", { rows: this.rows })
        .then((res) => {
          this.rows = res.data.rows.map(function (row) {
            return Object.assign({}, row);
          });
          this.cachePreview();
          this.$notify("AI 改写完成", "success");
          if (res.data.warnings && res.data.warnings.length) {
            this.$notify(res.data.warnings[0], "error");
          }
        })
        .finally(() => {
          this.rewritingAI = false;
        });
    },
    checkExistingRuns() {
      api
        .get("/api/runs", {
          params: {
            week_start: this.weekStart,
            week_end: this.weekEnd,
            mode: "live",
            status: "success",
            limit: 5
          }
        })
        .then((res) => {
          this.existingRuns = res.data.runs || [];
        })
        .catch(function () {});
    },
    openFill() {
      if (!this.rows.length) {
        this.$notify("请先采集提交", "error");
        return;
      }
      this.fillOpen = true;
      this.checkWorkbook();
    },
    checkWorkbook() {
      if (!this.config || (this.config.write_mode || "attach") !== "attach") {
        this.workbookStatus = {
          checking: false,
          excel_running: true,
          found: true,
          workbooks: [],
          workbook: null,
          target: "",
          message: ""
        };
        return;
      }
      this.workbookStatus = Object.assign({}, this.workbookStatus, {
        checking: true
      });
      api
        .get("/api/workbook/status")
        .then((res) => {
          this.workbookStatus = Object.assign({ checking: false }, res.data);
        })
        .catch(() => {
          this.workbookStatus = {
            checking: false,
            excel_running: false,
            found: false,
            workbooks: [],
            workbook: null,
            target: "",
            message: "检测失败"
          };
        });
    },
    confirmFill() {
      if (this.filling) {
        return;
      }
      if (
        this.config &&
        (this.config.write_mode || "attach") === "attach" &&
        !this.workbookStatus.found
      ) {
        this.$notify("请先在 Excel 中打开周报，再点击开始", "error");
        return;
      }
      this.filling = true;
      const rows = this.rows.map(function (row) {
        return {
          row_key: row.row_key,
          day: row.day,
          ticket: row.ticket,
          description: row.description,
          repos: row.repos,
          branches: row.branches,
          date_range: row.date_range,
          commit_count: row.commit_count,
          hours: row.hours,
          progress: row.progress,
          remark: row.remark
        };
      });
      api
        .post("/api/fill", {
          range_start: this.rangeStart,
          range_end: this.rangeEnd,
          rows: rows,
          dry_run: this.dryRun,
          author: this.author
        })
        .then((res) => {
          this.fillOpen = false;
          this.activeRun = {
            id: res.data.run_id,
            status: "pending",
            mode: res.data.mode,
            message: "任务已启动"
          };
          this.$notify("任务已启动，正在通过 Excel 处理…", "info");
          this.pollRun(res.data.run_id);
        })
        .catch(() => {
          this.filling = false;
        });
    },
    pollRun(id) {
      clearTimeout(this.pollTimer);
      this.pollTimer = setTimeout(() => {
        api
          .get("/api/runs/" + id)
          .then((res) => {
            this.activeRun = res.data;
            if (res.data.status === "success" || res.data.status === "failed") {
              this.filling = false;
              this.$notify(
                this.resultTitle,
                res.data.status === "success" ? "success" : "error"
              );
              if (res.data.status === "success") {
                this.checkExistingRuns();
                try {
                  sessionStorage.removeItem("weeklog.preview");
                } catch (error) {}
              }
            } else {
              this.pollRun(id);
            }
          })
          .catch(() => {
            this.filling = false;
          });
      }, 1500);
    }
  },
  template: `
    <div class="stack">
      <section class="hero">
        <div class="hero-copy">
          <p class="eyebrow">GIT → 周报 · {{ rangeStart }} / {{ rangeEnd }}</p>
          <h1>把 Git 提交<br><span>自动写进 SharePoint 周报</span></h1>
          <p class="lede">
            读取本地仓库，按作者和日期筛选，再按 Ticket 汇总成一行。可选按周或自定义时间范围，确认后由桌面 Excel 追加到在线文档，并保留快照、审计表与完整日志。
          </p>
          <div class="hero-actions">
            <button class="pill-button primary" @click="collect(false)" :disabled="collecting">
              <span v-if="collecting" class="spinner"></span>
              {{ collecting ? (config && config.ai_enabled ? "正在读取 Git 并调用 AI…" : "正在读取 Git…") : "采集提交" }}
            </button>
            <button class="pill-button ghost" @click="openFill" :disabled="!rows.length || filling">
              写入周报
            </button>
          </div>
        </div>

        <aside class="week-card">
          <div class="week-head">
            <div class="range-switch">
              <button :class="{ active: rangeMode === 'week' }" @click="setRangeMode('week')">按周</button>
              <button :class="{ active: rangeMode === 'custom' }" @click="setRangeMode('custom')">自定义</button>
            </div>
            <button class="week-today" @click="goToday">回到本周</button>
          </div>

          <div v-if="rangeMode === 'week'" class="week-nav">
            <button class="icon-button" @click="prevWeek" title="上一周">‹</button>
            <div class="week-nav-label">
              <strong>{{ weekRangeLabel }}</strong>
              <small>{{ weekAnchorLabel }}</small>
            </div>
            <button class="icon-button" @click="nextWeek" title="下一周">›</button>
          </div>
          <div v-if="rangeMode === 'week'" class="week-strip">
            <div
              v-for="day in weekDays"
              :key="day.date"
              class="day"
              :class="{ today: day.isToday, weekend: day.isWeekend }"
            >
              <span class="day-label">{{ day.label }}</span>
              <span class="day-num">{{ day.num }}</span>
              <span class="day-dot" :class="{ filled: day.count > 0 }">{{ day.count || "" }}</span>
            </div>
          </div>
          <div v-else class="range-picker">
            <label class="range-field">
              <span>开始</span>
              <input type="date" v-model="customStart" @change="onCustomRangeChange">
            </label>
            <label class="range-field">
              <span>结束</span>
              <input type="date" v-model="customEnd" @change="onCustomRangeChange">
            </label>
            <div class="range-summary">{{ rangeDays }} 天 · {{ rangeStart }} ~ {{ rangeEnd }}</div>
          </div>

          <div class="week-meta">
            <span>{{ author || "全部作者" }}</span>
            <span>{{ repos.length }} 个仓库</span>
            <span v-if="config && config.ai_enabled" class="mode-chip ai">AI 改写</span>
            <span class="mode-chip" :class="dryRun ? 'dry' : 'live'">
              {{ dryRun ? "演练模式" : "正式写入" }}
            </span>
          </div>
        </aside>
      </section>

      <section class="context-bar card">
        <label class="inline-field">
          <span>作者</span>
          <input v-model="author" @change="onAuthorChange" placeholder="全部作者">
        </label>
        <div class="repo-chips">
          <span v-for="name in repoNames" :key="name" class="repo-chip">{{ name }}</span>
        </div>
        <router-link to="/settings" class="mini-button">调整配置</router-link>
      </section>

      <section v-if="collecting && !preview" class="card loading-panel">
        <span class="spinner dark"></span>
        {{ config && config.ai_enabled ? "正在读取 Git 并调用 AI 改写描述，请稍候…" : "正在扫描仓库并汇总 Ticket，请稍候…" }}
      </section>

      <section v-if="preview" class="preview card">
        <header class="section-head">
          <div>
            <h2>汇总预览</h2>
            <p>{{ aggregateLabel }} · {{ uniqueTicketCount }} 个 Ticket · {{ rows.length }} 行 · {{ totalCommits }} 条提交 · {{ weekStart }} ~ {{ weekEnd }} · 日期写入{{ dateModeLabel }}</p>
          </div>
          <div class="section-actions">
            <button v-if="config && config.ai_enabled" class="mini-button" @click="rewriteDescriptions" :disabled="rewritingAI">
              {{ rewritingAI ? "AI 改写中…" : "AI 重写描述" }}
            </button>
            <button class="mini-button" @click="collect(false)" :disabled="collecting">重新采集</button>
            <button class="pill-button primary" @click="openFill" :disabled="filling">写入周报</button>
          </div>
        </header>

        <div v-if="preview.warnings && preview.warnings.length" class="existing-banner error">
          <span>{{ preview.warnings[0] }}</span>
        </div>

        <div v-if="existingRuns.length" class="existing-banner">
          <span>该时间范围已有 {{ existingRuns.length }} 条成功写入记录</span>
          <router-link :to="'/runs/' + existingRuns[0].id">查看 #{{ existingRuns[0].id }}</router-link>
          <span class="muted">再次写入会在表格中追加新行</span>
        </div>

        <div class="table-wrap">
          <table class="data-table editable">
            <thead>
              <tr>
                <th>#</th>
                <th>Ticket</th>
                <th>任务描述</th>
                <th>仓库</th>
                <th>日期</th>
                <th class="center">提交</th>
                <th>工时(h)</th>
                <th>进度(%)</th>
                <th>备注</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(row, index) in rows" :key="row.row_key || row.ticket">
                <td class="index">{{ index + 1 }}</td>
                <td><span class="ticket-chip">{{ row.ticket }}</span></td>
                <td class="desc">
                  <div v-if="row.ai_generated" class="ai-badge">AI</div>
                  <textarea v-model="row.description" class="cell-textarea" rows="3"></textarea>
                </td>
                <td class="muted">{{ row.repos }}</td>
                <td class="muted">{{ row.date_range }}</td>
                <td class="center">{{ row.commit_count }}</td>
                <td><input v-model="row.hours" class="cell-input" placeholder="—"></td>
                <td><input v-model="row.progress" class="cell-input" placeholder="—"></td>
                <td><input v-model="row.remark" class="cell-input wide" placeholder="—"></td>
              </tr>
            </tbody>
          </table>
        </div>

        <div class="raw-toggle" @click="rawOpen = !rawOpen">
          <span>{{ rawOpen ? "收起" : "查看" }} Git 提交明细（{{ totalCommits }}）</span>
          <span class="chev" :class="{ open: rawOpen }">›</span>
        </div>
        <div v-if="rawOpen" class="raw-list">
          <div v-for="commit in preview.commits" :key="commit.commit_hash" class="raw-item">
            <span class="raw-time">{{ commit.committed_at.slice(5, 16) }}</span>
            <span class="raw-ticket">{{ commit.ticket || "—" }}</span>
            <span class="raw-subject">{{ commit.subject }}</span>
            <span class="raw-repo">{{ commit.repo }} · {{ commit.branch }}</span>
          </div>
        </div>
      </section>

      <section v-if="activeRun" class="result-card card" :class="activeRun.status">
        <div class="result-icon">
          {{ activeRun.status === "success" ? "✓" : activeRun.status === "failed" ? "!" : "…" }}
        </div>
        <div class="result-copy">
          <strong>{{ resultTitle }}</strong>
          <p>{{ activeRun.message || "正在通过 Excel 处理，请勿关闭程序…" }}</p>
          <p v-if="activeRun.sheet_name" class="muted">
            工作表「{{ activeRun.sheet_name }}」 · 第 {{ activeRun.start_row }}–{{ activeRun.end_row }} 行
          </p>
        </div>
        <div class="result-actions">
          <router-link v-if="activeRun.id" class="mini-button primary" :to="'/runs/' + activeRun.id">
            查看记录
          </router-link>
        </div>
      </section>

      <section v-if="!preview && !collecting" class="empty-state card">
        <div class="empty-art"><span></span><span></span><span></span></div>
        <h3>先采集一次提交</h3>
        <p>
          程序会读取设置里的仓库，按作者和日期过滤，再按 Ticket 汇总。当前周：{{ weekStart }} ~ {{ weekEnd }}。
        </p>
        <button class="pill-button primary" @click="collect(false)">采集提交</button>
      </section>

      <div v-if="fillOpen" class="modal-backdrop" @click.self="fillOpen = false">
        <div class="modal">
          <h3>确认写入周报</h3>
          <p v-if="config && (config.write_mode || 'attach') === 'attach'">
            程序会附着到你已经打开的 Excel 周报，识别表头后把汇总行追加到表格末尾。
          </p>
          <p v-else>程序会启动 Excel 打开文档链接，识别表头后把汇总行追加到表格末尾。</p>

          <div
            v-if="config && (config.write_mode || 'attach') === 'attach'"
            class="workbook-check"
            :class="{ ok: workbookStatus.found, bad: !workbookStatus.checking && !workbookStatus.found }"
          >
            <div class="workbook-check-text">
              <strong v-if="workbookStatus.checking">正在检测已打开的 Excel…</strong>
              <strong v-else-if="workbookStatus.found">已找到工作簿：{{ workbookStatus.workbook.name }}</strong>
              <strong v-else-if="!workbookStatus.excel_running">未检测到正在运行的 Excel</strong>
              <strong v-else>Excel 已运行，但未找到目标工作簿</strong>
              <small v-if="!workbookStatus.found && !workbookStatus.checking">
                请先在 SharePoint 中点「在桌面应用中打开」，保持周报在 Excel 中打开。
                <template v-if="workbookNames">当前打开：{{ workbookNames }}</template>
              </small>
            </div>
            <button class="mini-button" @click="checkWorkbook" :disabled="workbookStatus.checking">重新检测</button>
          </div>

          <div class="modal-facts">
            <div><span>时间范围</span><strong>{{ weekStart }} ~ {{ weekEnd }}</strong></div>
            <div><span>写入行数</span><strong>{{ rows.length }}</strong></div>
            <div><span>作者过滤</span><strong>{{ author || "全部作者" }}</strong></div>
            <div><span>执行人</span><strong>{{ config && config.author_display ? config.author_display : (author || "—") }}</strong></div>
            <div v-if="config && config.target_mode !== 'weekly'">
              <span>{{ config.target_mode === 'both' ? '日报工作表' : '工作表' }}</span>
              <strong>{{ config.target_sheet || '自动识别' }}</strong>
            </div>
            <div v-if="config && config.target_mode !== 'existing'">
              <span>{{ config.target_mode === 'both' ? '周报汇总' : '新工作表' }}</span>
              <strong>{{ weeklySheetName }}</strong>
            </div>
            <div v-if="config && config.target_mode !== 'existing'">
              <span>模板</span>
              <strong>{{ config.weekly_template_sheet || '自动识别' }}</strong>
            </div>
            <div v-if="config && config.target_mode === 'both'">
              <span>周报汇总行数</span>
              <strong>{{ uniqueTicketCount }}</strong>
            </div>
            <div><span>日期写入</span><strong>{{ dateModeLabel }}</strong></div>
            <div><span>AI 改写</span><strong>{{ config && config.ai_enabled ? "开" : "关" }}</strong></div>
            <div><span>写入方式</span><strong>{{ config && (config.write_mode || 'attach') === 'attach' ? "附着已打开的 Excel" : "程序打开链接" }}</strong></div>
            <div><span>写入模式</span><strong>{{ dryRun ? "演练（不保存）" : "正式写入" }}</strong></div>
          </div>
          <label class="switch-row">
            <input type="checkbox" v-model="dryRun">
            <span v-if="config && (config.write_mode || 'attach') === 'attach'">演练模式：只分析表头与插入位置，不改动已打开的工作簿。</span>
            <span v-else>演练模式：打开文档并回读校验，但不保存修改。</span>
          </label>
          <p v-if="!dryRun" class="modal-warn">
            正式写入会直接修改 SharePoint 在线文档。程序会先保存一份本地快照，并把每一行写入结果记入填充记录。
          </p>
          <p v-if="existingRuns.length" class="modal-warn">
            注意：该时间范围已有成功写入记录 #{{ existingRuns[0].id }}，再次写入会追加新行。
          </p>
          <p class="field-hint" style="margin-top: 12px">
            <template v-if="config && (config.write_mode || 'attach') === 'attach'">
              请保持周报在 Excel 中打开；程序写入后会自动保存，不会关闭你的工作簿。
            </template>
            <template v-else>执行期间请留意 Excel 是否弹出公司账号登录窗口。</template>
          </p>
          <div class="modal-actions">
            <button class="pill-button ghost" @click="fillOpen = false">取消</button>
            <button class="pill-button primary" @click="confirmFill" :disabled="!fillReady">
              {{ filling ? "正在执行…" : "开始" }}
            </button>
          </div>
        </div>
      </div>
    </div>
  `
};

const HistoryPage = {
  data() {
    return { runs: [], loading: true, timer: null };
  },
  created() {
    this.load();
  },
  beforeDestroy() {
    clearTimeout(this.timer);
  },
  methods: {
    load() {
      this.loading = this.runs.length === 0;
      api
        .get("/api/runs", { params: { limit: 100 } })
        .then((res) => {
          this.runs = res.data.runs;
          this.loading = false;
          clearTimeout(this.timer);
          const active = this.runs.some(function (run) {
            return run.status === "running" || run.status === "pending";
          });
          if (active) {
            this.timer = setTimeout(this.load, 2500);
          }
        })
        .catch(() => {
          this.loading = false;
        });
    }
  },
  template: `
    <div>
      <div class="page-head">
        <div>
          <p class="eyebrow">AUDIT TRAIL</p>
          <h1>填充记录</h1>
          <p>每次写入的周区间、Ticket 行、快照、审计表与运行日志都留档在这里。</p>
        </div>
        <button class="pill-button ghost" @click="load">刷新</button>
      </div>

      <div v-if="loading" class="run-list">
        <div class="card skeleton-card" v-for="index in 3" :key="index">
          <div class="skeleton" style="width: 32%"></div>
          <div class="skeleton" style="width: 62%; margin-top: 10px"></div>
        </div>
      </div>

      <div v-else-if="!runs.length" class="empty-state card">
        <div class="empty-art"><span></span><span></span><span></span></div>
        <h3>还没有填充记录</h3>
        <p>回到采集页执行一次演练或正式写入，这里就会出现完整的操作留痕。</p>
        <router-link to="/" class="pill-button primary">去采集</router-link>
      </div>

      <div v-else class="run-list">
        <article
          v-for="run in runs"
          :key="run.id"
          class="run-card card"
          @click="$router.push('/runs/' + run.id)"
        >
          <div class="run-main">
            <div class="run-title">
              <span class="status-tag" :class="run.status">{{ statusText(run.status) }}</span>
              <strong>{{ run.week_start }} ~ {{ run.week_end }}</strong>
              <span class="mode-tag" :class="run.mode">{{ run.mode === "dry_run" ? "演练" : "正式" }}</span>
            </div>
            <p class="run-sub">
              {{ run.author || "全部作者" }} · {{ run.rows_count }} 行 · {{ run.commits_count }} 条提交
              <span v-if="run.sheet_name"> · {{ run.sheet_name }} 第 {{ run.start_row }}–{{ run.end_row }} 行</span>
            </p>
            <p v-if="run.message" class="run-msg">{{ run.message }}</p>
          </div>
          <div class="run-side">
            <span class="run-time">{{ shortTime(run.created_at) }}</span>
            <div class="run-actions" @click.stop>
              <a v-if="run.has_snapshot" class="mini-button" :href="'/api/runs/' + run.id + '/snapshot'">快照</a>
              <a v-if="run.has_audit" class="mini-button" :href="'/api/runs/' + run.id + '/audit'">审计表</a>
              <router-link class="mini-button primary" :to="'/runs/' + run.id">详情</router-link>
            </div>
          </div>
        </article>
      </div>
    </div>
  `
};

const RunDetailPage = {
  props: { id: String },
  data() {
    return { run: null, loading: true, tab: "rows", timer: null, retryOpen: false };
  },
  created() {
    this.load();
  },
  beforeDestroy() {
    clearTimeout(this.timer);
  },
  methods: {
    load() {
      api
        .get("/api/runs/" + this.id)
        .then((res) => {
          this.run = res.data;
          this.loading = false;
          clearTimeout(this.timer);
          if (this.run.status === "running" || this.run.status === "pending") {
            this.timer = setTimeout(this.load, 1800);
          }
        })
        .catch(() => {
          this.loading = false;
        });
    },
    retry() {
      this.retryOpen = false;
      api.post("/api/runs/" + this.id + "/retry").then((res) => {
        this.$notify("已重新启动", "success");
        this.$router.push("/runs/" + res.data.run_id);
      });
    }
  },
  template: `
    <div>
    <div v-if="run">
      <div class="detail-head">
        <button class="mini-button" @click="$router.push('/history')">← 返回记录</button>
        <div class="detail-title">
          <span class="status-tag" :class="run.status">{{ statusText(run.status) }}</span>
          <h1>运行 #{{ run.id }}</h1>
          <p>
            {{ run.week_start }} ~ {{ run.week_end }} · {{ run.author || "全部作者" }} ·
            {{ run.mode === "dry_run" ? "演练模式" : "正式写入" }}
          </p>
        </div>
        <div class="detail-actions">
          <a v-if="run.has_snapshot" class="pill-button ghost" :href="'/api/runs/' + run.id + '/snapshot'">下载快照</a>
          <a v-if="run.has_audit" class="pill-button ghost" :href="'/api/runs/' + run.id + '/audit'">下载审计表</a>
          <button class="pill-button primary" @click="retryOpen = true">重试</button>
        </div>
      </div>

      <div class="stat-grid">
        <div class="stat card"><span>Ticket 行</span><strong>{{ run.rows_count }}</strong></div>
        <div class="stat card"><span>Git 提交</span><strong>{{ run.commits_count }}</strong></div>
        <div class="stat card">
          <span>Excel 行</span>
          <strong>{{ run.start_row ? run.start_row + "–" + run.end_row : "—" }}</strong>
        </div>
        <div class="stat card">
          <span>状态</span>
          <strong style="font-size: 18px">{{ statusText(run.status) }}</strong>
        </div>
      </div>

      <div class="detail-grid">
        <section class="card meta-card">
          <h3>运行信息</h3>
          <dl>
            <dt>模式</dt><dd>{{ run.mode === "dry_run" ? "演练（未保存）" : "正式写入" }}</dd>
            <dt>仓库</dt><dd>{{ run.repos.join("；") || "—" }}</dd>
            <dt>目标文档</dt><dd>{{ run.target_url || "—" }}</dd>
            <dt>工作表</dt><dd>{{ run.sheet_name || "—" }}</dd>
            <dt>表头行</dt><dd>{{ run.header_row || "—" }}</dd>
            <dt>开始</dt><dd>{{ run.started_at || "—" }}</dd>
            <dt>结束</dt><dd>{{ run.finished_at || "—" }}</dd>
            <dt>说明</dt><dd>{{ run.message || "—" }}</dd>
          </dl>
        </section>
        <section class="card log-card">
          <h3>运行日志</h3>
          <pre>{{ run.log || "暂无日志" }}</pre>
        </section>
      </div>

      <section class="card" style="overflow: hidden">
        <div class="tabs">
          <button :class="{ active: tab === 'rows' }" @click="tab = 'rows'">
            汇总行 ({{ run.rows.length }})
          </button>
          <button :class="{ active: tab === 'commits' }" @click="tab = 'commits'">
            Git 提交 ({{ run.commits.length }})
          </button>
        </div>

        <div v-if="tab === 'rows'" class="table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Ticket</th>
                <th>任务描述</th>
                <th>仓库</th>
                <th>日期</th>
                <th class="center">提交</th>
                <th>工时</th>
                <th>进度</th>
                <th>备注</th>
                <th class="center">Excel 行</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="row in run.rows" :key="row.id">
                <td class="index">{{ row.seq }}</td>
                <td><span class="ticket-chip">{{ row.ticket }}</span></td>
                <td class="desc">{{ row.description }}</td>
                <td class="muted">{{ row.repos }}</td>
                <td class="muted">{{ row.date_range }}</td>
                <td class="center">{{ row.commit_count }}</td>
                <td>{{ row.hours || "—" }}</td>
                <td>{{ row.progress || "—" }}</td>
                <td class="muted">{{ row.remark || "—" }}</td>
                <td class="center">{{ row.excel_row || "—" }}</td>
              </tr>
            </tbody>
          </table>
        </div>

        <div v-else class="table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>提交时间</th>
                <th>Ticket</th>
                <th>提交说明</th>
                <th>仓库</th>
                <th>分支</th>
                <th>哈希</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="commit in run.commits" :key="commit.commit_hash">
                <td class="muted">{{ commit.committed_at }}</td>
                <td><span class="ticket-chip">{{ commit.ticket || "—" }}</span></td>
                <td class="desc">{{ commit.subject }}</td>
                <td class="muted">{{ commit.repo }}</td>
                <td class="muted">{{ commit.branch }}</td>
                <td class="muted">{{ commit.commit_hash.slice(0, 8) }}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      <div v-if="retryOpen" class="modal-backdrop" @click.self="retryOpen = false">
        <div class="modal">
          <h3>重新执行这次填充？</h3>
          <p>会使用相同的周区间和汇总行再跑一次。正式模式下可能产生重复行，请先确认表格状态。</p>
          <div class="modal-actions">
            <button class="pill-button ghost" @click="retryOpen = false">取消</button>
            <button class="pill-button primary" @click="retry">确认重试</button>
          </div>
        </div>
      </div>
    </div>

    <div v-else-if="loading" class="card skeleton-card">
      <div class="skeleton" style="width: 38%"></div>
      <div class="skeleton" style="width: 70%; margin-top: 12px"></div>
      <div class="skeleton" style="width: 52%; margin-top: 12px"></div>
    </div>

    <div v-else class="empty-state card">
      <h3>记录不存在</h3>
      <p>这条填充记录可能已被删除。</p>
      <router-link to="/history" class="pill-button primary">返回记录</router-link>
    </div>
    </div>
  `
};

const SettingsPage = {
  data() {
    return {
      config: null,
      samples: [],
      aliasText: {},
      reposText: "",
      footerText: "",
      saving: false,
      inspecting: false,
      inspectResult: null,
      testing: false,
      testResult: null,
      checkingWorkbook: false,
      workbookStatus: null,
      repoStatus: [],
      checkingRepos: false,
      testingAI: false,
      aiTestResult: ""
    };
  },
  computed: {
    fields() {
      return Object.keys(FIELD_LABELS);
    },
    fieldLabels() {
      return FIELD_LABELS;
    },
    workbookNames() {
      if (!this.workbookStatus) {
        return "";
      }
      return (this.workbookStatus.workbooks || [])
        .map(function (item) {
          return item.name;
        })
        .join("、");
    }
  },
  created() {
    this.load();
  },
  methods: {
    load() {
      Promise.all([api.get("/api/config"), api.get("/api/samples")])
        .then((results) => {
          this.config = results[0].data.config;
          this.samples = results[1].data.samples;
          this.reposText = this.config.repos.join("\n");
          this.footerText = this.config.footer_keywords.join(", ");
          const aliasText = {};
          const self = this;
          this.fields.forEach(function (field) {
            aliasText[field] = (self.config.column_aliases[field] || []).join(", ");
          });
          this.aliasText = aliasText;
        })
        .catch(function () {});
    },
    save() {
      this.saving = true;
      const payload = {
        sharepoint_url: this.config.sharepoint_url,
        write_mode: this.config.write_mode,
        target_sheet: this.config.target_sheet,
        target_mode: this.config.target_mode,
        weekly_template_sheet: this.config.weekly_template_sheet,
        weekly_sheet_prefix: this.config.weekly_sheet_prefix,
        weekly_overwrite: this.config.weekly_overwrite,
        header_row: Number(this.config.header_row) || 0,
        author: this.config.author,
        author_display: this.config.author_display,
        date_mode: this.config.date_mode,
        aggregate_mode: this.config.aggregate_mode,
        repos: this.reposText.split("\n"),
        week_start_day: Number(this.config.week_start_day) || 0,
        footer_keywords: this.footerText,
        column_aliases: {},
        dry_run: this.config.dry_run,
        copy_row_format: this.config.copy_row_format,
        insert_before_footer: this.config.insert_before_footer,
        ignore_merge_commits: this.config.ignore_merge_commits,
        reuse_empty_rows: this.config.reuse_empty_rows,
        max_description_chars: Number(this.config.max_description_chars) || 1000,
        sample_path: this.config.sample_path,
        ai_enabled: this.config.ai_enabled,
        ai_base_url: this.config.ai_base_url,
        ai_api_key: this.config.ai_api_key,
        ai_model: this.config.ai_model,
        ai_prompt: this.config.ai_prompt,
        ai_timeout: Number(this.config.ai_timeout) || 60,
        ai_temperature: Number(this.config.ai_temperature) || 0.3
      };
      const self = this;
      this.fields.forEach(function (field) {
        payload.column_aliases[field] = splitList(self.aliasText[field]);
      });
      api
        .put("/api/config", payload)
        .then((res) => {
          this.config = res.data.config;
          this.reposText = this.config.repos.join("\n");
          this.footerText = this.config.footer_keywords.join(", ");
          this.$notify("设置已保存", "success");
        })
        .finally(() => {
          this.saving = false;
        });
    },
    checkRepos() {
      this.checkingRepos = true;
      api
        .post("/api/repos/status", {
          repos: this.reposText.split("\n")
        })
        .then((res) => {
          this.repoStatus = res.data.repos;
          const bad = this.repoStatus.filter(function (repo) {
            return !repo.is_git;
          });
          this.$notify(
            bad.length ? bad.length + " 个仓库不可用" : "全部仓库可用",
            bad.length ? "error" : "success"
          );
        })
        .finally(() => {
          this.checkingRepos = false;
        });
    },
    checkWorkbookStatus() {
      this.checkingWorkbook = true;
      api
        .get("/api/workbook/status")
        .then((res) => {
          this.workbookStatus = res.data;
          this.$notify(
            res.data.found ? "已找到目标工作簿" : "未找到目标工作簿",
            res.data.found ? "success" : "error"
          );
        })
        .finally(() => {
          this.checkingWorkbook = false;
        });
    },
    inspectLocal() {
      this.inspecting = true;
      api
        .post("/api/template/inspect", { path: this.config.sample_path })
        .then((res) => {
          this.inspectResult = res.data;
          this.$notify("模板解析完成", "success");
        })
        .finally(() => {
          this.inspecting = false;
        });
    },
    testSharepoint() {
      this.testing = true;
      api
        .post("/api/sharepoint/inspect")
        .then((res) => {
          this.testResult = res.data;
          this.$notify("已连接在线文档", "success");
        })
        .finally(() => {
          this.testing = false;
        });
    },
    testAI() {
      this.testingAI = true;
      this.aiTestResult = "";
      api
        .post("/api/ai/test")
        .then((res) => {
          this.aiTestResult = "连接成功：" + (res.data.reply || "").slice(0, 120);
          this.$notify("AI 连接成功", "success");
        })
        .catch(() => {
          this.aiTestResult = "连接失败，请检查 Base URL、Key 和模型";
        })
        .finally(() => {
          this.testingAI = false;
        });
    },
    uploadSample(event) {
      const file = event.target.files[0];
      if (!file) {
        return;
      }
      const form = new FormData();
      form.append("file", file);
      api
        .post("/api/samples/upload", form)
        .then((res) => {
          this.config.sample_path = res.data.sample.path;
          return api.get("/api/samples");
        })
        .then((res) => {
          this.samples = res.data.samples;
          this.$notify("模板已上传", "success");
        });
      event.target.value = "";
    },
    applyDetected(sheet) {
      this.config.target_sheet = sheet.name;
      this.config.header_row = sheet.header_row || 0;
      this.$notify(
        "已应用「" + sheet.name + "」第 " + sheet.header_row + " 行表头，记得保存",
        "success"
      );
    }
  },
  template: `
    <div>
    <div v-if="config">
      <div class="page-head">
        <div>
          <p class="eyebrow">CONFIGURATION</p>
          <h1>设置</h1>
          <p>配置 SharePoint 目标、Git 数据源和字段映射，保存后立即生效。</p>
        </div>
        <button class="pill-button ghost" @click="checkRepos" :disabled="checkingRepos">
          {{ checkingRepos ? "检测中…" : "检测仓库" }}
        </button>
      </div>

      <div class="settings-layout">
        <section class="settings-group card">
          <header>
            <h2>SharePoint 目标</h2>
            <p>优先由程序打开文档链接（推荐）；附着模式仅作备用，效果不稳定。</p>
          </header>
          <div class="form-grid">
            <label class="form-row full">
              <span>在线文档链接</span>
              <input v-model="config.sharepoint_url" placeholder="https://...">
              <small class="field-hint">请使用 SharePoint 文件的直接路径（形如 .../Shared Documents/.../xxx.xlsx），不要用 doc.aspx 查看器链接。</small>
            </label>
            <label class="form-row full">
              <span>写入方式</span>
              <select v-model="config.write_mode">
                <option value="open">由程序打开文档链接（推荐）</option>
                <option value="attach">附着到已打开的 Excel（不推荐）</option>
              </select>
              <small class="field-hint">推荐打开模式：程序启动隐藏的 Excel 直接打开上面的文件路径，写入后保存。附着模式需先在 Excel 中打开周报，效果不稳定，仅备用。</small>
            </label>
            <label class="form-row full">
              <span>写入目标</span>
              <select v-model="config.target_mode">
                <option value="existing">写入现有工作表（日报/汇总）</option>
                <option value="weekly">每周新建周报汇总 sheet</option>
                <option value="both">日报 + 每周新建周报汇总 sheet（两者都写）</option>
              </select>
              <small class="field-hint">周报汇总/两者都写会复制模板 sheet 新建周报汇总；两者都写会先写日报，再写周报汇总。</small>
            </label>
            <label class="form-row" v-if="config.target_mode !== 'weekly'">
              <span>{{ config.target_mode === 'both' ? '日报工作表' : '工作表名称' }}</span>
              <input v-model="config.target_sheet" placeholder="留空则使用当前活动工作表">
            </label>
            <label class="form-row" v-if="config.target_mode !== 'existing'">
              <span>模板工作表</span>
              <input v-model="config.weekly_template_sheet" placeholder="留空自动找最后一个含「周报汇总」的 sheet">
            </label>
            <label class="form-row" v-if="config.target_mode !== 'existing'">
              <span>新 sheet 前缀</span>
              <input v-model="config.weekly_sheet_prefix" placeholder="周报汇总">
              <small class="field-hint">新 sheet 名 = 前缀 + 起始日期，如「周报汇总 10.19」。</small>
            </label>
            <div class="form-row" v-if="config.target_mode !== 'existing'">
              <div class="switch-line">
                <div>
                  <strong>同周重复写入时覆盖</strong>
                  <small>关闭后会在名称后加 (2)、(3)。</small>
                </div>
                <label class="toggle">
                  <input type="checkbox" v-model="config.weekly_overwrite">
                  <i></i>
                </label>
              </div>
            </div>
            <label class="form-row">
              <span>表头行</span>
              <input type="number" min="0" v-model.number="config.header_row">
              <small class="field-hint">0 = 自动识别。</small>
            </label>
            <div class="form-row full">
              <div class="switch-line">
                <div>
                  <strong>默认演练模式</strong>
                  <small>只分析表头与插入位置，不保存修改。</small>
                </div>
                <label class="toggle">
                  <input type="checkbox" v-model="config.dry_run">
                  <i></i>
                </label>
              </div>
            </div>
          </div>
        </section>

        <section class="settings-group card">
          <header>
            <h2>Excel 连接</h2>
            <p>附着模式（仅备用）下，程序会找到你已经在 Excel 中打开的周报。</p>
          </header>
          <div
            class="workbook-check"
            :class="{ ok: workbookStatus && workbookStatus.found, bad: workbookStatus && !workbookStatus.found }"
          >
            <div class="workbook-check-text">
              <strong v-if="!workbookStatus">尚未检测</strong>
              <strong v-else-if="workbookStatus.found">已找到：{{ workbookStatus.workbook.name }}</strong>
              <strong v-else-if="!workbookStatus.excel_running">未检测到正在运行的 Excel</strong>
              <strong v-else>Excel 已运行，但未找到目标工作簿</strong>
              <small v-if="workbookStatus && workbookNames">当前打开：{{ workbookNames }}</small>
              <small v-else-if="!workbookStatus">点击右侧按钮检测当前 Excel 中打开的工作簿。</small>
              <small v-else-if="!workbookStatus.excel_running && config.write_mode === 'attach'">
                附着模式需要 Excel 已正常打开；优先切换到「由程序打开文档链接」，不需要保持 Excel 打开。
              </small>
            </div>
            <button class="mini-button" @click="checkWorkbookStatus" :disabled="checkingWorkbook">
              {{ checkingWorkbook ? "检测中…" : "检测" }}
            </button>
          </div>
          <p class="field-hint">
            <template v-if="config.write_mode === 'attach'">
              附着模式效果不稳定，仅备用。先在 Excel 中正常打开周报（不要用 SharePoint 网页预览窗口），再回来检测。
            </template>
            <template v-else>
              当前为「由程序打开文档链接」模式（推荐），不需要保持 Excel 打开；附着模式才需要检测。
            </template>
          </p>
        </section>

        <section class="settings-group card">
          <header>
            <h2>Git 数据源</h2>
            <p>支持多个本地仓库，按作者模糊匹配，统计所选周（默认周一至周日）。</p>
          </header>
          <div class="form-grid">
            <label class="form-row">
              <span>作者过滤</span>
              <input v-model="config.author" placeholder="留空统计所有人">
            </label>
            <label class="form-row">
              <span>执行人显示名</span>
              <input v-model="config.author_display" placeholder="留空则使用 Git 作者名">
              <small class="field-hint">写入周报的「姓名 / 执行人」列。</small>
            </label>
            <label class="form-row">
              <span>每周起始日</span>
              <select v-model.number="config.week_start_day">
                <option :value="0">周一</option>
                <option :value="6">周日</option>
              </select>
            </label>
            <label class="form-row full">
              <span>仓库路径（每行一个）</span>
              <textarea v-model="reposText"></textarea>
            </label>
          </div>
          <div class="repo-status-list" v-if="repoStatus.length">
            <div v-for="repo in repoStatus" :key="repo.path" class="repo-status">
              <span class="name">{{ repo.name }}</span>
              <span class="path">{{ repo.path }}</span>
              <span v-if="repo.is_git" class="ok">正常</span>
              <span v-else class="bad">{{ repo.error || "不可用" }}</span>
              <span v-if="repo.last_commit" class="meta">{{ shortTime(repo.last_commit.committed_at) }}</span>
            </div>
          </div>
        </section>

        <section class="settings-group card">
          <header>
            <h2>字段映射</h2>
            <p>填写周报表头里可能出现的名称，用逗号分隔；程序会自动匹配列。</p>
          </header>
          <div class="alias-grid">
            <label class="form-row" v-for="field in fields" :key="field">
              <span>{{ fieldLabels[field] }}</span>
              <input v-model="aliasText[field]">
            </label>
          </div>
        </section>

        <section class="settings-group card">
          <header>
            <h2>写入细节</h2>
            <p>控制插入位置、格式复用与描述长度。</p>
          </header>
          <div class="form-grid">
            <label class="form-row">
              <span>页脚关键词</span>
              <input v-model="footerText" placeholder="合计, 总计, 小计">
              <small class="field-hint">遇到这些开头的行时，把新行插在它前面。</small>
            </label>
            <label class="form-row full">
              <span>汇总粒度</span>
              <select v-model="config.aggregate_mode" :disabled="config.target_mode !== 'existing'">
                <option value="daily">按 Ticket + 日期（日报，一天可多条）</option>
                <option value="ticket">按 Ticket（整个范围一行）</option>
              </select>
              <small class="field-hint" v-if="config.target_mode === 'weekly'">周报汇总模式固定按 Ticket 汇总。</small>
              <small class="field-hint" v-else-if="config.target_mode === 'both'">日报 + 周报汇总：日报按 Ticket + 日期，周报汇总按 Ticket。</small>
              <small class="field-hint" v-else>日报建议用「按 Ticket + 日期」：同一天不同 Ticket 写成多行；同一 Ticket 同一天合并为一行。</small>
            </label>
            <label class="form-row">
              <span>日期写入方式</span>
              <select v-model="config.date_mode" :disabled="config.target_mode !== 'existing'">
                <option value="last">最后提交日</option>
                <option value="first">首次提交日</option>
                <option value="week_start">范围起始日</option>
                <option value="range">日期范围（文本）</option>
              </select>
              <small class="field-hint" v-if="config.target_mode === 'weekly'">周报汇总 sheet 没有日期列。</small>
              <small class="field-hint" v-else-if="config.target_mode === 'both'">日报按提交当天写；周报汇总无日期列。</small>
              <small class="field-hint" v-else>仅在「按 Ticket」粒度下生效；按 Ticket + 日期时写提交当天。</small>
            </label>
            <label class="form-row">
              <span>描述最大长度</span>
              <input type="number" min="100" v-model.number="config.max_description_chars">
            </label>
            <div class="form-row">
              <div class="switch-line">
                <div>
                  <strong>复用未填充的日期行</strong>
                  <small>优先从表格末尾的空日期行开始写入。</small>
                </div>
                <label class="toggle">
                  <input type="checkbox" v-model="config.reuse_empty_rows">
                  <i></i>
                </label>
              </div>
            </div>
            <div class="form-row">
              <div class="switch-line">
                <div>
                  <strong>复用上一行格式</strong>
                  <small>复制边框、字体与底色。</small>
                </div>
                <label class="toggle">
                  <input type="checkbox" v-model="config.copy_row_format">
                  <i></i>
                </label>
              </div>
            </div>
            <div class="form-row">
              <div class="switch-line">
                <div>
                  <strong>插入到页脚之前</strong>
                  <small>保持合计行在表格底部。</small>
                </div>
                <label class="toggle">
                  <input type="checkbox" v-model="config.insert_before_footer">
                  <i></i>
                </label>
              </div>
            </div>
            <div class="form-row">
              <div class="switch-line">
                <div>
                  <strong>忽略 Merge 提交</strong>
                  <small>避免「Merge branch…」进入未分类。</small>
                </div>
                <label class="toggle">
                  <input type="checkbox" v-model="config.ignore_merge_commits">
                  <i></i>
                </label>
              </div>
            </div>
          </div>
        </section>

        <section class="settings-group card">
          <header>
            <h2>AI 改写</h2>
            <p>调用 OpenAI 兼容接口，把 Git 提交改写成适合周报的简洁任务描述。失败时保留原始提交说明。</p>
          </header>
          <div class="form-grid">
            <div class="form-row full">
              <div class="switch-line">
                <div>
                  <strong>启用 AI 改写</strong>
                  <small>采集时自动改写任务描述。</small>
                </div>
                <label class="toggle">
                  <input type="checkbox" v-model="config.ai_enabled">
                  <i></i>
                </label>
              </div>
            </div>
            <label class="form-row full">
              <span>Base URL</span>
              <input v-model="config.ai_base_url" placeholder="https://opencode.ai/zen/v1">
              <small class="field-hint">OpenCode Zen / Go 使用 https://opencode.ai/zen/v1，程序会自动追加 /chat/completions。</small>
            </label>
            <label class="form-row full">
              <span>API Key</span>
              <input type="password" v-model="config.ai_api_key" placeholder="sk-...">
            </label>
            <label class="form-row">
              <span>模型</span>
              <input v-model="config.ai_model" placeholder="deepseek-v4.1-flash">
            </label>
            <label class="form-row">
              <span>超时（秒）</span>
              <input type="number" min="10" v-model.number="config.ai_timeout">
            </label>
            <label class="form-row">
              <span>温度</span>
              <input type="number" step="0.1" min="0" max="2" v-model.number="config.ai_temperature">
            </label>
            <label class="form-row full">
              <span>系统提示词</span>
              <textarea v-model="config.ai_prompt" style="min-height: 150px"></textarea>
            </label>
            <div class="form-row full" style="flex-direction: row; align-items: center; gap: 12px; flex-wrap: wrap">
              <button class="mini-button" @click="testAI" :disabled="testingAI">
                {{ testingAI ? "测试中…" : "测试 AI 连接" }}
              </button>
              <span class="field-hint">修改后先保存设置，再测试连接。</span>
              <span v-if="aiTestResult" class="field-hint">{{ aiTestResult }}</span>
            </div>
          </div>
        </section>

        <section class="settings-group card">
          <header>
            <h2>模板检查</h2>
            <p>上传一份周报副本，解析工作表、表头和前几行内容，用来校准字段映射。</p>
          </header>
          <div class="form-grid">
            <label class="form-row full">
              <span>本地模板路径</span>
              <input v-model="config.sample_path" placeholder="D:\\...\\samples\\周报.xlsx">
            </label>
            <div class="form-row full" style="flex-direction: row; align-items: center; gap: 12px; flex-wrap: wrap">
              <label class="upload-label">
                上传模板
                <input type="file" class="hidden-input" accept=".xlsx,.xlsm" @change="uploadSample">
              </label>
              <select v-model="config.sample_path" style="max-width: 420px">
                <option value="">选择已上传的模板…</option>
                <option v-for="sample in samples" :key="sample.path" :value="sample.path">{{ sample.name }}</option>
              </select>
              <button class="mini-button" @click="inspectLocal" :disabled="inspecting">
                {{ inspecting ? "解析中…" : "解析本地模板" }}
              </button>
              <button class="mini-button" @click="testSharepoint" :disabled="testing">
                {{ testing ? "检测中…" : (config.write_mode === "attach" ? "检测已打开的工作簿" : "测试在线连接") }}
              </button>
            </div>
          </div>
          <p class="field-hint" style="margin-top: 10px">
            <template v-if="config.write_mode === 'attach'">
              「检测已打开的工作簿」会读取当前 Excel 中打开的周报，不会启动新的 Excel 进程。
            </template>
            <template v-else>
              「测试在线连接」会启动一个隐藏的 Excel 进程，首次可能需要登录公司账号，通常需要 10–30 秒。
            </template>
          </p>
          <inspect-view v-if="inspectResult" :result="inspectResult" @apply="applyDetected"></inspect-view>
          <inspect-view v-if="testResult" :result="testResult" @apply="applyDetected"></inspect-view>
        </section>

        <div class="save-bar">
          <span>修改后点击保存。正式写入前建议先用演练模式验证。</span>
          <button class="pill-button primary" @click="save" :disabled="saving">
            {{ saving ? "保存中…" : "保存设置" }}
          </button>
        </div>
      </div>
    </div>

    <div v-else class="card skeleton-card">
      <div class="skeleton" style="width: 40%"></div>
      <div class="skeleton" style="width: 66%; margin-top: 12px"></div>
    </div>
    </div>
  `
};

const App = {
  data() {
    return { toast: toastState, meta: { excel_com: true } };
  },
  created() {
    api
      .get("/api/meta")
      .then((res) => {
        this.meta = res.data;
      })
      .catch(function () {});
  },
  template: `
    <div class="shell">
      <header class="topbar">
        <router-link to="/" class="brand">
          <span class="brand-mark">W</span>
          <span class="brand-text">
            <strong>WeekLog</strong>
            <small>Git → 周报</small>
          </span>
        </router-link>
        <nav class="nav">
          <router-link to="/" exact>采集填充</router-link>
          <router-link to="/history">填充记录</router-link>
          <router-link to="/settings">设置</router-link>
        </nav>
        <div class="topbar-right">
          <span class="env-dot" :class="{ off: !meta.excel_com }"></span>
          <span>{{ meta.excel_com ? "Excel COM 就绪" : "Excel COM 不可用" }}</span>
        </div>
      </header>
      <main class="page">
        <transition name="page" mode="out-in">
          <router-view :key="$route.fullPath"></router-view>
        </transition>
      </main>
      <transition name="toast">
        <div v-if="toast.visible" class="toast" :class="toast.type">{{ toast.message }}</div>
      </transition>
    </div>
  `
};

const router = new VueRouter({
  mode: "history",
  routes: [
    { path: "/", component: HomePage },
    { path: "/history", component: HistoryPage },
    { path: "/runs/:id", component: RunDetailPage, props: true },
    { path: "/settings", component: SettingsPage },
    { path: "*", redirect: "/" }
  ]
});

new Vue({
  router: router,
  render: function (h) {
    return h(App);
  }
}).$mount("#app");
