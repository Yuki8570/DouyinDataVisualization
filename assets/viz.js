/* Shared Douyin / Shipinhao Excel parsers + dashboard renderer */
(function (global) {
  const fmt = (n) => Number(n || 0).toLocaleString("zh-CN");
  const pct = (n) => `${((Number(n) || 0) * 100).toFixed(1)}%`;

  function shortTitle(title, n = 18) {
    const base = String(title || "")
      .split("#")[0]
      .trim()
      .replace(/[！!？?。．\s]+$/g, "");
    return base.length <= n ? base : base.slice(0, n);
  }

  function toNum(v) {
    if (v === null || v === undefined || v === "" || v === "-") return null;
    if (typeof v === "number") return v;
    const s = String(v).replace(/,/g, "").trim();
    const m = s.match(/-?\d+(\.\d+)?/);
    return m ? Number(m[0]) : null;
  }

  function parseSeconds(v) {
    if (v === null || v === undefined || v === "" || v === "-") return 0;
    if (typeof v === "number") return v;
    return toNum(String(v).replace("秒", "")) || 0;
  }

  function excelDateToString(v) {
    if (v == null || v === "") return "";
    if (v instanceof Date && !isNaN(v)) {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, "0");
      const d = String(v.getDate()).padStart(2, "0");
      const hh = String(v.getHours()).padStart(2, "0");
      const mm = String(v.getMinutes()).padStart(2, "0");
      const ss = String(v.getSeconds()).padStart(2, "0");
      return `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
    }
    if (typeof v === "number" && global.XLSX && XLSX.SSF) {
      try {
        const parsed = XLSX.SSF.parse_date_code(v);
        if (parsed) {
          const y = parsed.y;
          const m = String(parsed.m).padStart(2, "0");
          const d = String(parsed.d).padStart(2, "0");
          const hh = String(parsed.H || 0).padStart(2, "0");
          const mm = String(parsed.M || 0).padStart(2, "0");
          const ss = String(parsed.S || 0).padStart(2, "0");
          return `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
        }
      } catch (_) {}
    }
    return String(v);
  }

  function dateLabel(datetime) {
    const s = String(datetime || "");
    const m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[2]}-${m[3]}`;
    return s.slice(5, 10) || s;
  }

  function detectFormat(headers) {
    const h = (headers || []).map((x) => String(x || "").trim().replace(/^"|"$/g, ""));
    if (h.includes("视频描述") && h.includes("视频ID")) return "shipinhao";
    if (h.includes("作品名称") && h.includes("播放量")) return "douyin";
    if (h.includes("日期") && h.includes("总播放量") && (h.includes("投稿量") || h.includes("总点赞量")))
      return "jiangwei_douyin";
    if (h.includes("时间") && h.includes("播放") && h.includes("喜欢") && h.includes("推荐"))
      return "jiangwei_shipinhao";
    if (h.includes("视频描述") || (h.includes("喜欢") && h.includes("关注量")))
      return "shipinhao";
    return null;
  }

  function findHeaderRow(rows) {
    for (let i = 0; i < Math.min(rows.length, 12); i++) {
      const detected = detectFormat(rows[i]);
      if (detected) return { index: i, format: detected };
    }
    return null;
  }

  function headerIndex(headers) {
    const map = {};
    headers.forEach((h, i) => {
      map[String(h || "").trim().replace(/^"|"$/g, "")] = i;
    });
    return map;
  }

  function normalizeDateString(v) {
    const raw = excelDateToString(v);
    const s = String(raw || "").trim().replace(/\//g, "-");
    const m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) {
      return `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(m[3]).padStart(2, "0")}`;
    }
    return s.slice(0, 10);
  }

  function parsePctCell(v) {
    if (v === null || v === undefined || v === "" || v === "-") return 0;
    if (typeof v === "number") return v > 1 ? v / 100 : v;
    const n = toNum(String(v).replace("%", ""));
    if (n == null) return 0;
    return n > 1 ? n / 100 : n;
  }

  function parseDouyinRows(rows) {
    if (!rows || rows.length < 2) return [];
    const idx = headerIndex(rows[0]);
    const get = (r, name) => r[idx[name]];
    return rows.slice(1).filter((r) => r && r[idx["作品名称"]]).map((r) => {
      const fullTitle = String(get(r, "作品名称") || "");
      const datetime = excelDateToString(get(r, "发布时间"));
      return {
        platform: "douyin",
        kind: "works",
        title: shortTitle(fullTitle),
        fullTitle,
        date: dateLabel(datetime),
        datetime,
        genre: String(get(r, "体裁") || ""),
        status: String(get(r, "审核状态") || ""),
        views: Math.round(toNum(get(r, "播放量")) || 0),
        finishRate: toNum(get(r, "完播率")) || 0,
        finish5s: toNum(get(r, "5s完播率")) || toNum(get(r, "5秒完播率")) || 0,
        coverCtr: toNum(get(r, "封面点击率")),
        bounce2s: toNum(get(r, "2s跳出率")) || toNum(get(r, "2秒跳出率")) || 0,
        avgWatchSec: toNum(get(r, "平均播放时长")) || 0,
        likes: Math.round(toNum(get(r, "点赞量")) || 0),
        shares: Math.round(toNum(get(r, "分享量")) || 0),
        comments: Math.round(toNum(get(r, "评论量")) || 0),
        favorites: Math.round(toNum(get(r, "收藏量")) || 0),
        profileVisits: Math.round(toNum(get(r, "主页访问量")) || 0),
        followerDelta: Math.round(toNum(get(r, "粉丝增量")) || 0),
      };
    });
  }

  function parseShipinhaoRows(rows) {
    if (!rows || rows.length < 2) return [];
    const idx = headerIndex(rows[0]);
    const get = (r, name) => r[idx[name]];
    return rows.slice(1).filter((r) => r && r[idx["视频描述"]]).map((r) => {
      const fullTitle = String(get(r, "视频描述") || "");
      const datetime = excelDateToString(get(r, "发布时间"));
      return {
        platform: "shipinhao",
        kind: "works",
        title: shortTitle(fullTitle),
        fullTitle,
        videoId: String(get(r, "视频ID") || ""),
        date: dateLabel(datetime),
        datetime,
        finishRate: toNum(get(r, "完播率")) || 0,
        avgWatchSec: parseSeconds(get(r, "平均播放时长")),
        views: Math.round(toNum(get(r, "播放量")) || 0),
        recommend: Math.round(toNum(get(r, "推荐")) || 0),
        likes: Math.round(toNum(get(r, "喜欢")) || 0),
        comments: Math.round(toNum(get(r, "评论量")) || 0),
        shares: Math.round(toNum(get(r, "分享量")) || 0),
        follows: Math.round(toNum(get(r, "关注量")) || 0),
        forwardChat: Math.round(toNum(get(r, "转发聊天和朋友圈")) || 0),
        setRingtone: Math.round(toNum(get(r, "设为铃声")) || 0),
        setStatus: Math.round(toNum(get(r, "设为状态")) || 0),
        setMomentsCover: Math.round(toNum(get(r, "设为朋友圈封面")) || 0),
        wecomClicks: Math.round(toNum(get(r, "企微链接点击次数")) || 0),
        wecomUsers: Math.round(toNum(get(r, "企微链接点击人数")) || 0),
        contactAdds: Math.round(toNum(get(r, "添加到通讯录次数")) || 0),
        contactUsers: Math.round(toNum(get(r, "添加到通讯录人数")) || 0),
      };
    });
  }

  function parseJiangweiDouyinDaily(rows) {
    if (!rows || rows.length < 2) return [];
    const idx = headerIndex(rows[0]);
    const get = (r, name) => r[idx[name]];
    return rows
      .slice(1)
      .filter((r) => r && get(r, "日期") != null && String(get(r, "日期")).trim() !== "")
      .map((r) => {
        const datetime = normalizeDateString(get(r, "日期"));
        return {
          platform: "jiangwei_douyin",
          kind: "daily",
          date: datetime,
          datetime,
          posts: Math.round(toNum(get(r, "投稿量")) || 0),
          views: Math.round(toNum(get(r, "总播放量")) || 0),
          likes: Math.round(toNum(get(r, "总点赞量")) || 0),
          shares: Math.round(toNum(get(r, "总分享量")) || 0),
          comments: Math.round(toNum(get(r, "总评论量")) || 0),
          finish5s: parsePctCell(get(r, "5秒完播率")),
          bounce2s: parsePctCell(get(r, "2秒跳出率")),
          coverCtr: parsePctCell(get(r, "封面点击率")),
          avgWatchSec: parseSeconds(get(r, "平均播放时长")),
        };
      })
      .sort((a, b) => a.date.localeCompare(b.date));
  }

  function parseJiangweiShipinhaoDaily(rows) {
    if (!rows || rows.length < 2) return [];
    const idx = headerIndex(rows[0]);
    const get = (r, name) => r[idx[name]];
    return rows
      .slice(1)
      .filter((r) => r && get(r, "时间") != null && String(get(r, "时间")).trim() !== "")
      .map((r) => {
        const datetime = normalizeDateString(get(r, "时间"));
        return {
          platform: "jiangwei_shipinhao",
          kind: "daily",
          date: datetime,
          datetime,
          views: Math.round(toNum(get(r, "播放")) || 0),
          recommend: Math.round(toNum(get(r, "推荐")) || 0),
          likes: Math.round(toNum(get(r, "喜欢")) || 0),
          comments: Math.round(toNum(get(r, "评论")) || 0),
          shares: Math.round(toNum(get(r, "分享")) || 0),
          follows: Math.round(toNum(get(r, "关注")) || 0),
        };
      })
      .sort((a, b) => a.date.localeCompare(b.date));
  }

  function parseRowsByFormat(format, rows) {
    if (format === "douyin") return parseDouyinRows(rows);
    if (format === "shipinhao") return parseShipinhaoRows(rows);
    if (format === "jiangwei_douyin") return parseJiangweiDouyinDaily(rows);
    if (format === "jiangwei_shipinhao") return parseJiangweiShipinhaoDaily(rows);
    return [];
  }

  function parseWorkbook(workbook) {
    const sheets = {};
    let platform = null;
    workbook.SheetNames.forEach((name) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
        header: 1,
        defval: null,
        raw: true,
      });
      if (!rows.length) return;
      const found = findHeaderRow(rows);
      if (!found) return;
      if (!platform) platform = found.format;
      if (found.format !== platform) return;
      const sliced = [rows[found.index]].concat(rows.slice(found.index + 1));
      const posts = parseRowsByFormat(found.format, sliced);
      if (posts.length) sheets[name] = posts;
    });
    if (!platform || !Object.keys(sheets).length) {
      throw new Error(
        "无法识别表格格式。支持：公司号作品明细（抖音/视频号）、江炜日维度数据（抖音全量指标 / 视频号详情 CSV）。"
      );
    }
    return { platform, sheets, kind: platform.startsWith("jiangwei_") ? "daily" : "works" };
  }

  function isDailyState(state) {
    if (state && state.kind === "daily") return true;
    if (state && String(state.platform || "").startsWith("jiangwei_")) return true;
    const sample = state && state.posts && state.posts[0];
    return !!(sample && sample.kind === "daily");
  }

  function engage(p) {
    if (p.platform === "shipinhao") {
      return ((p.likes + p.shares + p.comments + p.follows) / Math.max(p.views, 1)) * 1000;
    }
    return (
      ((p.likes + p.shares + p.comments + p.favorites) / Math.max(p.views, 1)) * 1000
    );
  }

  function sortPosts(posts, key) {
    const list = [...posts];
    if (key === "views") return list.sort((a, b) => b.views - a.views);
    if (key === "engage") return list.sort((a, b) => engage(b) - engage(a));
    if (key === "finish") return list.sort((a, b) => b.finishRate - a.finishRate);
    if (key === "profile")
      return list.sort(
        (a, b) => (b.profileVisits || b.follows || 0) - (a.profileVisits || a.follows || 0)
      );
    if (key === "fans")
      return list.sort(
        (a, b) => (b.followerDelta || b.follows || 0) - (a.followerDelta || a.follows || 0)
      );
    return list.sort((a, b) => String(a.datetime).localeCompare(String(b.datetime)));
  }

  function linePath(points) {
    return points
      .map((p, i) => (i === 0 ? `M ${p.x} ${p.y}` : `L ${p.x} ${p.y}`))
      .join(" ");
  }
  function areaPath(points, baselineY) {
    if (!points.length) return "";
    return `${linePath(points)} L ${points[points.length - 1].x} ${baselineY} L ${points[0].x} ${baselineY} Z`;
  }

  function renderDonut(svgId, legendId, items, colors) {
    const svg = document.getElementById(svgId);
    const legend = document.getElementById(legendId);
    if (!svg || !legend) return;
    const filtered = items.filter((x) => x.value > 0);
    const total = filtered.reduce((s, x) => s + x.value, 0) || 1;
    const cx = 80, cy = 80, r = 58, stroke = 22;
    let angle = -Math.PI / 2;
    const arcs = filtered
      .map((item, i) => {
        const slice = (item.value / total) * Math.PI * 2;
        const a1 = angle;
        const a2 = angle + slice;
        angle = a2;
        const large = slice > Math.PI ? 1 : 0;
        const x1 = cx + r * Math.cos(a1);
        const y1 = cy + r * Math.sin(a1);
        const x2 = cx + r * Math.cos(a2);
        const y2 = cy + r * Math.sin(a2);
        return `<path d="M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2}" fill="none" stroke="${colors[i % colors.length]}" stroke-width="${stroke}"><title>${item.label}: ${item.value}</title></path>`;
      })
      .join("");
    svg.innerHTML = `
      <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#e4ebe7" stroke-width="${stroke}" />
      ${arcs}
      <text x="${cx}" y="${cy - 4}" text-anchor="middle" font-family="Syne, sans-serif" font-size="20" font-weight="700" fill="#132820">${fmt(total)}</text>
      <text x="${cx}" y="${cy + 14}" text-anchor="middle" font-size="11" fill="#6b7f74">合计</text>
    `;
    legend.innerHTML = filtered
      .map(
        (item, i) => `
      <div>
        <span><span class="swatch" style="background:${colors[i % colors.length]}"></span>${item.label}</span>
        <strong>${fmt(item.value)}（${((item.value / total) * 100).toFixed(0)}%）</strong>
      </div>`
      )
      .join("");
  }

  function renderTrend(posts) {
    const chrono = sortPosts(posts, "date");
    const svg = document.getElementById("trendChart");
    if (!svg || !chrono.length) return;
    const W = 560, H = 220, pad = { t: 16, r: 16, b: 36, l: 52 };
    const iw = W - pad.l - pad.r, ih = H - pad.t - pad.b;
    const maxV = Math.max(...chrono.map((p) => p.views), 1) * 1.08;
    const pts = chrono.map((p, i) => ({
      x: pad.l + (i / Math.max(chrono.length - 1, 1)) * iw,
      y: pad.t + ih - (p.views / maxV) * ih,
      label: p.date,
      v: p.views,
    }));
    let grid = "";
    for (let i = 0; i <= 4; i++) {
      const y = pad.t + (ih * i) / 4;
      const val = Math.round(maxV * (1 - i / 4));
      grid += `<line x1="${pad.l}" y1="${y}" x2="${W - pad.r}" y2="${y}" stroke="rgba(19,40,32,0.12)" /><text class="axis" x="${pad.l - 8}" y="${y + 4}" text-anchor="end">${fmt(val)}</text>`;
    }
    const labels = pts
      .map((p, i) =>
        i % 2 === 0 || i === pts.length - 1
          ? `<text class="axis" x="${p.x}" y="${H - 12}" text-anchor="middle">${p.label}</text>`
          : ""
      )
      .join("");
    const dots = pts
      .map(
        (p) =>
          `<circle class="dot" cx="${p.x}" cy="${p.y}" r="3.5"><title>${p.label}: ${fmt(p.v)}</title></circle>`
      )
      .join("");
    svg.innerHTML = `
      <defs>
        <linearGradient id="areaFill" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="#0f6b5c" stop-opacity="0.28"/>
          <stop offset="100%" stop-color="#0f6b5c" stop-opacity="0.02"/>
        </linearGradient>
      </defs>
      <g>${grid}</g>
      <path class="area-views" d="${areaPath(pts, pad.t + ih)}" />
      <path class="line-views" d="${linePath(pts)}" />
      ${dots}${labels}
    `;
  }

  function renderRates(posts) {
    const chrono = sortPosts(posts, "date");
    const svg = document.getElementById("rateChart");
    if (!svg || !chrono.length) return;
    const W = 400, H = 220, pad = { t: 16, r: 12, b: 36, l: 40 };
    const iw = W - pad.l - pad.r, ih = H - pad.t - pad.b;
    const max = 70;
    const toPts = (getter) =>
      chrono.map((p, i) => ({
        x: pad.l + (i / Math.max(chrono.length - 1, 1)) * iw,
        y: pad.t + ih - ((getter(p) * 100) / max) * ih,
        label: p.date,
      }));
    const finish = toPts((p) => p.finishRate);
    const bounce = posts[0] && "bounce2s" in posts[0]
      ? toPts((p) => p.bounce2s || 0)
      : null;
    let grid = "";
    for (let i = 0; i <= 4; i++) {
      const y = pad.t + (ih * i) / 4;
      const val = Math.round(max * (1 - i / 4));
      grid += `<line x1="${pad.l}" y1="${y}" x2="${W - pad.r}" y2="${y}" stroke="rgba(19,40,32,0.12)" /><text class="axis" x="${pad.l - 6}" y="${y + 4}" text-anchor="end">${val}%</text>`;
    }
    const labels = finish
      .map((p, i) =>
        i % 2 === 0 || i === finish.length - 1
          ? `<text class="axis" x="${p.x}" y="${H - 12}" text-anchor="middle">${p.label}</text>`
          : ""
      )
      .join("");
    svg.innerHTML = `${grid}<path class="line-finish" d="${linePath(finish)}" />${
      bounce ? `<path class="line-bounce" d="${linePath(bounce)}" />` : ""
    }${labels}`;
    const legend = document.getElementById("rateLegend");
    if (legend) {
      legend.innerHTML = bounce
        ? `<span><i style="background:#1f7a45"></i>完播率</span><span><i style="background:#c56a1a;height:2px;border-top:2px dashed #c56a1a;background:transparent"></i>2s跳出率</span>`
        : `<span><i style="background:#1f7a45"></i>完播率</span>`;
    }
  }

  let sortKey = "views";

  function renderDailyDashboard(state) {
    const root = document.getElementById("dashboard");
    if (!root) return;
    const days = [...(state.posts || [])].sort((a, b) =>
      String(a.date).localeCompare(String(b.date))
    );
    const isDy = state.platform === "jiangwei_douyin";
    if (!days.length) {
      root.innerHTML = `<div class="panel"><p style="color:var(--muted)">当前没有可展示的日数据。</p></div>`;
      return;
    }

    const views = days.reduce((s, d) => s + (d.views || 0), 0);
    const likes = days.reduce((s, d) => s + (d.likes || 0), 0);
    const shares = days.reduce((s, d) => s + (d.shares || 0), 0);
    const comments = days.reduce((s, d) => s + (d.comments || 0), 0);
    const posts = days.reduce((s, d) => s + (d.posts || 0), 0);
    const follows = days.reduce((s, d) => s + (d.follows || 0), 0);
    const recommend = days.reduce((s, d) => s + (d.recommend || 0), 0);
    const peak = [...days].sort((a, b) => b.views - a.views)[0];
    const eng =
      ((likes + shares + comments + (isDy ? 0 : follows)) / Math.max(views, 1)) * 1000;
    const avgWatch = isDy
      ? days.reduce((s, d) => s + (d.avgWatchSec || 0), 0) / days.length
      : 0;
    const avgFinish5 = isDy
      ? days.reduce((s, d) => s + (d.finish5s || 0), 0) / days.length
      : 0;

    root.innerHTML = `
      <div class="stats reveal">
        <div class="stat"><div class="val accent">${fmt(views)}</div><div class="lbl">总播放量</div></div>
        <div class="stat"><div class="val">${fmt(likes)}</div><div class="lbl">${isDy ? "总点赞" : "总喜欢"}</div></div>
        <div class="stat"><div class="val">${fmt(shares)}</div><div class="lbl">总分享</div></div>
        <div class="stat"><div class="val">${fmt(comments)}</div><div class="lbl">总评论</div></div>
        <div class="stat"><div class="val up">${isDy ? posts : follows}</div><div class="lbl">${isDy ? "投稿量合计" : "关注合计"}</div></div>
      </div>
      <div class="stats-secondary reveal">
        <div class="stat"><div class="val">${days.length}</div><div class="lbl">统计天数</div></div>
        <div class="stat"><div class="val">${eng.toFixed(1)}</div><div class="lbl">互动 / 千次播放</div></div>
        <div class="stat"><div class="val">${isDy ? pct(avgFinish5) : fmt(recommend)}</div><div class="lbl">${isDy ? "日均5秒完播率" : "推荐合计"}</div></div>
        <div class="stat"><div class="val">${isDy ? avgWatch.toFixed(1) + "s" : fmt(Math.round(views / days.length))}</div><div class="lbl">${isDy ? "日均播放时长" : "日均播放"}</div></div>
      </div>
      <aside class="insight reveal">
        <div class="insight-mark">洞</div>
        <div>
          <h2>江炜个人号 · ${isDy ? "抖音" : "视频号"}日数据</h2>
          <p>
            ${days[0].date} 至 ${days[days.length - 1].date}，共 ${days.length} 天。
            峰值日 <strong>${peak.date}</strong> 播放 ${fmt(peak.views)}。
            ${isDy ? `期间投稿 ${posts} 条。` : `期间新增关注 ${fmt(follows)}，推荐 ${fmt(recommend)}。`}
          </p>
        </div>
      </aside>
      <section class="block reveal">
        <div class="section-head"><h2>每日播放趋势</h2><p>按日期 · 播放量</p></div>
        <div class="panel"><svg class="chart-svg" id="trendChart" viewBox="0 0 560 220"></svg></div>
      </section>
      <section class="block reveal">
        <div class="grid-2">
          <div>
            <div class="section-head"><h2>${isDy ? "完播 / 跳出 / 封面点击" : "互动趋势"}</h2><p>${isDy ? "百分比" : "喜欢/分享/评论/关注"}</p></div>
            <div class="panel">
              <svg class="chart-svg" id="rateChart" viewBox="0 0 400 220"></svg>
              <div class="legend" id="rateLegend"></div>
            </div>
          </div>
          <div>
            <div class="section-head"><h2>互动构成</h2><p>区间合计</p></div>
            <div class="panel">
              <div class="donut-wrap">
                <svg width="160" height="160" viewBox="0 0 160 160" id="engageDonut"></svg>
                <div class="donut-legend" id="engageLegend"></div>
              </div>
            </div>
          </div>
        </div>
      </section>
      <section class="block reveal">
        <div class="section-head"><h2>每日明细</h2><p>按时间正序</p></div>
        <div class="table-wrap"><table><thead id="tableHead"></thead><tbody id="tableBody"></tbody></table></div>
      </section>
    `;

    // reuse trend chart with fake posts shape
    renderTrend(days.map((d) => ({ ...d, datetime: d.date })));

    const svg = document.getElementById("rateChart");
    const legend = document.getElementById("rateLegend");
    if (svg) {
      const W = 400, H = 220, pad = { t: 16, r: 12, b: 36, l: 40 };
      const iw = W - pad.l - pad.r, ih = H - pad.t - pad.b;
      if (isDy) {
        const max = 80;
        const series = [
          { name: "5秒完播率", color: "#1f7a45", get: (d) => d.finish5s },
          { name: "2秒跳出率", color: "#c56a1a", get: (d) => d.bounce2s, dash: true },
          { name: "封面点击率", color: "#0f6b5c", get: (d) => d.coverCtr },
        ];
        let grid = "";
        for (let i = 0; i <= 4; i++) {
          const y = pad.t + (ih * i) / 4;
          const val = Math.round(max * (1 - i / 4));
          grid += `<line x1="${pad.l}" y1="${y}" x2="${W - pad.r}" y2="${y}" stroke="rgba(19,40,32,0.12)" /><text class="axis" x="${pad.l - 6}" y="${y + 4}" text-anchor="end">${val}%</text>`;
        }
        const paths = series
          .map((s) => {
            const pts = days.map((d, i) => ({
              x: pad.l + (i / Math.max(days.length - 1, 1)) * iw,
              y: pad.t + ih - (((s.get(d) || 0) * 100) / max) * ih,
            }));
            return `<path d="${linePath(pts)}" fill="none" stroke="${s.color}" stroke-width="2.2" stroke-linecap="round" ${s.dash ? 'stroke-dasharray="5 4"' : ""} />`;
          })
          .join("");
        const labels = days
          .map((d, i) =>
            i % 4 === 0 || i === days.length - 1
              ? `<text class="axis" x="${pad.l + (i / Math.max(days.length - 1, 1)) * iw}" y="${H - 12}" text-anchor="middle">${d.date.slice(5)}</text>`
              : ""
          )
          .join("");
        svg.innerHTML = grid + paths + labels;
        if (legend) {
          legend.innerHTML = series
            .map((s) => `<span><i style="background:${s.color}"></i>${s.name}</span>`)
            .join("");
        }
      } else {
        const maxV = Math.max(...days.map((d) => Math.max(d.likes, d.shares, d.comments, d.follows)), 1) * 1.2;
        const series = [
          { name: "喜欢", color: "#0f6b5c", get: (d) => d.likes },
          { name: "分享", color: "#3d9a7f", get: (d) => d.shares },
          { name: "评论", color: "#c56a1a", get: (d) => d.comments },
          { name: "关注", color: "#7aa82e", get: (d) => d.follows },
        ];
        let grid = "";
        for (let i = 0; i <= 4; i++) {
          const y = pad.t + (ih * i) / 4;
          const val = Math.round(maxV * (1 - i / 4));
          grid += `<line x1="${pad.l}" y1="${y}" x2="${W - pad.r}" y2="${y}" stroke="rgba(19,40,32,0.12)" /><text class="axis" x="${pad.l - 6}" y="${y + 4}" text-anchor="end">${fmt(val)}</text>`;
        }
        const paths = series
          .map((s) => {
            const pts = days.map((d, i) => ({
              x: pad.l + (i / Math.max(days.length - 1, 1)) * iw,
              y: pad.t + ih - ((s.get(d) || 0) / maxV) * ih,
            }));
            return `<path d="${linePath(pts)}" fill="none" stroke="${s.color}" stroke-width="2.1" stroke-linecap="round" />`;
          })
          .join("");
        const labels = days
          .map((d, i) =>
            i % 4 === 0 || i === days.length - 1
              ? `<text class="axis" x="${pad.l + (i / Math.max(days.length - 1, 1)) * iw}" y="${H - 12}" text-anchor="middle">${d.date.slice(5)}</text>`
              : ""
          )
          .join("");
        svg.innerHTML = grid + paths + labels;
        if (legend) {
          legend.innerHTML = series
            .map((s) => `<span><i style="background:${s.color}"></i>${s.name}</span>`)
            .join("");
        }
      }
    }

    if (isDy) {
      renderDonut(
        "engageDonut",
        "engageLegend",
        [
          { label: "点赞", value: likes },
          { label: "分享", value: shares },
          { label: "评论", value: comments },
        ],
        ["#0f6b5c", "#3d9a7f", "#c56a1a"]
      );
    } else {
      renderDonut(
        "engageDonut",
        "engageLegend",
        [
          { label: "喜欢", value: likes },
          { label: "分享", value: shares },
          { label: "评论", value: comments },
          { label: "关注", value: follows },
        ],
        ["#0f6b5c", "#3d9a7f", "#c56a1a", "#8a9a90"]
      );
    }

    const head = document.getElementById("tableHead");
    const body = document.getElementById("tableBody");
    if (isDy) {
      head.innerHTML = `<tr>
        <th>日期</th><th class="num">投稿</th><th class="num">播放</th><th class="num">点赞</th><th class="num">分享</th><th class="num">评论</th>
        <th class="num">5秒完播</th><th class="num">2秒跳出</th><th class="num">封面点击</th><th class="num">均播时长</th>
      </tr>`;
      body.innerHTML = days
        .map((d) => {
          const tone = d.views >= 5000 ? "hi" : d.views < 1000 ? "lo" : "";
          const dot = tone ? `<span class="dot-tone ${tone}"></span>` : "";
          return `<tr>
            <td>${dot}${d.date}</td>
            <td class="num">${d.posts}</td>
            <td class="num">${fmt(d.views)}</td>
            <td class="num">${d.likes}</td>
            <td class="num">${d.shares}</td>
            <td class="num">${d.comments}</td>
            <td class="num">${pct(d.finish5s)}</td>
            <td class="num">${pct(d.bounce2s)}</td>
            <td class="num">${pct(d.coverCtr)}</td>
            <td class="num">${(d.avgWatchSec || 0).toFixed(1)}s</td>
          </tr>`;
        })
        .join("");
    } else {
      head.innerHTML = `<tr>
        <th>日期</th><th class="num">播放</th><th class="num">推荐</th><th class="num">喜欢</th>
        <th class="num">评论</th><th class="num">分享</th><th class="num">关注</th>
      </tr>`;
      body.innerHTML = days
        .map((d) => {
          const tone = d.views >= 800 ? "hi" : d.views < 200 ? "lo" : "";
          const dot = tone ? `<span class="dot-tone ${tone}"></span>` : "";
          return `<tr>
            <td>${dot}${d.date}</td>
            <td class="num">${fmt(d.views)}</td>
            <td class="num">${d.recommend}</td>
            <td class="num">${d.likes}</td>
            <td class="num">${d.comments}</td>
            <td class="num">${d.shares}</td>
            <td class="num">${d.follows}</td>
          </tr>`;
        })
        .join("");
    }
  }

  function renderDashboard(state) {
    if (isDailyState(state)) {
      renderDailyDashboard(state);
      return;
    }
    const root = document.getElementById("dashboard");
    if (!root) return;
    const posts = state.posts || [];
    const platform = state.platform;
    const isDy = platform === "douyin";

    if (!posts.length) {
      root.innerHTML = `<div class="panel"><p style="color:var(--muted)">当前没有可展示的数据，请上传 Excel 或切换月份。</p></div>`;
      return;
    }

    const views = posts.reduce((s, p) => s + p.views, 0);
    const likes = posts.reduce((s, p) => s + p.likes, 0);
    const shares = posts.reduce((s, p) => s + p.shares, 0);
    const comments = posts.reduce((s, p) => s + p.comments, 0);
    const favorites = posts.reduce((s, p) => s + (p.favorites || 0), 0);
    const profile = posts.reduce((s, p) => s + (p.profileVisits || 0), 0);
    const fans = posts.reduce((s, p) => s + (p.followerDelta || 0), 0);
    const follows = posts.reduce((s, p) => s + (p.follows || 0), 0);
    const recommend = posts.reduce((s, p) => s + (p.recommend || 0), 0);
    const wfin = views
      ? posts.reduce((s, p) => s + p.views * p.finishRate, 0) / views
      : 0;
    const engVal = isDy
      ? ((likes + shares + comments + favorites) / Math.max(views, 1)) * 1000
      : ((likes + shares + comments + follows) / Math.max(views, 1)) * 1000;
    const byViews = sortPosts(posts, "views");
    const top = byViews[0];
    const topShare = views ? (top.views / views) * 100 : 0;
    const finishBest = [...posts].sort((a, b) => b.finishRate - a.finishRate)[0];

    root.innerHTML = `
      <div class="stats reveal">
        <div class="stat"><div class="val accent">${fmt(views)}</div><div class="lbl">总播放量</div></div>
        <div class="stat"><div class="val">${fmt(likes)}</div><div class="lbl">${isDy ? "总点赞" : "总喜欢"}</div></div>
        <div class="stat"><div class="val">${fmt(isDy ? profile : follows)}</div><div class="lbl">${isDy ? "主页访问" : "关注量"}</div></div>
        <div class="stat"><div class="val up">${isDy ? "+" + fans : fmt(recommend)}</div><div class="lbl">${isDy ? "粉丝净增" : "推荐次数"}</div></div>
        <div class="stat"><div class="val">${pct(wfin)}</div><div class="lbl">加权完播率</div></div>
      </div>
      <div class="stats-secondary reveal">
        <div class="stat"><div class="val">${fmt(shares)}</div><div class="lbl">总分享</div></div>
        <div class="stat"><div class="val">${fmt(comments)}</div><div class="lbl">总评论</div></div>
        <div class="stat"><div class="val">${fmt(isDy ? favorites : posts.reduce((s,p)=>s+(p.forwardChat||0),0))}</div><div class="lbl">${isDy ? "总收藏" : "转发聊天/朋友圈"}</div></div>
        <div class="stat"><div class="val">${engVal.toFixed(1)}</div><div class="lbl">互动 / 千次播放</div></div>
      </div>
      <aside class="insight reveal">
        <div class="insight-mark">洞</div>
        <div>
          <h2>数据洞察</h2>
          <p>
            共 <strong>${posts.length}</strong> 条作品；最高播放「${top.title}」${fmt(top.views)} 次，约占总播放
            <strong>${topShare.toFixed(1)}%</strong>。完播最好：「${finishBest.title}」${pct(finishBest.finishRate)}。
            ${isDy ? `粉丝净增 +${fans}，主页访问 ${fmt(profile)}。` : `新增关注 ${fmt(follows)}，推荐 ${fmt(recommend)} 次。`}
          </p>
        </div>
      </aside>
      <section class="block reveal">
        <div class="section-head"><h2>播放量排行</h2><p>按作品 · 次</p></div>
        <div class="panel"><div class="bars" id="viewBars"></div></div>
      </section>
      <section class="block reveal">
        <div class="grid-2">
          <div>
            <div class="section-head"><h2>发布节奏与播放趋势</h2><p>按时间 · 播放量</p></div>
            <div class="panel"><svg class="chart-svg" id="trendChart" viewBox="0 0 560 220"></svg></div>
          </div>
          <div>
            <div class="section-head"><h2>${isDy ? "完播率 vs 2s 跳出" : "完播率趋势"}</h2><p>百分比</p></div>
            <div class="panel">
              <svg class="chart-svg" id="rateChart" viewBox="0 0 400 220"></svg>
              <div class="legend" id="rateLegend"></div>
            </div>
          </div>
        </div>
      </section>
      <section class="block reveal">
        <div class="grid-2b">
          <div>
            <div class="section-head"><h2>互动构成</h2><p>${isDy ? "赞/转/评/藏" : "喜欢/分享/评论/关注"}</p></div>
            <div class="panel">
              <div class="donut-wrap">
                <svg width="160" height="160" viewBox="0 0 160 160" id="engageDonut"></svg>
                <div class="donut-legend" id="engageLegend"></div>
              </div>
            </div>
          </div>
          <div>
            <div class="section-head"><h2>${isDy ? "体裁分布" : "传播动作"}</h2><p>${isDy ? "作品数量" : "转发/铃声/状态等"}</p></div>
            <div class="panel">
              <div class="donut-wrap">
                <svg width="160" height="160" viewBox="0 0 160 160" id="secondDonut"></svg>
                <div class="donut-legend" id="secondLegend"></div>
              </div>
            </div>
          </div>
        </div>
      </section>
      <section class="block reveal">
        <div class="section-head">
          <h2>作品明细</h2>
          <div class="filters" id="sortFilters"></div>
        </div>
        <div class="table-wrap"><table><thead id="tableHead"></thead><tbody id="tableBody"></tbody></table></div>
      </section>
    `;

    const max = byViews[0].views || 1;
    document.getElementById("viewBars").innerHTML = byViews
      .slice(0, 15)
      .map(
        (p, i) => `
      <div class="bar-row" title="${p.fullTitle.replace(/"/g, "&quot;")}">
        <div class="bar-label">${p.title}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${(p.views / max) * 100}%;animation-delay:${i * 0.04}s"></div></div>
        <div class="bar-val">${fmt(p.views)}</div>
      </div>`
      )
      .join("");

    renderTrend(posts);
    renderRates(posts);

    if (isDy) {
      renderDonut(
        "engageDonut",
        "engageLegend",
        [
          { label: "点赞", value: likes },
          { label: "分享", value: shares },
          { label: "评论", value: comments },
          { label: "收藏", value: favorites },
        ],
        ["#0f6b5c", "#3d9a7f", "#c56a1a", "#8a9a90"]
      );
      const genres = {};
      posts.forEach((p) => {
        genres[p.genre || "未标注"] = (genres[p.genre || "未标注"] || 0) + 1;
      });
      renderDonut(
        "secondDonut",
        "secondLegend",
        Object.entries(genres).map(([label, value]) => ({ label, value })),
        ["#0a4a3f", "#b6e05a", "#6b7f74", "#c56a1a"]
      );
    } else {
      renderDonut(
        "engageDonut",
        "engageLegend",
        [
          { label: "喜欢", value: likes },
          { label: "分享", value: shares },
          { label: "评论", value: comments },
          { label: "关注", value: follows },
        ],
        ["#0f6b5c", "#3d9a7f", "#c56a1a", "#8a9a90"]
      );
      renderDonut(
        "secondDonut",
        "secondLegend",
        [
          { label: "转发聊天/朋友圈", value: posts.reduce((s, p) => s + (p.forwardChat || 0), 0) },
          { label: "设为铃声", value: posts.reduce((s, p) => s + (p.setRingtone || 0), 0) },
          { label: "设为状态", value: posts.reduce((s, p) => s + (p.setStatus || 0), 0) },
          { label: "朋友圈封面", value: posts.reduce((s, p) => s + (p.setMomentsCover || 0), 0) },
        ],
        ["#0a4a3f", "#b6e05a", "#6b7f74", "#c56a1a"]
      );
    }

    const opts = isDy
      ? [
          ["views", "按播放"],
          ["engage", "按互动率"],
          ["finish", "按完播"],
          ["profile", "按主页"],
          ["fans", "按涨粉"],
          ["date", "按时间"],
        ]
      : [
          ["views", "按播放"],
          ["engage", "按互动率"],
          ["finish", "按完播"],
          ["fans", "按关注"],
          ["date", "按时间"],
        ];

    function paintTable() {
      const rows = sortPosts(posts, sortKey);
      const head = document.getElementById("tableHead");
      const body = document.getElementById("tableBody");
      if (isDy) {
        head.innerHTML = `<tr>
          <th>作品</th><th>日期</th><th>体裁</th>
          <th class="num">播放</th><th class="num">完播率</th><th class="num">5s完播</th><th class="num">2s跳出</th>
          <th class="num">均播时长</th><th class="num">点赞</th><th class="num">分享</th><th class="num">评论</th>
          <th class="num">收藏</th><th class="num">主页</th><th class="num">涨粉</th><th class="num">互动/千次</th>
        </tr>`;
        body.innerHTML = rows
          .map((p) => {
            const tone = p.views >= 8000 ? "hi" : p.views < 1200 ? "lo" : "";
            const dot = tone ? `<span class="dot-tone ${tone}"></span>` : "";
            return `<tr>
              <td class="title" title="${p.fullTitle.replace(/"/g, "&quot;")}">${dot}${p.title}</td>
              <td>${p.date}</td><td>${p.genre}</td>
              <td class="num">${fmt(p.views)}</td>
              <td class="num">${pct(p.finishRate)}</td>
              <td class="num">${pct(p.finish5s)}</td>
              <td class="num">${pct(p.bounce2s)}</td>
              <td class="num">${(p.avgWatchSec || 0).toFixed(1)}s</td>
              <td class="num">${p.likes}</td><td class="num">${p.shares}</td><td class="num">${p.comments}</td>
              <td class="num">${p.favorites}</td><td class="num">${p.profileVisits}</td>
              <td class="num">${p.followerDelta > 0 ? "+" + p.followerDelta : p.followerDelta}</td>
              <td class="num">${engage(p).toFixed(1)}</td>
            </tr>`;
          })
          .join("");
      } else {
        head.innerHTML = `<tr>
          <th>视频</th><th>日期</th>
          <th class="num">播放</th><th class="num">完播率</th><th class="num">均播时长</th>
          <th class="num">喜欢</th><th class="num">分享</th><th class="num">评论</th><th class="num">关注</th>
          <th class="num">推荐</th><th class="num">转发</th><th class="num">企微点击</th><th class="num">互动/千次</th>
        </tr>`;
        body.innerHTML = rows
          .map((p) => {
            const tone = p.views >= 5000 ? "hi" : p.views < 800 ? "lo" : "";
            const dot = tone ? `<span class="dot-tone ${tone}"></span>` : "";
            return `<tr>
              <td class="title" title="${p.fullTitle.replace(/"/g, "&quot;")}">${dot}${p.title}</td>
              <td>${p.date}</td>
              <td class="num">${fmt(p.views)}</td>
              <td class="num">${pct(p.finishRate)}</td>
              <td class="num">${(p.avgWatchSec || 0).toFixed(1)}s</td>
              <td class="num">${p.likes}</td><td class="num">${p.shares}</td><td class="num">${p.comments}</td>
              <td class="num">${p.follows}</td><td class="num">${p.recommend}</td>
              <td class="num">${p.forwardChat}</td><td class="num">${p.wecomClicks}</td>
              <td class="num">${engage(p).toFixed(1)}</td>
            </tr>`;
          })
          .join("");
      }
    }

    const filters = document.getElementById("sortFilters");
    filters.innerHTML = opts
      .map(
        ([k, label]) =>
          `<button class="filter-btn${sortKey === k ? " active" : ""}" data-sort="${k}">${label}</button>`
      )
      .join("");
    filters.querySelectorAll("button").forEach((btn) => {
      btn.addEventListener("click", () => {
        sortKey = btn.dataset.sort;
        renderDashboard(state);
      });
    });
    paintTable();
  }

  const STORAGE_KEY = "sunevo-dash-store-v2";

  function sheetSortKey(name) {
    const m = String(name).match(/(\d{1,2})\s*月/);
    if (m) return Number(m[1]);
    const y = String(name).match(/(\d{4}).*?(\d{1,2})/);
    if (y) return Number(y[1]) * 100 + Number(y[2]);
    return 999;
  }

  function sortSheetNames(names) {
    return [...names].sort((a, b) => sheetSortKey(a) - sheetSortKey(b) || a.localeCompare(b, "zh"));
  }

  function inferMonthSheetName(sheetName, posts) {
    const name = String(sheetName || "").trim();
    if (/\d{1,2}\s*月/.test(name)) {
      const m = name.match(/(\d{1,2})\s*月/);
      return `${Number(m[1])}月数据`;
    }
    const counts = {};
    (posts || []).forEach((p) => {
      const hit = String(p.datetime || p.date || "").match(/(?:^|\D)(\d{4})-(\d{2})/);
      if (!hit) return;
      const key = `${Number(hit[2])}月数据`;
      counts[key] = (counts[key] || 0) + 1;
    });
    const best = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
    return best ? best[0] : name || "未命名数据";
  }

  /** Split posts into month buckets like "9月数据" using publish datetime. */
  function regroupByMonth(posts) {
    const out = {};
    (posts || []).forEach((p) => {
      const hit = String(p.datetime || p.date || "").match(/(?:^|\D)(\d{4})-(\d{2})/);
      const key = hit ? `${Number(hit[2])}月数据` : "未标注月份";
      if (!out[key]) out[key] = [];
      out[key].push(p);
    });
    return out;
  }

  function normalizeUploadedSheets(sheets) {
    const normalized = {};
    Object.entries(sheets || {}).forEach(([name, posts]) => {
      // If one sheet spans multiple months, split; else keep as one month bucket.
      const byMonth = regroupByMonth(posts);
      const monthKeys = Object.keys(byMonth).filter((k) => k !== "未标注月份");
      if (monthKeys.length >= 2) {
        monthKeys.forEach((k) => {
          normalized[k] = (normalized[k] || []).concat(byMonth[k]);
        });
        if (byMonth["未标注月份"] && byMonth["未标注月份"].length) {
          const fallback = inferMonthSheetName(name, byMonth["未标注月份"]);
          normalized[fallback] = (normalized[fallback] || []).concat(byMonth["未标注月份"]);
        }
      } else {
        const key = inferMonthSheetName(name, posts);
        normalized[key] = (normalized[key] || []).concat(posts);
      }
    });
    return normalized;
  }

  function readStore() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        return {
          douyin: null,
          shipinhao: null,
          jiangwei_douyin: null,
          jiangwei_shipinhao: null,
        };
      }
      const parsed = JSON.parse(raw);
      return {
        douyin: parsed.douyin || null,
        shipinhao: parsed.shipinhao || null,
        jiangwei_douyin: parsed.jiangwei_douyin || null,
        jiangwei_shipinhao: parsed.jiangwei_shipinhao || null,
      };
    } catch (_) {
      return {
        douyin: null,
        shipinhao: null,
        jiangwei_douyin: null,
        jiangwei_shipinhao: null,
      };
    }
  }

  function writeStore(store) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
  }

  function getSavedPlatform(platform) {
    const store = readStore();
    return store[platform] || null;
  }

  function clearSavedPlatform(platform) {
    const store = readStore();
    store[platform] = null;
    writeStore(store);
  }

  function mergeSheetsWithBuiltin(platform, overlaySheets) {
    const builtin =
      (global.DEFAULT_DATA && global.DEFAULT_DATA[platform]) || {};
    const merged = structuredClone(builtin);
    Object.entries(overlaySheets || {}).forEach(([name, rows]) => {
      if (Array.isArray(rows) && rows.length > 0) merged[name] = rows;
    });
    return merged;
  }

  /**
   * Merge uploaded month sheets into existing saved/builtin sheets.
   * - New month names are appended
   * - Same month name is replaced with the newly uploaded data
   * - Builtin months are kept unless explicitly replaced by non-empty upload
   */
  function mergeAndSave(platform, incomingSheets, meta) {
    const saved = getSavedPlatform(platform);
    const baseSheets =
      saved && saved.sheets
        ? mergeSheetsWithBuiltin(platform, saved.sheets)
        : structuredClone((global.DEFAULT_DATA && global.DEFAULT_DATA[platform]) || {});
    const normalized = normalizeUploadedSheets(incomingSheets);
    const merged = { ...baseSheets };
    const added = [];
    const replaced = [];
    Object.entries(normalized).forEach(([name, posts]) => {
      if (!posts || !posts.length) return;
      if (merged[name]) replaced.push(name);
      else added.push(name);
      merged[name] = posts;
    });
    const payload = {
      sheets: merged,
      updatedAt: new Date().toISOString(),
      lastFile: (meta && meta.fileName) || "",
      months: sortSheetNames(Object.keys(merged)),
    };
    const store = readStore();
    store[platform] = payload;
    writeStore(store);
    return { ...payload, added, replaced };
  }

  function resolveAccountSheets(platform) {
    const builtin =
      (global.DEFAULT_DATA && global.DEFAULT_DATA[platform]) || {};
    const saved = getSavedPlatform(platform);
    if (!saved || !saved.sheets) {
      return {
        sheets: structuredClone(builtin),
        source: "builtin",
        saved: null,
      };
    }
    const merged = mergeSheetsWithBuiltin(platform, saved.sheets);
    const hasOverlay = Object.entries(saved.sheets).some(
      ([, rows]) => Array.isArray(rows) && rows.length > 0
    );
    return {
      sheets: merged,
      source: hasOverlay ? "saved" : "builtin",
      saved: hasOverlay ? saved : null,
    };
  }

  function exportStoreJson(platform) {
    const saved = getSavedPlatform(platform);
    const builtin = (global.DEFAULT_DATA && global.DEFAULT_DATA[platform]) || {};
    const data = saved && saved.sheets ? saved : { sheets: builtin, updatedAt: null, lastFile: "" };
    const blob = new Blob([JSON.stringify(data, null, 2)], {
      type: "application/json;charset=utf-8",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `sunevo-${platform}-data.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  global.SunevoDash = {
    detectFormat,
    parseDouyinRows,
    parseShipinhaoRows,
    parseJiangweiDouyinDaily,
    parseJiangweiShipinhaoDaily,
    parseWorkbook,
    sortPosts,
    renderDashboard,
    isDailyState,
    fmt,
    pct,
    sortSheetNames,
    inferMonthSheetName,
    normalizeUploadedSheets,
    getSavedPlatform,
    clearSavedPlatform,
    mergeAndSave,
    exportStoreJson,
    readStore,
    resolveAccountSheets,
    mergeSheetsWithBuiltin,
  };
})(window);
