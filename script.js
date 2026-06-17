const fallbackEvents = [
  {
    id: "sample",
    name: "ตัวอย่างกิจกรรม",
    shortName: "ตัวอย่าง",
    icon: "broadcast",
    cycle: "กิจกรรมที่ 1",
    period: "-",
    resetDate: "ตัดรอบ/จัดส่ง: ทุกวันพุธ",
    latest: "รออัปเดต",
    status: "กำลังดำเนินการ",
    reward: "รางวัลกิจกรรม",
    winners: [],
    pendingRewards: [],
  },
];

// Visible now = not manually hidden, AND within the scheduled time window (if set)
function isVisibleNow(item) {
  if (!item) return false;
  if (item.visibility === "private" || item.visibility === "archived") return false;
  const now = Date.now();
  if (item.publishAt) {
    const t = new Date(item.publishAt).getTime();
    if (!isNaN(t) && now < t) return false; // not yet time to show
  }
  if (item.hideAt) {
    const t = new Date(item.hideAt).getTime();
    if (!isNaN(t) && now >= t) return false; // past hide time
  }
  return true;
}
function isPublicEvent(e) { return isVisibleNow(e); }

let _allEventsRaw = Array.isArray(window.WARZ_EVENTS) && window.WARZ_EVENTS.length ? window.WARZ_EVENTS : fallbackEvents;
let events = _allEventsRaw.filter(isVisibleNow);
if (!events.length) events = fallbackEvents;
let masterCodes = Array.isArray(window.WARZ_MASTER_CODES) ? window.WARZ_MASTER_CODES : [];

let activeEvent = events[0];
let searchMode = "uid";
let statusFilter = "all";
let rewardFilter = "all";
let currentPage = 1;
const PAGE_SIZE = 100;
let sortKey = null;     // 'facebook' | 'reward' | 'status' | 'updated'
let sortDir = 1;        // 1 asc, -1 desc

const tabs = document.querySelector("#event-tabs");
const uidLookupInput = document.querySelector("#uid-lookup-input");
const uidLookupResult = document.querySelector("#uid-lookup-result");
const winnerRows = document.querySelector("#winner-rows");
const searchInput = document.querySelector("#winner-search");
const summary = document.querySelector("#event-summary");
const activityCards = document.querySelector("#activity-cards");
const codeList = document.querySelector("#code-list");
const toast = document.querySelector("#toast");
const header = document.querySelector(".site-header");
const navToggle = document.querySelector(".nav-toggle");
const searchResult = document.querySelector("#search-result");
const currentEventName = document.querySelector("#current-event-name");
const currentEventMeta = document.querySelector("#current-event-meta");
const clearSearch = document.querySelector("#clear-search");
const rewardAnnouncement = document.querySelector("#reward-announcement");
const eventCoverBanner = document.querySelector("#event-cover-banner");
const eventExtra = document.querySelector("#event-extra");
const deliveryProgress = document.querySelector("#delivery-progress");
const pageSections = document.querySelectorAll(".page-section");
const pageLinks = document.querySelectorAll(".page-link");
const searchModeTabs = document.querySelector("#search-mode-tabs");
const statusFilterSelect = document.querySelector("#status-filter");
const rewardFilterSelect = document.querySelector("#reward-filter");
const winnerCountLabel = document.querySelector("#winner-count-label");

function normalize(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function statusClass(status) {
  const text = normalize(status);
  if (text.includes("จัดส่งแล้ว") || text.includes("รับรางวัลแล้ว")) return "done";
  if (text.includes("พร้อมใช้")) return "ready";
  if (text.includes("ดำเนินการ") || text.includes("รอกดรับ") || text.includes("ติดต่อแก้ข้อมูล")) return "pending";
  if (text.includes("หมดเขต") || text.includes("ใช้แล้ว")) return "expired";
  return "checking";
}

function rewardName(reward) {
  if (typeof reward === "string") return reward;
  if (reward.hasItem && reward.itemEn) return reward.itemEn;
  return reward.forumReward || reward.name || "รางวัลกิจกรรม";
}

function rewardCategory(reward) {
  if (typeof reward === "string") return reward;
  return reward.forumReward || reward.name || "รางวัลกิจกรรม";
}

function itemName(reward) {
  if (typeof reward === "string") return reward;
  return reward.itemEn || reward.name || reward.forumReward || "รางวัลกิจกรรม";
}

function winnerStatus(winner) {
  return winner.claimStatus || activeEvent.status || "กำลังดำเนินการ";
}

const REWARD_CATEGORIES = [
  "Luckydraw", "ได้ทุกคน", "ถูกใจทีมงาน",
  "อันดับ 1", "อันดับ 2", "อันดับ 3", "อันดับที่ 4 - 10", "อื่นๆ",
];

function normalizeRewardCategory(str) {
  const s = String(str || "").toLowerCase();
  if (!s) return "อื่นๆ";
  if (REWARD_CATEGORIES.includes(str)) return str;
  if (s.includes("lucky")) return "Luckydraw";
  if (s.includes("ทุกคน")) return "ได้ทุกคน";
  if (s.includes("ถูกใจ")) return "ถูกใจทีมงาน";
  if (s.includes("อันดับ")) {
    if (/(4|5|6|7|8|9|10|4\s*-\s*10)/.test(s)) return "อันดับที่ 4 - 10";
    if (s.includes("1")) return "อันดับ 1";
    if (s.includes("2")) return "อันดับ 2";
    if (s.includes("3")) return "อันดับ 3";
  }
  return "อื่นๆ";
}

function winnerCategory(winner) {
  if (winner.rewardCategory) return winner.rewardCategory;
  const fr = (winner.rewards || []).map((r) => (typeof r === "string" ? r : r.forumReward)).find(Boolean);
  return normalizeRewardCategory(fr);
}

function winnerRewardType(winner) {
  return winnerCategory(winner);
}

function winnerRewardItems(winner) {
  return [winnerCategory(winner)];
}

function collectEventRewards(event) {
  const map = new Map();

  // If the event has its own reward list (set in admin), it is authoritative.
  // Otherwise derive from the winners' item rewards (legacy / imported data).
  const source = (event.pendingRewards && event.pendingRewards.length)
    ? event.pendingRewards
    : event.winners.flatMap((winner) => winner.rewards || []);

  source.forEach((reward) => {
    if (!reward) return;
    if (typeof reward !== "string" && (!reward.hasItem || !reward.itemEn)) return;
    const key = typeof reward === "string" ? reward : reward.itemId || reward.name;
    if (!map.has(key)) {
      map.set(key, {
        itemEn: itemName(reward),
        amount: typeof reward === "string" ? "1" : reward.amount || "1",
        imageUrl: typeof reward === "string" ? "" : reward.imageUrl || "",
      });
    }
  });

  return [...map.values()].filter((reward) => reward.itemEn).slice(0, 12);
}

function getFilteredWinners() {
  return activeEvent.winners.filter((winner) => {
    const matchesStatus = statusFilter === "all" || winnerStatus(winner) === statusFilter;
    const matchesReward = rewardFilter === "all" || winnerRewardType(winner) === rewardFilter;
    return matchesStatus && matchesReward;
  });
}

function isGuildEvent(event) {
  return event && event.eventType === "guild";
}

function getSearchValueByMode(winner) {
  if (winner.guild) return winner.guild; // guild events: match by guild name
  return searchMode === "facebook" ? winner.facebook : winner.uid || winner.character;
}

function getSearchPlaceholder() {
  if (isGuildEvent(activeEvent)) return "ค้นหาชื่อกิลด์";
  return searchMode === "facebook" ? "ค้นหาชื่อ Facebook" : "ค้นหา UID";
}

function getSearchExamples() {
  if (isGuildEvent(activeEvent)) return "ตัวอย่าง: ชื่อกิลด์ของคุณ";
  return searchMode === "facebook" ? "ตัวอย่าง: Sarada Ken, Jui Na Ja" : "ตัวอย่าง: 024VHD, HFL131, XB9Y20";
}

// Prefix match: the value (or any word in it) must START with the keyword.
// e.g. "7" → only UIDs starting with 7; "dang" → name with a word starting with "dang"
function matchesPrefix(value, keyword) {
  const v = normalize(value);
  if (!keyword) return false;
  if (v.startsWith(keyword)) return true;
  return v.split(/\s+/).some((word) => word.startsWith(keyword));
}

function getSearchMatches(keyword) {
  return events
    .flatMap((event) =>
      event.winners
        .filter((winner) => matchesPrefix(getSearchValueByMode(winner), keyword))
        .map((winner) => ({ event, winner }))
    )
    .slice(0, 30);
}

function renderTabs() {
  tabs.innerHTML = events
    .map(
      (event, index) => `
        <button class="event-tab ${event.id === activeEvent.id ? "is-active" : ""}" data-event="${event.id}" type="button">
          <span class="tab-avatar">${event.coverImage
            ? `<img src="${escapeHtml(event.coverImage)}" alt="" loading="lazy" onerror="this.parentNode.textContent='${String(index + 1).padStart(2, "0")}'" />`
            : String(index + 1).padStart(2, "0")}</span>
          <span>
            <strong>${escapeHtml(event.shortName || event.name)}</strong>
            <small>${event.winners.length.toLocaleString("th-TH")} รายชื่อ</small>
          </span>
        </button>
      `
    )
    .join("");

  tabs.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", () => {
      activeEvent = events.find((event) => event.id === button.dataset.event) || events[0];
      statusFilter = "all";
      rewardFilter = "all";
      renderAll();
    });
  });
}

function renderSearchModeTabs() {
  if (!searchModeTabs) return;

  // Hide the UID/Facebook mode toggle for guild events (search is by guild name)
  searchModeTabs.style.display = isGuildEvent(activeEvent) ? "none" : "";

  searchModeTabs.querySelectorAll("button").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.searchMode === searchMode);
  });

  searchInput.placeholder = getSearchPlaceholder();
}

function renderSearchResult() {
  const keyword = normalize(searchInput.value);
  const guildMode = isGuildEvent(activeEvent);

  if (!keyword) {
    searchResult.innerHTML = `
      <div class="empty-result">
        <strong>${guildMode ? "ค้นหาด้วยชื่อกิลด์" : searchMode === "facebook" ? "ค้นหาด้วยชื่อ Facebook" : "ค้นหาด้วย UID"}</strong>
        <span>${getSearchExamples()}</span>
      </div>
    `;
    return;
  }

  const matches = getSearchMatches(keyword);

  if (!matches.length) {
    searchResult.innerHTML = `
      <div class="empty-result is-warning">
        <strong>ไม่พบข้อมูลรางวัล</strong>
        <span>ไม่พบชื่อ/UID นี้ในกิจกรรม — ถ้าคิดว่าผิดพลาด ติดต่อทีมงานได้เลย</span>
        <a class="support-inline-btn" href="https://liff.thehof.gg/th/warzth/ticket/home" target="_blank" rel="noopener noreferrer">💬 ติดต่อ Customer Support</a>
      </div>
    `;
    return;
  }

  searchResult.innerHTML = matches
    .map(
      ({ event, winner }) => `
        <article class="reward-result">
          <div class="result-player">
            ${winner.guild
              ? `<span class="uid">กิลด์</span>
                 <h3>${escapeHtml(winner.guild)}</h3>
                 <p>${escapeHtml(event.name)}</p>`
              : `<span class="uid">${searchMode === "facebook" ? `Facebook ${escapeHtml(winner.facebook || "-")}` : `UID ${escapeHtml(winner.uid || winner.character || "-")}`}</span>
                 <h3>${escapeHtml(searchMode === "facebook" ? winner.uid || winner.character || "-" : winner.facebook || "-")}</h3>
                 <p>${escapeHtml(event.name)} • ${escapeHtml(winner.claimMethod || "-")}</p>`}
          </div>
          <div class="reward-stack">
            <span class="reward-chip reward-chip-category"><b>${escapeHtml(winnerCategory(winner))}</b></span>
            ${(winner.rewards || [])
              .filter((reward) => reward && (reward.imageUrl || reward.itemEn))
              .map(
                (reward) => `
                  <span class="reward-chip">
                    ${reward.imageUrl ? `<img src="${escapeHtml(reward.imageUrl)}" alt="" loading="lazy" />` : ""}
                    <b>${escapeHtml(itemName(reward))}</b>
                  </span>
                `
              )
              .join("")}
          </div>
          <div class="result-status">
            <span class="status ${statusClass(winnerStatus(winner))}">${escapeHtml(winnerStatus(winner))}</span>
            <small>อัปเดต ${escapeHtml(winner.updatedAt || event.latest)}</small>
            <button type="button" class="share-result-btn" onclick="shareResult('${escapeHtml(winner.guild || winner.uid || winner.character || winner.facebook || "")}')">🔗 แชร์/คัดลอกลิงก์</button>
          </div>
        </article>
      `
    )
    .join("");
}

function shareResult(value) {
  if (!value) return;
  const url = `${location.origin}${location.pathname}?q=${encodeURIComponent(value)}`;
  if (navigator.share) {
    navigator.share({ title: "WARZ — ผลรางวัล", text: `ผลรางวัล WARZ ของ ${value}`, url }).catch(() => {});
  } else {
    navigator.clipboard.writeText(url).then(() => showToast("คัดลอกลิงก์แล้ว ✓")).catch(() => showToast(url));
  }
}

function renderFilterOptions() {
  const statuses = [...new Set(activeEvent.winners.map((winner) => winnerStatus(winner)).filter(Boolean))];
  const rewards = [...new Set(activeEvent.winners.map((winner) => winnerRewardType(winner)).filter(Boolean))];

  statusFilterSelect.innerHTML = ['<option value="all">ทุกสถานะ</option>']
    .concat(statuses.map((status) => `<option value="${escapeHtml(status)}">${escapeHtml(status)}</option>`))
    .join("");

  rewardFilterSelect.innerHTML = ['<option value="all">ทุกรางวัล</option>']
    .concat(rewards.map((reward) => `<option value="${escapeHtml(reward)}">${escapeHtml(reward)}</option>`))
    .join("");

  statusFilterSelect.value = statuses.includes(statusFilter) ? statusFilter : "all";
  rewardFilterSelect.value = rewards.includes(rewardFilter) ? rewardFilter : "all";
}

function renderHeroStats() {
  const el = document.getElementById("hero-stats");
  if (!el) return;
  const totalWinners = events.reduce((s, e) => s + e.winners.length, 0);
  const stats = [
    [events.length, "กิจกรรม"],
    [totalWinners.toLocaleString("th-TH"), "ผู้ได้รับรางวัล"],
    [masterCodes.length, "Master Code"],
  ];
  el.innerHTML = stats
    .map(([n, label]) => `<div class="hero-stat"><strong>${n}</strong><span>${label}</span></div>`)
    .join("");
}

function eventInitial(event) {
  const s = (event.shortName || event.name || "?").trim();
  return s ? s[0] : "?";
}

function renderHeroRecent() {
  const el = document.getElementById("hero-recent");
  if (!el) return;
  if (!events.length) { el.innerHTML = ""; return; }
  el.innerHTML = `
    <span class="hero-recent-label">กิจกรรมตอนนี้</span>
    <div class="hero-events-strip">
      ${events
        .map(
          (event) => `
        <button class="hero-event-chip" type="button" data-event="${escapeHtml(event.id)}" title="${escapeHtml(event.name)}">
          <span class="hero-event-avatar">
            ${event.coverImage
              ? `<img src="${escapeHtml(event.coverImage)}" alt="" loading="lazy" onerror="this.parentNode.textContent='${escapeHtml(eventInitial(event))}'" />`
              : escapeHtml(eventInitial(event))}
          </span>
          <span class="hero-event-name">${escapeHtml(event.shortName || event.name)}</span>
        </button>`
        )
        .join("")}
    </div>
  `;
  el.querySelectorAll(".hero-event-chip").forEach((btn) => {
    btn.addEventListener("click", () => {
      activeEvent = events.find((e) => e.id === btn.dataset.event) || events[0];
      statusFilter = "all";
      rewardFilter = "all";
      renderAll();
      showPage("rewards");
    });
  });
}

function exportWinnersCsv() {
  const rows = getSortedWinners();
  if (!rows.length) { showToast("ไม่มีรายชื่อให้ดาวน์โหลด"); return; }
  const header = ["ลำดับ", "Facebook", "UID", "รางวัล", "สถานะ", "อัปเดต"];
  const csvRows = rows.map((w, i) => [
    i + 1,
    w.facebook || "",
    w.uid || "",
    winnerRewardItems(w).join(" / "),
    winnerStatus(w),
    w.updatedAt || activeEvent.latest || "",
  ]);
  const esc = (v) => {
    const s = String(v ?? "");
    return /[",\n]/.test(s) ? `"${s.replaceAll('"', '""')}"` : s;
  };
  const csv = "﻿" + [header, ...csvRows].map((r) => r.map(esc).join(",")).join("\r\n");
  const safeName = (activeEvent.name || "winners").replace(/[\\/:*?"<>|]/g, "_").slice(0, 40);
  const a = Object.assign(document.createElement("a"), {
    href: URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" })),
    download: `${safeName}.csv`,
  });
  a.click();
  showToast(`ดาวน์โหลด ${rows.length} รายชื่อแล้ว`);
}

function winnerSortValue(winner, key) {
  switch (key) {
    case "facebook": return normalize(winner.guild || winner.facebook || winner.uid || "");
    case "reward":   return normalize(winnerRewardItems(winner).join(", "));
    case "status":   return normalize(winnerStatus(winner));
    case "updated":  return normalize(winner.updatedAt || activeEvent.latest || "");
    default:         return "";
  }
}

function getSortedWinners() {
  const base = [...getFilteredWinners()].sort((a, b) => (a.note ? 1 : 0) - (b.note ? 1 : 0));
  if (!sortKey) return base;
  return base.sort((a, b) => {
    const va = winnerSortValue(a, sortKey);
    const vb = winnerSortValue(b, sortKey);
    return va.localeCompare(vb, "th") * sortDir;
  });
}

function renderSortIndicators() {
  document.querySelectorAll(".sortable-table th.sortable").forEach((th) => {
    const ind = th.querySelector(".sort-ind");
    if (!ind) return;
    ind.textContent = th.dataset.sort === sortKey ? (sortDir === 1 ? "▲" : "▼") : "";
    th.classList.toggle("is-sorted", th.dataset.sort === sortKey);
  });
}

function renderWinnerIdHeader() {
  const th = document.querySelector('.sortable-table th[data-sort="facebook"]');
  if (!th) return;
  const label = isGuildEvent(activeEvent) ? "กิลด์" : "Facebook / UID";
  th.innerHTML = `${label} <span class="sort-ind"></span>`;
}

function renderWinners() {
  const guild = isGuildEvent(activeEvent);
  const sorted = getSortedWinners();
  const totalPages = Math.max(1, Math.ceil(sorted.length / PAGE_SIZE));
  if (currentPage > totalPages) currentPage = totalPages;

  const start = (currentPage - 1) * PAGE_SIZE;
  const pageRows = sorted.slice(start, start + PAGE_SIZE);

  winnerCountLabel.textContent = `${sorted.length.toLocaleString("th-TH")} ${guild ? "กิลด์" : "รายชื่อ"}`;
  renderWinnerIdHeader();

  winnerRows.innerHTML = pageRows.length
    ? pageRows
        .map(
          (winner, index) => `
            <tr${winner.note ? ' class="row-problem"' : ""}>
              <td class="rank">${start + index + 1}</td>
              <td>
                ${guild
                  ? `<strong>${escapeHtml(winner.guild || "-")}</strong>`
                  : `<strong>${escapeHtml(winner.facebook || "-")}</strong>
                     <small>UID ${escapeHtml(winner.uid || "-")} / ${escapeHtml(winner.claimMethod || "-")}</small>`}
              </td>
              <td>${escapeHtml(winnerRewardItems(winner).join(", "))}</td>
              <td>
                <span class="status ${statusClass(winnerStatus(winner))}">${escapeHtml(winnerStatus(winner))}</span>
                ${winner.note ? renderWinnerNote(winner.note) : ""}
              </td>
              <td>${escapeHtml(winner.updatedAt || activeEvent.latest)}</td>
            </tr>
          `
        )
        .join("")
    : '<tr><td colspan="5">ไม่พบรายชื่อที่ตรงกับตัวกรองที่เลือก</td></tr>';

  renderPagination(totalPages);
  renderSortIndicators();
}

const TICKET_URL = "https://liff.thehof.gg/th/warzth/ticket/home";

function renderWinnerNote(note) {
  // If the note tells the user to contact CS / open a ticket, make it a clickable link
  if (/ติดต่อ\s*cs|ticket|ติดต่อทีมงาน/i.test(note)) {
    return `<a class="problem-note problem-note-link" href="${TICKET_URL}" target="_blank" rel="noopener noreferrer">⚠ ${escapeHtml(note)} →</a>`;
  }
  return `<small class="problem-note">${escapeHtml(note)}</small>`;
}

function sortWinnersBy(key) {
  if (sortKey === key) sortDir = -sortDir;
  else { sortKey = key; sortDir = 1; }
  currentPage = 1;
  renderWinners();
}

function renderPagination(totalPages) {
  const paginationEl = document.getElementById("winners-pagination");
  if (!paginationEl) return;
  if (totalPages <= 1) { paginationEl.innerHTML = ""; return; }

  const maxButtons = 7;
  let pages = [];

  if (totalPages <= maxButtons) {
    pages = Array.from({ length: totalPages }, (_, i) => i + 1);
  } else {
    pages = [1];
    let start = Math.max(2, currentPage - 2);
    let end   = Math.min(totalPages - 1, currentPage + 2);
    if (start > 2)            pages.push("…");
    for (let i = start; i <= end; i++) pages.push(i);
    if (end < totalPages - 1) pages.push("…");
    pages.push(totalPages);
  }

  paginationEl.innerHTML = `
    <button class="page-btn page-prev" ${currentPage === 1 ? "disabled" : ""} onclick="goToPage(${currentPage - 1})">‹ ก่อนหน้า</button>
    <div class="page-numbers">
      ${pages.map(p =>
        p === "…"
          ? `<span class="page-ellipsis">…</span>`
          : `<button class="page-btn page-num ${p === currentPage ? "is-active" : ""}" onclick="goToPage(${p})">${p}</button>`
      ).join("")}
    </div>
    <button class="page-btn page-next" ${currentPage === totalPages ? "disabled" : ""} onclick="goToPage(${currentPage + 1})">ถัดไป ›</button>
  `;
}

function goToPage(page) {
  const sorted = getSortedWinners();
  const totalPages = Math.max(1, Math.ceil(sorted.length / PAGE_SIZE));
  currentPage = Math.max(1, Math.min(page, totalPages));
  renderWinners();
  document.querySelector(".winners-table-wrap")?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderSummary() {
  currentEventName.textContent = activeEvent.name;
  currentEventMeta.textContent = `${activeEvent.cycle} / ${activeEvent.period}`;

  const rows = [
    ["สถานะ", activeEvent.status],
    ["จำนวนรายชื่อ", `${activeEvent.winners.length.toLocaleString("th-TH")} คน`],
    ["ระยะเวลากดรับ", activeEvent.period],
    ["วันที่จัดส่ง", activeEvent.latest],
  ];

  summary.innerHTML = rows
    .map(
      ([label, value]) => `
        <div class="summary-row">
          <dt>${escapeHtml(label)}</dt>
          <dd>${escapeHtml(value)}</dd>
        </div>
      `
    )
    .join("");
}

let _deadlineTimer = null;

function deadlineText(dl) {
  const end = new Date(dl).getTime();
  if (isNaN(end)) return "";
  const diff = end - Date.now();
  if (diff <= 0) return `<span class="dl-ended">⛔ หมดเขตกดรับรางวัลแล้ว</span>`;
  const d = Math.floor(diff / 86400000);
  const h = Math.floor((diff % 86400000) / 3600000);
  const m = Math.floor((diff % 3600000) / 60000);
  const s = Math.floor((diff % 60000) / 1000);
  let t;
  if (d > 0) t = `${d} วัน ${h} ชม. ${m} นาที`;
  else if (h > 0) t = `${h} ชม. ${m} นาที ${s} วิ`;
  else t = `${m} นาที ${s} วิ`;
  const urgent = diff < 86400000 ? " dl-urgent" : "";
  return `<span class="dl-live${urgent}">⏰ หมดเขตกดรับใน ${t}</span>`;
}

function renderEventExtra() {
  if (!eventExtra) return;
  if (_deadlineTimer) { clearInterval(_deadlineTimer); _deadlineTimer = null; }
  const ev = activeEvent || {};
  const parts = [];
  if (ev.fbPostUrl) {
    parts.push(`<a class="event-fb-btn" href="${escapeHtml(ev.fbPostUrl)}" target="_blank" rel="noopener noreferrer">📘 ดูโพสต์กิจกรรมนี้</a>`);
  }
  if (ev.claimDeadline) {
    parts.push(`<span class="deadline-badge" id="deadline-badge"></span>`);
  }
  eventExtra.innerHTML = parts.join("");
  eventExtra.style.display = parts.length ? "flex" : "none";

  if (ev.claimDeadline) {
    const update = () => {
      const badge = document.getElementById("deadline-badge");
      if (!badge) { if (_deadlineTimer) clearInterval(_deadlineTimer); return; }
      badge.innerHTML = deadlineText(ev.claimDeadline);
    };
    update();
    _deadlineTimer = setInterval(update, 1000);
  }
}

function isDeliveredStatus(status) {
  const s = normalize(status);
  return s.includes("จัดส่งแล้ว") || s.includes("รับรางวัลแล้ว");
}

function renderDeliveryProgress() {
  if (!deliveryProgress) return;
  const winners = activeEvent.winners || [];
  const total = winners.length;
  if (!total) { deliveryProgress.innerHTML = ""; deliveryProgress.style.display = "none"; return; }
  const done = winners.filter((w) => isDeliveredStatus(winnerStatus(w))).length;
  const pct = Math.round((done / total) * 100);
  deliveryProgress.style.display = "block";
  deliveryProgress.innerHTML = `
    <div class="dp-head">
      <span class="dp-label">ความคืบหน้าการจัดส่ง</span>
      <span class="dp-count"><strong>${done.toLocaleString("th-TH")}</strong> / ${total.toLocaleString("th-TH")} (${pct}%)</span>
    </div>
    <div class="dp-bar"><div class="dp-fill" style="width:${pct}%"></div></div>
  `;
}

function renderEventCover() {
  if (!eventCoverBanner) return;
  const cover = activeEvent && activeEvent.coverImage;
  eventCoverBanner.innerHTML = cover
    ? `<img src="${escapeHtml(cover)}" alt="${escapeHtml(activeEvent.name)}" loading="lazy" />`
    : "";
  eventCoverBanner.style.display = cover ? "block" : "none";
}

function getEventRewardSets(event) {
  if (Array.isArray(event.rewardSets) && event.rewardSets.length) {
    return event.rewardSets
      .map((s) => ({
        category: s.category,
        items: (s.items || []).filter((r) => r && r.itemEn).map((r) => ({
          itemEn: itemName(r), amount: r.amount || "1", imageUrl: r.imageUrl || "",
        })),
      }))
      .filter((s) => s.items.length);
  }
  // Fallback: derive from winners grouped by each winner's reward category
  const byCat = new Map();
  for (const w of event.winners) {
    const cat = winnerCategory(w);
    const items = (w.rewards || []).filter((r) => r && r.hasItem && r.itemEn);
    if (!items.length) continue;
    if (!byCat.has(cat)) byCat.set(cat, new Map());
    const m = byCat.get(cat);
    for (const r of items) {
      const k = r.itemId || r.itemEn;
      if (!m.has(k)) m.set(k, { itemEn: itemName(r), amount: r.amount || "1", imageUrl: r.imageUrl || "" });
    }
  }
  return [...byCat.entries()].map(([category, m]) => ({ category, items: [...m.values()] }));
}

function renderRewardAnnouncement() {
  const sets = getEventRewardSets(activeEvent);

  if (!sets.length) {
    rewardAnnouncement.innerHTML = `
      <div class="side-empty">
        <strong>รอประกาศของรางวัล</strong>
        <span>เมื่อมีรายการรางวัลในกิจกรรม จะแสดงตรงนี้อัตโนมัติ</span>
      </div>`;
    return;
  }

  rewardAnnouncement.innerHTML = sets
    .map(
      (set) => `
        <div class="reward-group">
          <div class="reward-group-head">${escapeHtml(set.category)}</div>
          ${set.items
            .map(
              (reward) => `
            <article class="side-reward">
              <div class="side-reward-img">
                ${reward.imageUrl
                  ? `<img src="${escapeHtml(reward.imageUrl)}" alt="${escapeHtml(reward.itemEn)}" loading="lazy" />`
                  : "<span>WARZ</span>"}
              </div>
              <div>
                <strong>${escapeHtml(reward.itemEn)}</strong>
                <small>Amount: ${escapeHtml(reward.amount)}</small>
              </div>
            </article>`
            )
            .join("")}
        </div>`
    )
    .join("");
}

function renderUidEventCard(event, winner) {
  if (!winner) {
    return `
      <article class="uid-event-card not-found">
        <div class="uid-card-head">
          <span class="uid-event-num">${escapeHtml(event.cycle)}</span>
          <span class="uid-no-reward">ไม่ได้รับรางวัล</span>
        </div>
        <h3>${escapeHtml(event.name)}</h3>
        <p class="uid-period">${escapeHtml(event.period)}</p>
      </article>
    `;
  }

  const reward = (winner.rewards || [])[0];
  const category = winnerCategory(winner);
  const item = reward && reward.hasItem ? reward.itemEn : "";

  return `
    <article class="uid-event-card is-found${winner.note ? " has-problem" : ""}">
      <div class="uid-card-head">
        <span class="uid-event-num">${escapeHtml(event.cycle)}</span>
        <span class="status ${statusClass(winner.claimStatus || event.status)}">${escapeHtml(winner.claimStatus || event.status)}</span>
      </div>
      <h3>${escapeHtml(event.name)}</h3>
      <div class="uid-reward-row">
        ${reward && reward.imageUrl ? `<img src="${escapeHtml(reward.imageUrl)}" alt="${escapeHtml(item || category)}" />` : ""}
        <div class="uid-reward-text">
          <b>${escapeHtml(category)}</b>
          ${item ? `<small>${escapeHtml(item)}</small>` : ""}
        </div>
      </div>
      ${winner.note ? `<div class="uid-note">⚠ ${escapeHtml(winner.note)}</div>` : ""}
      <p class="uid-period">${escapeHtml(event.period)}</p>
    </article>
  `;
}

function renderUidLookupResult() {
  if (!uidLookupResult) return;
  const keyword = normalize(uidLookupInput.value);

  if (!keyword) {
    uidLookupResult.innerHTML = `
      <div class="empty-result">
        <strong>กรอก UID เพื่อเริ่มค้นหา</strong>
        <span>ตัวอย่าง: 024VHD, XB9Y20, 5VD4O6 (ตัวพิมพ์ใหญ่-เล็กก็ได้)</span>
      </div>
    `;
    return;
  }

  const results = events.map((event) => ({
    event,
    winner: event.winners.find((w) => matchesPrefix(w.uid || w.character, keyword)),
  }));

  const foundCount = results.filter((r) => r.winner).length;

  if (foundCount === 0) {
    uidLookupResult.innerHTML = `
      <div class="empty-result is-warning">
        <strong>ไม่พบ UID "${escapeHtml(uidLookupInput.value.trim())}" ในกิจกรรมใดเลย</strong>
        <span>ตรวจสอบ UID ให้ถูกต้อง เช่น 024VHD</span>
      </div>
    `;
    return;
  }

  const displayUid = results.find((r) => r.winner)?.winner?.uid || uidLookupInput.value.trim().toUpperCase();

  uidLookupResult.innerHTML = `
    <div class="uid-lookup-summary">
      <div>
        <strong class="uid">${escapeHtml(displayUid)}</strong>
        <span>${escapeHtml(results.find((r) => r.winner)?.winner?.facebook || "")}</span>
      </div>
      <span class="uid-summary-badge">${foundCount} / ${events.length} กิจกรรม</span>
    </div>
    <div class="uid-event-grid">
      ${results.map(({ event, winner }) => renderUidEventCard(event, winner)).join("")}
    </div>
  `;
}

function renderActivityCards() {
  activityCards.innerHTML = events
    .map(
      (event) => `
        <article class="activity-card compact-card${event.coverImage ? " has-cover" : ""}">
          ${event.coverImage ? `<div class="activity-cover"><img src="${escapeHtml(event.coverImage)}" alt="${escapeHtml(event.name)}" loading="lazy" /></div>` : ""}
          <div>
            <h3>${escapeHtml(event.name)}</h3>
            <p>${escapeHtml(event.period)}</p>
            <strong>${event.winners.length.toLocaleString("th-TH")} รายชื่อ / ${escapeHtml(event.status)}</strong>
            <button class="open-event-btn" type="button" data-event="${event.id}">เปิดกิจกรรมนี้</button>
          </div>
        </article>
      `
    )
    .join("");

  activityCards.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", () => {
      activeEvent = events.find((event) => event.id === button.dataset.event) || events[0];
      statusFilter = "all";
      rewardFilter = "all";
      renderAll();
      showPage("rewards");
    });
  });
}

function copyIcon() {
  return '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 8h11v11H8z" /><path d="M5 15H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v1" /></svg>';
}

function normalizeMasterCode(entry) {
  if (Array.isArray(entry)) {
    const [code, reward, status, expires] = entry;
    return { code, reward, status, expiresAt: expires, eventName: "-", items: [] };
  }

  return {
    code: entry.code || "",
    eventName: entry.eventName || "-",
    reward: entry.reward || entry.itemEn || entry.itemName || "รางวัลพิเศษ",
    status: entry.status || "พร้อมใช้",
    expiresAt: entry.expiresAt || entry.expires || "-",
    itemReceived: entry.itemReceived || entry.itemEn || entry.itemName || "",
    items: Array.isArray(entry.items) ? entry.items : [],
  };
}

function renderCodes() {
  const codes = masterCodes.filter(isVisibleNow).map(normalizeMasterCode).filter((entry) => entry.code);

  if (!codes.length) {
    codeList.innerHTML = `
      <div class="code-empty">
        <strong>ยังไม่มี Master Code ในรอบนี้</strong>
        <span>เมื่อมีโค้ดจากชีต Excel หน้านี้จะอัปเดตอัตโนมัติ</span>
      </div>
    `;
    return;
  }

  const isExpired = (status) => /(หมดเขต|ใช้แล้ว)/i.test(normalize(status));

  codeList.innerHTML = codes
    .map(
      (entry) => `
        <div class="code-card${isExpired(entry.status) ? " is-expired" : ""}">
          <div class="code-card-top">
            <span class="code-event-tag">${escapeHtml(entry.eventName)}</span>
            <div class="code-card-meta">
              <span class="status ${statusClass(entry.status)}">${escapeHtml(entry.status)}</span>
              <span class="code-expiry-text">หมดเขต ${escapeHtml(entry.expiresAt)}</span>
            </div>
          </div>
          <div class="code-display-row">
            <code class="code-value-big">${escapeHtml(entry.code)}</code>
            <button class="copy-btn-card" type="button" data-code="${escapeHtml(entry.code)}" ${isExpired(entry.status) ? "disabled" : ""}>
              ${copyIcon()} คัดลอก
            </button>
          </div>
          ${
            entry.items.length
              ? `<div class="code-items-wrap">
                  ${entry.items
                    .map(
                      (item) => `
                        <span class="code-item-chip">
                          ${item.imageUrl ? `<img src="${escapeHtml(item.imageUrl)}" alt="${escapeHtml(item.itemEn)}" loading="lazy" />` : ""}
                          <span>
                            <b>${escapeHtml(item.itemEn || entry.reward)}</b>
                            <small>x${escapeHtml(item.amount || "1")}</small>
                          </span>
                        </span>
                      `
                    )
                    .join("")}
                </div>`
              : entry.reward
              ? `<div class="code-items-wrap"><span class="code-item-chip"><span><b>${escapeHtml(entry.reward)}</b></span></span></div>`
              : ""
          }
        </div>
      `
    )
    .join("");

  codeList.querySelectorAll(".copy-btn-card:not(:disabled)").forEach((button) => {
    button.addEventListener("click", async () => {
      await navigator.clipboard.writeText(button.dataset.code);
      showToast(`คัดลอกโค้ด ${button.dataset.code} แล้ว`);
    });
  });
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("is-visible");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove("is-visible"), 2200);
}

function renderAll() {
  currentPage = 1;
  renderTabs();
  renderSearchModeTabs();
  renderEventCover();
  renderEventExtra();
  renderDeliveryProgress();
  renderSummary();
  renderRewardAnnouncement();
  renderFilterOptions();
  renderSearchResult();
  renderWinners();
}

function showPage(pageId) {
  const safePageId = [...pageSections].some((section) => section.id === pageId) ? pageId : "home";
  pageSections.forEach((section) => {
    section.classList.toggle("is-active", section.id === safePageId);
  });

  pageLinks.forEach((link) => {
    link.classList.toggle("is-active", link.dataset.page === safePageId);
  });

  header.classList.remove("is-open");
  navToggle.setAttribute("aria-expanded", "false");
  window.scrollTo({ top: 0, behavior: "smooth" });

  if (location.hash !== `#${safePageId}`) {
    history.replaceState(null, "", `#${safePageId}`);
  }
}

searchInput.addEventListener("input", renderSearchResult);

clearSearch.addEventListener("click", () => {
  searchInput.value = "";
  renderSearchResult();
  searchInput.focus();
});

searchModeTabs?.querySelectorAll("button").forEach((button) => {
  button.addEventListener("click", () => {
    searchMode = button.dataset.searchMode || "uid";
    renderSearchModeTabs();
    renderSearchResult();
    searchInput.focus();
  });
});

statusFilterSelect?.addEventListener("change", () => {
  statusFilter = statusFilterSelect.value;
  currentPage = 1;
  renderWinners();
});

rewardFilterSelect?.addEventListener("change", () => {
  rewardFilter = rewardFilterSelect.value;
  currentPage = 1;
  renderWinners();
});

navToggle.addEventListener("click", () => {
  const isOpen = header.classList.toggle("is-open");
  navToggle.setAttribute("aria-expanded", String(isOpen));
});

pageLinks.forEach((link) => {
  link.addEventListener("click", () => {
    showPage(link.dataset.page);
  });
});

uidLookupInput?.addEventListener("input", renderUidLookupResult);

// Sortable table headers
document.querySelectorAll(".sortable-table th.sortable").forEach((th) => {
  th.addEventListener("click", () => sortWinnersBy(th.dataset.sort));
});

// CSV export
document.getElementById("export-csv-btn")?.addEventListener("click", exportWinnersCsv);

// Try to load live-published data (from the admin "เผยแพร่" button).
// Falls back silently to the baked-in events-data.js on any error.
async function loadLiveData() {
  try {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 4000);
    const res = await fetch("/api/data", { signal: ctrl.signal, cache: "no-store" });
    clearTimeout(timer);
    if (!res.ok) return false; // 204 = nothing published yet
    const data = await res.json();
    if (Array.isArray(data.events) && data.events.length) {
      _allEventsRaw = data.events;
      const pub = _allEventsRaw.filter(isVisibleNow);
      events = pub.length ? pub : fallbackEvents;
      masterCodes = Array.isArray(data.codes) ? data.codes : masterCodes;
      activeEvent = events.find((e) => e.id === activeEvent?.id) || events[0];
      return true;
    }
  } catch (_) {
    /* keep baked-in fallback data */
  }
  return false;
}

function renderContent() {
  renderHeroStats();
  renderHeroRecent();
  renderActivityCards();
  renderCodes();
  renderAll();
  renderUidLookupResult();
}

function applyDeepLink() {
  const q = new URLSearchParams(location.search).get("q");
  if (!q) return false;
  searchMode = "uid";
  renderSearchModeTabs();
  searchInput.value = q;
  renderSearchResult();
  showPage("rewards");
  setTimeout(() => searchResult.scrollIntoView({ behavior: "smooth", block: "center" }), 200);
  return true;
}

// Re-check scheduled visibility while the page is open, so events/codes flip
// on/off automatically at their publishAt / hideAt time without a manual reload.
function recomputeScheduledVisibility() {
  const pub = _allEventsRaw.filter(isVisibleNow);
  const next = pub.length ? pub : fallbackEvents;
  const changed = next.length !== events.length || next.some((e, i) => e.id !== events[i]?.id);
  if (changed) {
    events = next;
    activeEvent = events.find((e) => e.id === activeEvent?.id) || events[0];
    renderContent();
  } else {
    renderCodes(); // codes filter live; cheap to refresh in case a code flipped
  }
}

(async function bootstrap() {
  // Render baked-in data immediately (no blank wait), then refresh if live data exists
  renderContent();
  if (!applyDeepLink()) showPage(location.hash.replace("#", "") || "home");
  const changed = await loadLiveData();
  if (changed) { renderContent(); applyDeepLink(); }
  setInterval(recomputeScheduledVisibility, 30000); // every 30s
})();
