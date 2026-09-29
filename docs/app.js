/* =============================================
   PhD Timeline Gantt — Application Logic
   =============================================
   Reads .xlsx from OneDrive (via CORS proxy or
   embedded JSON fallback), parses with SheetJS,
   and renders an interactive Gantt chart.
   ============================================= */

(function () {
  'use strict';

  // ---- Configuration ----
  const ONEDRIVE_SHARE_URL =
    'https://anu365-my.sharepoint.com/:x:/r/personal/u8363323_anu_edu_au/Documents/Doctoral%20Documents/QTimeline/PhD_Timeline_ANU.xlsx?d=wca152be551d24e248e6073fba2c31b4f&csf=1&web=1&e=TR8lMw;

  // OneDrive download URL (derived from share link)
  const ONEDRIVE_DOWNLOAD_URL =
    'https://anu365-my.sharepoint.com/personal/u8363323_anu_edu_au/_layouts/15/download.aspx?UniqueId=ca152be5-51d2-4e24-8e60-73fba2c31b4f';

  const TIMELINE_SHEET = 'Timeline';
  const DATA_ROW_START = 14; // 1-indexed row where task data begins
  const DATA_ROW_END = 57;   // 1-indexed row where task data ends
  const ENROLMENT_DATE_CELL = 'D3'; // Cell containing enrolment start date

  const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // ---- DOM references ----
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  const dom = {
    loadingOverlay: $('#loading-overlay'),
    errorOverlay: $('#error-overlay'),
    errorMessage: $('#error-message'),
    retryBtn: $('#retry-btn'),
    app: $('#app'),
    headerStats: $('#header-stats'),
    keyDates: $('#key-dates'),
    phaseFilter: $('#phase-filter'),
    statusFilter: $('#status-filter'),
    typeFilter: $('#type-filter'),
    taskPanelBody: $('#task-panel-body'),
    timelineHeader: $('#timeline-header'),
    timelineBody: $('#timeline-body'),
    tooltip: $('#tooltip'),
    refreshBtn: $('#refresh-btn'),
    lastRefresh: $('#last-refresh'),
  };

  // ---- State ----
  let allTasks = [];
  let enrolmentDate = null;
  let keyDatesData = [];
  let phases = [];

  // ---- Utility functions ----

  /** Convert Excel serial date to JS Date */
  function excelDateToJS(serial) {
    if (serial instanceof Date) return serial;
    if (typeof serial === 'string') {
      const d = new Date(serial);
      if (!isNaN(d)) return d;
    }
    if (typeof serial === 'number') {
      // Excel epoch: 1 Jan 1900, but Excel has a leap-year bug (day 60 = 29 Feb 1900)
      const epoch = new Date(1899, 11, 30);
      return new Date(epoch.getTime() + serial * 86400000);
    }
    return null;
  }

  /** Add N months to a date (replicates Excel EDATE) */
  function addMonths(date, n) {
    const d = new Date(date);
    d.setMonth(d.getMonth() + n);
    return d;
  }

  /** Format date as "dd Mon yyyy" */
  function fmtDate(d) {
    if (!d) return '—';
    return `${d.getDate()} ${MONTH_NAMES[d.getMonth()]} ${d.getFullYear()}`;
  }

  /** Relative time description */
  function relativeTime(d) {
    if (!d) return '';
    const now = new Date();
    const diffMs = d - now;
    const diffDays = Math.round(diffMs / 86400000);
    if (diffDays < 0) return `${Math.abs(diffDays)} days ago`;
    if (diffDays === 0) return 'today';
    if (diffDays < 30) return `in ${diffDays} days`;
    const diffMonths = Math.round(diffDays / 30.44);
    if (diffMonths < 12) return `in ${diffMonths} month${diffMonths > 1 ? 's' : ''}`;
    const diffYears = (diffDays / 365.25).toFixed(1);
    return `in ~${diffYears} years`;
  }

  /** Days between two dates */
  function daysBetween(a, b) {
    return Math.round((b - a) / 86400000);
  }

  /** Check if due date is less than 1 month from today */
  function isDueLessThanOneMonth(endDate) {
    if (!endDate) return false;
    const now = new Date();
    const diff = endDate - now;
    // Less than 30 days AND not yet past (or already past = also urgent)
    return diff < 30 * 86400000;
  }

  /** Determine highlight class for a task row */
  function getHighlightClass(task) {
    const status = (task.status || '').toLowerCase().trim();

    // Done = green
    if (status === 'done') return 'highlight-done';

    const urgent = isDueLessThanOneMonth(task.endDate);
    if (!urgent) return '';

    // Due < 1 month checks
    if (status === 'in progress') return 'highlight-urgent-progress';
    if (status === 'not started') return 'highlight-urgent-not-started';
    if (status === 'delayed') return 'highlight-urgent-delayed';

    return '';
  }

  /** Determine bar colour class */
  function getBarClass(task) {
    const status = (task.status || '').toLowerCase().trim();
    if (status === 'done') return 'bar-done';

    const urgent = isDueLessThanOneMonth(task.endDate);
    if (urgent && status === 'in progress') return 'bar-urgent-progress';
    if (urgent && status === 'not started') return 'bar-urgent-not-started';
    if (urgent && status === 'delayed') return 'bar-urgent-delayed';

    return 'bar-default';
  }

  /** Status badge HTML */
  function statusBadge(status) {
    const s = (status || '').toLowerCase().trim();
    const cls = s === 'done' ? 'done'
      : s === 'in progress' ? 'in-progress'
        : s === 'delayed' ? 'delayed'
          : 'not-started';
    const label = s === 'in progress' ? 'In Progress'
      : s === 'not started' ? 'Not Started'
        : status;
    return `<span class="status-badge ${cls}">${label}</span>`;
  }

  /** Type badge HTML */
  function typeBadge(type) {
    const t = (type || '').toLowerCase().trim();
    const cls = t === 'milestone' ? 'milestone' : 'task';
    return `<span class="type-badge ${cls}">${type}</span>`;
  }

  /** Get phase index for colouring */
  function phaseIndex(phaseName) {
    const idx = phases.indexOf(phaseName);
    return idx >= 0 ? idx % 5 : 0;
  }

  // ---- Data loading ----

  async function loadFromOneDrive() {
    // Try multiple strategies to get the file

    // Strategy 1: Try fetch as arraybuffer with CORS (will work for public files)
    const strategies = [
      // OneDrive direct download
      ONEDRIVE_DOWNLOAD_URL,
      // Alternative: use ?download=1 parameter
      ONEDRIVE_SHARE_URL.replace('&e=', '&download=1&e='),
    ];

    let arrayBuffer = null;

    for (const url of strategies) {
      try {
        const resp = await fetch(url, {
          mode: 'cors',
          credentials: 'include',
        });
        if (resp.ok) {
          arrayBuffer = await resp.arrayBuffer();
          break;
        }
      } catch (e) {
        // CORS error — expected, try next strategy
        console.warn('Fetch failed for', url, e.message);
      }
    }

    // Strategy 2: Try CORS proxy
    if (!arrayBuffer) {
      const proxyUrls = [
        `https://api.allorigins.win/raw?url=${encodeURIComponent(ONEDRIVE_DOWNLOAD_URL)}`,
        `https://corsproxy.io/?${encodeURIComponent(ONEDRIVE_DOWNLOAD_URL)}`,
      ];

      for (const proxyUrl of proxyUrls) {
        try {
          const resp = await fetch(proxyUrl);
          if (resp.ok) {
            arrayBuffer = await resp.arrayBuffer();
            break;
          }
        } catch (e) {
          console.warn('Proxy fetch failed:', e.message);
        }
      }
    }

    if (arrayBuffer) {
      return parseWorkbook(arrayBuffer);
    }

    // Strategy 3: Fall back to embedded data
    return loadFromFallback();
  }

  async function loadFromFallback() {
    // Try to load from a local data.json (generated by the convert script)
    try {
      const resp = await fetch('data.json');
      if (resp.ok) {
        const data = await resp.json();
        enrolmentDate = new Date(data.enrolmentDate);
        allTasks = data.tasks.map(t => ({
          ...t,
          startDate: t.startDate ? new Date(t.startDate) : null,
          endDate: t.endDate ? new Date(t.endDate) : null,
        }));
        keyDatesData = (data.keyDates || []).map(kd => ({
          ...kd,
          date: kd.date ? new Date(kd.date) : null,
        }));
        return true;
      }
    } catch (e) {
      // noop
    }

    throw new Error(
      'Could not load the Excel file from OneDrive (CORS restriction). ' +
      'Please run the converter script to generate data.json, or open this page from a local server.'
    );
  }

  function parseWorkbook(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    const ws = wb.Sheets[TIMELINE_SHEET];

    if (!ws) {
      throw new Error(`Sheet "${TIMELINE_SHEET}" not found in workbook.`);
    }

    // Get enrolment date from D3
    const enrolCell = ws[ENROLMENT_DATE_CELL];
    enrolmentDate = enrolCell ? excelDateToJS(enrolCell.v) : new Date('2026-09-29');

    // Compute key dates
    keyDatesData = [
      { label: 'Enrolment Start', date: new Date(enrolmentDate) },
      { label: 'CoC Target (M9)', date: addMonths(enrolmentDate, 9) },
      { label: 'Final Seminar (M30)', date: addMonths(enrolmentDate, 30) },
      { label: 'Thesis Submission (M36)', date: addMonths(enrolmentDate, 36) },
      { label: 'Scholarship End (M42)', date: (() => { const d = addMonths(enrolmentDate, 42); d.setDate(d.getDate() - 1); return d; })() },
      { label: 'MEP (M48)', date: (() => { const d = addMonths(enrolmentDate, 48); d.setDate(d.getDate() - 1); return d; })() },
    ];

    // Parse task rows
    allTasks = [];
    for (let r = DATA_ROW_START; r <= DATA_ROW_END; r++) {
      const cell = (col) => {
        const ref = XLSX.utils.encode_cell({ r: r - 1, c: col });
        return ws[ref] ? ws[ref].v : null;
      };

      const id = cell(0);
      if (!id) continue;

      const startM = cell(4);
      const endM = cell(5);

      // Compute dates using enrolment date + month offsets
      const startDate = startM != null ? addMonths(enrolmentDate, startM - 1) : null;
      let endDate = endM != null ? addMonths(enrolmentDate, endM) : null;
      if (endDate) {
        endDate.setDate(endDate.getDate() - 1);
      }

      allTasks.push({
        id,
        phase: cell(1) || '',
        task: cell(2) || '',
        type: cell(3) || 'Task',
        startM: startM || 1,
        endM: endM || 1,
        startDate,
        endDate,
        owner: cell(8) || '',
        status: cell(9) || 'Not started',
        notes: cell(10) || '',
      });
    }

    return true;
  }

  // ---- Rendering ----

  function computeMonthColumns() {
    // Generate month columns from enrolment start to 42 months after
    if (!enrolmentDate) return [];
    const cols = [];
    for (let m = 0; m < 44; m++) {
      const d = addMonths(enrolmentDate, m);
      cols.push({
        month: d.getMonth(),
        year: d.getFullYear(),
        date: d,
        monthNum: m + 1, // 1-indexed month of PhD
        label: MONTH_NAMES[d.getMonth()],
        yearLabel: String(d.getFullYear()).slice(-2),
      });
    }
    return cols;
  }

  function render() {
    const filteredTasks = getFilteredTasks();
    const monthCols = computeMonthColumns();
    const now = new Date();
    const currentMonthIdx = monthCols.findIndex(
      c => c.month === now.getMonth() && c.year === now.getFullYear()
    );

    renderStats();
    renderKeyDates();
    renderPhaseFilter();
    renderTimelineHeader(monthCols, currentMonthIdx);
    renderTasksAndBars(filteredTasks, monthCols, currentMonthIdx);
    renderTodayLine(monthCols);
  }

  function getFilteredTasks() {
    const phaseVal = dom.phaseFilter.value;
    const statusVal = dom.statusFilter.value;
    const typeVal = dom.typeFilter.value;

    return allTasks.filter(t => {
      if (phaseVal !== 'all' && t.phase !== phaseVal) return false;
      if (statusVal !== 'all' && t.status !== statusVal) return false;
      if (typeVal !== 'all' && t.type !== typeVal) return false;
      return true;
    });
  }

  function renderStats() {
    const total = allTasks.length;
    const done = allTasks.filter(t => t.status.toLowerCase() === 'done').length;
    const inProgress = allTasks.filter(t => t.status.toLowerCase() === 'in progress').length;
    const pct = total > 0 ? Math.round((done / total) * 100) : 0;

    dom.headerStats.innerHTML = `
      <div class="stat-chip done">
        <span>Done</span>
        <span class="stat-value">${done}/${total}</span>
      </div>
      <div class="stat-chip progress">
        <span>In Progress</span>
        <span class="stat-value">${inProgress}</span>
      </div>
      <div class="stat-chip">
        <span>Progress</span>
        <span class="stat-value">${pct}%</span>
      </div>
    `;
  }

  function renderKeyDates() {
    dom.keyDates.innerHTML = keyDatesData.map(kd => `
      <div class="key-date-card">
        <span class="kd-label">${kd.label}</span>
        <span class="kd-value">${fmtDate(kd.date)}</span>
        <span class="kd-relative">${relativeTime(kd.date)}</span>
      </div>
    `).join('');
  }

  function renderPhaseFilter() {
    phases = [...new Set(allTasks.map(t => t.phase))];
    const current = dom.phaseFilter.value;
    dom.phaseFilter.innerHTML = '<option value="all">All Phases</option>' +
      phases.map(p => `<option value="${p}" ${p === current ? 'selected' : ''}>${p}</option>`).join('');
  }

  function renderTimelineHeader(monthCols, currentMonthIdx) {
    dom.timelineHeader.innerHTML = monthCols.map((c, i) => `
      <div class="month-header ${i === currentMonthIdx ? 'current-month' : ''}">
        <span class="month-label">${c.label}</span>
        <span class="year-label">'${c.yearLabel}</span>
      </div>
    `).join('');
  }

  function renderTasksAndBars(tasks, monthCols, currentMonthIdx) {
    dom.taskPanelBody.innerHTML = '';
    dom.timelineBody.innerHTML = '';

    const totalMonths = monthCols.length;
    const monthW = 60; // px, matches CSS --month-w

    let lastPhase = '';
    let rowIndex = 0;

    tasks.forEach((task, idx) => {
      const animDelay = `animation-delay: ${idx * 20}ms;`;

      // Phase header
      if (task.phase !== lastPhase) {
        lastPhase = task.phase;
        const pIdx = phaseIndex(task.phase);

        // Task panel phase header
        const phEl = document.createElement('div');
        phEl.className = `phase-header phase-${pIdx}`;
        phEl.style.cssText = animDelay;
        phEl.textContent = task.phase;
        dom.taskPanelBody.appendChild(phEl);

        // Timeline phase row (empty spacer)
        const tlPhEl = document.createElement('div');
        tlPhEl.className = 'timeline-phase-row';
        tlPhEl.style.cssText = animDelay;
        for (let i = 0; i < totalMonths; i++) {
          const cell = document.createElement('div');
          cell.className = `timeline-cell ${i === currentMonthIdx ? 'current-month-col' : ''}`;
          tlPhEl.appendChild(cell);
        }
        dom.timelineBody.appendChild(tlPhEl);
        rowIndex++;
      }

      // Task row (left panel)
      const highlight = getHighlightClass(task);
      const trEl = document.createElement('div');
      trEl.className = `task-row ${highlight}`;
      trEl.style.cssText = animDelay;
      trEl.dataset.taskId = task.id;
      trEl.innerHTML = `
        <div class="tp-cell tp-id">${task.id}</div>
        <div class="tp-cell tp-task" title="${task.task}">${task.task}</div>
        <div class="tp-cell tp-type">${typeBadge(task.type)}</div>
        <div class="tp-cell tp-status">${statusBadge(task.status)}</div>
        <div class="tp-cell tp-owner">${task.owner}</div>
      `;
      dom.taskPanelBody.appendChild(trEl);

      // Timeline row (right panel)
      const tlRowEl = document.createElement('div');
      tlRowEl.className = 'timeline-row';
      tlRowEl.style.cssText = animDelay;

      // Background cells
      for (let i = 0; i < totalMonths; i++) {
        const cell = document.createElement('div');
        cell.className = `timeline-cell ${i === currentMonthIdx ? 'current-month-col' : ''}`;
        tlRowEl.appendChild(cell);
      }

      // Gantt bar
      const barClass = getBarClass(task);
      const startCol = task.startM - 1; // 0-indexed column
      const endCol = task.endM;         // endM is 1-indexed, so endCol is exclusive
      const barLeft = startCol * monthW;
      const barWidth = (endCol - startCol) * monthW;

      if (task.type === 'Milestone') {
        // Diamond milestone marker
        const bar = document.createElement('div');
        bar.className = `gantt-bar bar-milestone ${barClass}`;
        bar.style.left = `${barLeft + monthW / 2 - 11}px`;
        bar.style.cssText += animDelay;
        bar.innerHTML = '<div class="bar-fill"></div>';
        bar.addEventListener('mouseenter', (e) => showTooltip(e, task));
        bar.addEventListener('mouseleave', hideTooltip);
        bar.addEventListener('mousemove', moveTooltip);
        tlRowEl.appendChild(bar);
      } else {
        // Regular bar
        const bar = document.createElement('div');
        bar.className = `gantt-bar ${barClass}`;
        bar.style.left = `${barLeft}px`;
        bar.style.width = `${Math.max(barWidth, monthW)}px`;
        bar.style.cssText += animDelay;
        bar.innerHTML = `
          <div class="bar-fill" style="width: 100%;"></div>
          ${barWidth > monthW * 2 ? `<span class="bar-label">${task.task}</span>` : ''}
        `;
        bar.addEventListener('mouseenter', (e) => showTooltip(e, task));
        bar.addEventListener('mouseleave', hideTooltip);
        bar.addEventListener('mousemove', moveTooltip);
        tlRowEl.appendChild(bar);
      }

      dom.timelineBody.appendChild(tlRowEl);
      rowIndex++;
    });
  }

  function renderTodayLine(monthCols) {
    // Remove existing today lines
    dom.timelineBody.querySelectorAll('.today-line').forEach(el => el.remove());

    const now = new Date();
    const monthW = 60;

    // Find which month column "today" falls in
    for (let i = 0; i < monthCols.length; i++) {
      const colDate = monthCols[i].date;
      const nextDate = i + 1 < monthCols.length ? monthCols[i + 1].date : addMonths(colDate, 1);

      if (now >= colDate && now < nextDate) {
        // Calculate fractional position within this month
        const totalDays = daysBetween(colDate, nextDate);
        const elapsed = daysBetween(colDate, now);
        const frac = totalDays > 0 ? elapsed / totalDays : 0;
        const xPos = (i + frac) * monthW;

        const line = document.createElement('div');
        line.className = 'today-line';
        line.style.left = `${xPos}px`;
        dom.timelineBody.appendChild(line);
        break;
      }
    }
  }

  // ---- Tooltip ----

  function showTooltip(e, task) {
    const urgent = isDueLessThanOneMonth(task.endDate);
    const daysLeft = task.endDate ? daysBetween(new Date(), task.endDate) : null;

    dom.tooltip.innerHTML = `
      <div class="tt-title">${task.task}</div>
      <div class="tt-row"><span class="tt-label">ID</span><span class="tt-value">${task.id}</span></div>
      <div class="tt-row"><span class="tt-label">Phase</span><span class="tt-value">${task.phase}</span></div>
      <div class="tt-row"><span class="tt-label">Type</span><span class="tt-value">${task.type}</span></div>
      <div class="tt-row"><span class="tt-label">Status</span><span class="tt-value">${statusBadge(task.status)}</span></div>
      <div class="tt-row"><span class="tt-label">Start</span><span class="tt-value">${fmtDate(task.startDate)} (M${task.startM})</span></div>
      <div class="tt-row"><span class="tt-label">Due</span><span class="tt-value">${fmtDate(task.endDate)} (M${task.endM})</span></div>
      <div class="tt-row"><span class="tt-label">Owner</span><span class="tt-value">${task.owner}</span></div>
      ${daysLeft !== null ? `<div class="tt-row"><span class="tt-label">Days left</span><span class="tt-value" style="color: ${urgent ? 'var(--clr-not-started-urgent)' : 'var(--clr-text)'}">${daysLeft < 0 ? `${Math.abs(daysLeft)} days overdue` : `${daysLeft} days`}</span></div>` : ''}
      ${task.notes ? `<div class="tt-notes">📝 ${task.notes}</div>` : ''}
    `;
    dom.tooltip.classList.remove('hidden');
    moveTooltip(e);
  }

  function moveTooltip(e) {
    const pad = 16;
    const ttRect = dom.tooltip.getBoundingClientRect();
    let x = e.clientX + pad;
    let y = e.clientY + pad;

    if (x + ttRect.width > window.innerWidth - pad) {
      x = e.clientX - ttRect.width - pad;
    }
    if (y + ttRect.height > window.innerHeight - pad) {
      y = e.clientY - ttRect.height - pad;
    }

    dom.tooltip.style.left = `${x}px`;
    dom.tooltip.style.top = `${y}px`;
  }

  function hideTooltip() {
    dom.tooltip.classList.add('hidden');
  }

  // ---- Event listeners ----

  function setupEvents() {
    dom.phaseFilter.addEventListener('change', render);
    dom.statusFilter.addEventListener('change', render);
    dom.typeFilter.addEventListener('change', render);
    dom.refreshBtn.addEventListener('click', () => {
      dom.refreshBtn.classList.add('spinning');
      init().finally(() => {
        dom.refreshBtn.classList.remove('spinning');
      });
    });
    dom.retryBtn.addEventListener('click', init);

    // Sync scroll between task panel and timeline
    const wrapper = $('#gantt-wrapper');
    if (wrapper) {
      // Scroll task rows in sync with timeline body (handled by shared parent scroll)
    }

    // Highlight matching rows on hover
    document.addEventListener('mouseover', (e) => {
      const row = e.target.closest('.task-row');
      if (row) {
        const taskId = row.dataset.taskId;
        // Find corresponding timeline row
        const idx = [...dom.taskPanelBody.children].indexOf(row);
        const tlRow = dom.timelineBody.children[idx];
        if (tlRow) tlRow.style.background = 'var(--clr-surface-2)';
      }
    });

    document.addEventListener('mouseout', (e) => {
      const row = e.target.closest('.task-row');
      if (row) {
        const idx = [...dom.taskPanelBody.children].indexOf(row);
        const tlRow = dom.timelineBody.children[idx];
        if (tlRow) tlRow.style.background = '';
      }
    });
  }

  // ---- Scroll to today ----

  function scrollToToday() {
    const todayLine = dom.timelineBody.querySelector('.today-line');
    if (todayLine) {
      const panel = $('#timeline-panel');
      const lineLeft = parseInt(todayLine.style.left);
      panel.scrollLeft = Math.max(0, lineLeft - panel.clientWidth / 3);
    }
  }

  // ---- Initialization ----

  async function init() {
    dom.loadingOverlay.classList.remove('hidden');
    dom.errorOverlay.classList.add('hidden');
    dom.app.classList.add('hidden');

    try {
      await loadFromOneDrive();
      dom.loadingOverlay.classList.add('hidden');
      dom.app.classList.remove('hidden');
      render();

      const now = new Date();
      dom.lastRefresh.textContent = `${fmtDate(now)} ${now.toLocaleTimeString()}`;

      // Scroll to today after a short delay for layout
      setTimeout(scrollToToday, 300);
    } catch (err) {
      console.error('Failed to load timeline:', err);
      dom.loadingOverlay.classList.add('hidden');
      dom.errorOverlay.classList.remove('hidden');
      dom.errorMessage.textContent = err.message;
    }
  }

  // ---- Boot ----
  setupEvents();
  init();
})();
