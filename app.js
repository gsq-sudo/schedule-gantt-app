(function () {
  "use strict";

  const APP_STORAGE_KEY = "schedule-gantt:data";
  const CLIENT_STORAGE_KEY = "project-ledger:google-client-id";
  const DRIVE_ENABLED_KEY = "schedule-gantt:drive-enabled";
  const DRIVE_FILE_NAME = "schedule-gantt-data.json";
  const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.appdata";
  const MS_PER_DAY = 86400000;
  const MIN_DAY_WIDTH = 28;
  const MAX_DAY_WIDTH = 64;
  const ROW_AUTOSCROLL_EDGE = 88;
  const ROW_AUTOSCROLL_MAX_SPEED = 24;
  const OFFICE_COLOR_PALETTE = [
    { name: "Blue", colors: ["#deebf7", "#9dc3e6", "#5b9bd5", "#2f75b5", "#1f4e79"] },
    { name: "Orange", colors: ["#fce4d6", "#f8cbad", "#ed7d31", "#c55a11", "#843c0c"] },
    { name: "Gray", colors: ["#f2f2f2", "#d9e2f3", "#a6a6a6", "#666666", "#262626"] },
    { name: "Gold", colors: ["#fff2cc", "#ffe699", "#ffc000", "#bf9000", "#7f6000"] },
    { name: "Green", colors: ["#e2f0d9", "#a9d18e", "#70ad47", "#548235", "#375623"] },
    { name: "Teal", colors: ["#ddebf7", "#9cc2e5", "#00a2e8", "#0070c0", "#004c7f"] },
    { name: "Purple", colors: ["#e4dfec", "#b4a7d6", "#8064a2", "#5f497a", "#3f3151"] },
    { name: "Red", colors: ["#f4cccc", "#e06666", "#c00000", "#990000", "#660000"] },
    { name: "Aqua", colors: ["#d9ead3", "#93cddd", "#4bacc6", "#31859b", "#205867"] },
    { name: "Rose", colors: ["#eadcf8", "#d5a6bd", "#c0504d", "#953735", "#632423"] }
  ];
  const COLOR_POOL = OFFICE_COLOR_PALETTE.map((group) => group.colors[2]);

  const app = document.getElementById("app");
  let rowAutoScrollFrame = 0;

  const state = {
    activeTab: "gantt",
    ownerFilter: "all",
    selectedProjectId: "",
    drag: null,
    rowDrag: null,
    drive: {
      accessToken: "",
      tokenExpiresAt: 0,
      fileId: "",
      modifiedTime: "",
      connected: false,
      enabled: localStorage.getItem(DRIVE_ENABLED_KEY) === "true",
      pending: false,
      pendingAction: "",
      message: localStorage.getItem(DRIVE_ENABLED_KEY) === "true" ? "Reconnect Drive" : "Local only",
      error: ""
    },
    tokenClient: null,
    tokenTimer: null,
    saveTimer: null,
    toastTimer: null,
    data: loadData()
  };

  function emptyData() {
    const now = new Date().toISOString();
    return {
      schemaVersion: 1,
      updatedAt: now,
      settings: {
        dayWidth: 38,
        showArchivedOnTimeline: true
      },
      projects: []
    };
  }

  function loadData() {
    try {
      const raw = localStorage.getItem(APP_STORAGE_KEY);
      if (!raw) return emptyData();
      return normalizeData(JSON.parse(raw));
    } catch (error) {
      console.warn("Unable to read local schedule data.", error);
      return emptyData();
    }
  }

  function normalizeData(raw) {
    const base = emptyData();
    const data = {
      ...base,
      ...raw,
      settings: {
        ...base.settings,
        ...(raw && raw.settings ? raw.settings : {})
      },
      projects: Array.isArray(raw && raw.projects) ? raw.projects : []
    };

    data.settings.dayWidth = clampNumber(Number(data.settings.dayWidth || base.settings.dayWidth), MIN_DAY_WIDTH, MAX_DAY_WIDTH);
    data.settings.showArchivedOnTimeline = data.settings.showArchivedOnTimeline !== false;
    data.projects = data.projects.map((project, index) => normalizeProject(project, index));
    return data;
  }

  function normalizeProject(project, index = 0) {
    const today = todayInput();
    const startDate = isDateInput(project && project.startDate) ? project.startDate : today;
    const endDate = isDateInput(project && project.endDate) ? project.endDate : addDays(startDate, 13);
    const safeEndDate = compareDates(endDate, startDate) < 0 ? startDate : endDate;
    const mode = project && project.mode === "allocation" ? "allocation" : "effort";
    const duration = diffDaysInclusive(startDate, safeEndDate);
    const rawEffortDays = project && project.effortDays !== undefined ? Number(project.effortDays) : 5;
    const effortDays = Math.max(0, Number.isFinite(rawEffortDays) ? rawEffortDays : 5);
    const fallbackDailyPercent = roundTo((effortDays / Math.max(1, workDaysInclusive(startDate, safeEndDate))) * 100, 1);
    const rawDailyPercent = project && project.dailyPercent !== undefined ? Number(project.dailyPercent) : fallbackDailyPercent;
    const dailyPercent = Math.max(0, Number.isFinite(rawDailyPercent) ? rawDailyPercent : fallbackDailyPercent);

    return {
      id: (project && project.id) || uid(),
      name: (project && project.name) || "Untitled project",
      owner: (project && project.owner) || "Unassigned",
      startDate,
      endDate: safeEndDate,
      mode,
      effortDays: roundTo(effortDays, 2),
      dailyPercent: roundTo(dailyPercent, 2),
      dependencies: Array.isArray(project && project.dependencies) ? project.dependencies.filter(Boolean) : [],
      notes: (project && project.notes) || "",
      color: normalizeColor(project && project.color, index),
      status: project && project.status === "archived" ? "archived" : "active",
      completedAt: (project && project.completedAt) || "",
      order: Number.isFinite(Number(project && project.order)) ? Number(project.order) : index,
      createdAt: (project && project.createdAt) || new Date().toISOString(),
      updatedAt: (project && project.updatedAt) || new Date().toISOString()
    };
  }

  function saveLocal({ sync = true, renderNow = true } = {}) {
    state.data.updatedAt = new Date().toISOString();
    localStorage.setItem(APP_STORAGE_KEY, JSON.stringify(state.data));
    if (sync && state.drive.enabled) scheduleDriveSave();
    if (renderNow) render();
  }

  function enableDriveSync() {
    state.drive.enabled = true;
    localStorage.setItem(DRIVE_ENABLED_KEY, "true");
  }

  function hasValidAccessToken() {
    return Boolean(
      state.drive.accessToken &&
        (!state.drive.tokenExpiresAt || Date.now() < state.drive.tokenExpiresAt - 30000)
    );
  }

  function makeAuthError(message) {
    const error = new Error(message);
    error.name = "DriveAuthRequired";
    return error;
  }

  function isAuthError(error) {
    return error && error.name === "DriveAuthRequired";
  }

  function markDriveNeedsAuth(action = state.drive.pendingAction || "save", message = "Reconnect Drive") {
    enableDriveSync();
    state.drive.accessToken = "";
    state.drive.tokenExpiresAt = 0;
    state.drive.connected = false;
    state.drive.pending = true;
    state.drive.pendingAction = action;
    state.drive.message = message;
    window.clearTimeout(state.tokenTimer);
  }

  function scheduleTokenExpiryNotice() {
    window.clearTimeout(state.tokenTimer);
    if (!state.drive.tokenExpiresAt) return;
    const delay = Math.max(0, state.drive.tokenExpiresAt - Date.now() - 30000);
    state.tokenTimer = window.setTimeout(() => {
      if (!hasValidAccessToken()) {
        markDriveNeedsAuth("", "Reconnect Drive");
        render();
      }
    }, delay);
  }

  function uid() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return window.crypto.randomUUID();
    }
    return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function escapeHtml(value) {
    return String(value === null || value === undefined ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function clampNumber(value, min, max) {
    if (!Number.isFinite(value)) return min;
    return Math.min(max, Math.max(min, value));
  }

  function roundTo(value, places = 1) {
    const factor = 10 ** places;
    return Math.round((Number(value) || 0) * factor) / factor;
  }

  function paletteColors() {
    return OFFICE_COLOR_PALETTE.flatMap((group) => group.colors);
  }

  function normalizeColor(value, index = 0) {
    const color = String(value || "").trim().toLowerCase();
    if (/^#[0-9a-f]{6}$/i.test(color)) return color;
    return COLOR_POOL[index % COLOR_POOL.length];
  }

  function todayInput() {
    return formatDateInput(new Date());
  }

  function isDateInput(value) {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "")) && Boolean(dateFromInput(value));
  }

  function dateFromInput(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) return null;
    const [year, month, day] = value.split("-").map(Number);
    const date = new Date(year, month - 1, day);
    if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
    date.setHours(0, 0, 0, 0);
    return date;
  }

  function formatDateInput(date) {
    const safeDate = new Date(date);
    safeDate.setHours(0, 0, 0, 0);
    const year = safeDate.getFullYear();
    const month = String(safeDate.getMonth() + 1).padStart(2, "0");
    const day = String(safeDate.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function addDays(value, days) {
    const date = dateFromInput(value);
    if (!date) return todayInput();
    date.setDate(date.getDate() + days);
    return formatDateInput(date);
  }

  function diffDays(startDate, endDate) {
    const start = dateFromInput(startDate);
    const end = dateFromInput(endDate);
    if (!start || !end) return 0;
    return Math.round((end.getTime() - start.getTime()) / MS_PER_DAY);
  }

  function diffDaysInclusive(startDate, endDate) {
    return Math.max(1, diffDays(startDate, endDate) + 1);
  }

  function workDaysInclusive(startDate, endDate) {
    if (!dateFromInput(startDate) || !dateFromInput(endDate) || compareDates(endDate, startDate) < 0) return 0;
    let current = startDate;
    let days = 0;
    const guard = Math.min(730, diffDaysInclusive(startDate, endDate));
    for (let index = 0; index < guard; index += 1) {
      if (!isWeekend(current)) days += 1;
      current = addDays(current, 1);
    }
    return days;
  }

  function nextWorkday(value) {
    let current = value;
    for (let index = 0; index < 14 && isWeekend(current); index += 1) {
      current = addDays(current, 1);
    }
    return current;
  }

  function compareDates(a, b) {
    const first = dateFromInput(a);
    const second = dateFromInput(b);
    if (!first || !second) return 0;
    return first.getTime() - second.getTime();
  }

  function formatShortDate(value) {
    const date = dateFromInput(value);
    if (!date) return "";
    return new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric" }).format(date);
  }

  function formatLongDate(value) {
    const date = dateFromInput(value);
    if (!date) return "";
    return new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric", year: "numeric" }).format(date);
  }

  function formatMonth(value) {
    const date = dateFromInput(value);
    if (!date) return "";
    return new Intl.DateTimeFormat(undefined, { month: "short" }).format(date);
  }

  function formatDateTime(value) {
    if (!value) return "Not set";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "Not set";
    return new Intl.DateTimeFormat(undefined, {
      month: "short",
      day: "numeric",
      year: "numeric",
      hour: "numeric",
      minute: "2-digit"
    }).format(date);
  }

  function projectById(id) {
    return state.data.projects.find((project) => project.id === id);
  }

  function projectName(id) {
    const project = projectById(id);
    return project ? project.name : "Missing project";
  }

  function sortedProjects(projects = state.data.projects) {
    return [...projects].sort((a, b) => {
      const orderSort = Number(a.order || 0) - Number(b.order || 0);
      if (orderSort !== 0) return orderSort;
      return a.name.localeCompare(b.name);
    });
  }

  function normalizeProjectOrder(projects = state.data.projects) {
    sortedProjects(projects).forEach((project, index) => {
      project.order = index;
    });
  }

  function nextProjectOrder() {
    return state.data.projects.reduce((max, project) => Math.max(max, Number(project.order || 0)), -1) + 1;
  }

  function owners() {
    return [...new Set(state.data.projects.map((project) => project.owner || "Unassigned"))].sort((a, b) => a.localeCompare(b));
  }

  function activeProjects() {
    return state.data.projects.filter((project) => project.status !== "archived");
  }

  function archivedProjects() {
    return state.data.projects.filter((project) => project.status === "archived");
  }

  function getVisibleProjects() {
    const timelineProjects = state.data.settings.showArchivedOnTimeline
      ? state.data.projects
      : state.data.projects.filter((project) => project.status !== "archived");
    const projects = state.ownerFilter === "all"
      ? timelineProjects
      : timelineProjects.filter((project) => (project.owner || "Unassigned") === state.ownerFilter);
    return sortedProjects(projects);
  }

  function getProjectMath(project) {
    const duration = diffDaysInclusive(project.startDate, project.endDate);
    const workDays = workDaysInclusive(project.startDate, project.endDate);
    if (project.mode === "allocation") {
      const dailyPercent = Math.max(0, Number(project.dailyPercent || 0));
      return {
        duration,
        workDays,
        dailyPercent: roundTo(dailyPercent, 1),
        effortDays: roundTo((workDays * dailyPercent) / 100, 2)
      };
    }

    const effortDays = Math.max(0, Number(project.effortDays || 0));
    return {
      duration,
      workDays,
      effortDays: roundTo(effortDays, 2),
      dailyPercent: roundTo(workDays ? (effortDays / workDays) * 100 : 0, 1)
    };
  }

  function getDraftProject() {
    const startDate = todayInput();
    const endDate = addDays(startDate, 20);
    return {
      id: "",
      name: "",
      owner: "",
      startDate,
      endDate,
      mode: "effort",
      effortDays: 5,
      dailyPercent: roundTo((5 / diffDaysInclusive(startDate, endDate)) * 100, 1),
      dependencies: [],
      notes: "",
      color: COLOR_POOL[state.data.projects.length % COLOR_POOL.length],
      status: "active",
      completedAt: ""
    };
  }

  function getEditorProject() {
    return projectById(state.selectedProjectId) || getDraftProject();
  }

  function sanitizeProjectDates(project) {
    if (!isDateInput(project.startDate)) project.startDate = todayInput();
    if (!isDateInput(project.endDate)) project.endDate = project.startDate;
    if (compareDates(project.endDate, project.startDate) < 0) project.endDate = project.startDate;
  }

  function getDependencyReadyStart(project) {
    let readyStart = "";
    project.dependencies.forEach((dependencyId) => {
      const dependency = projectById(dependencyId);
      if (!dependency || !dependency.endDate) return;
      const nextStart = nextWorkday(addDays(dependency.endDate, 1));
      if (!readyStart || compareDates(nextStart, readyStart) > 0) readyStart = nextStart;
    });
    return readyStart;
  }

  function shiftProjectToStart(project, newStartDate) {
    const duration = diffDaysInclusive(project.startDate, project.endDate);
    project.startDate = newStartDate;
    project.endDate = addDays(newStartDate, duration - 1);
    project.updatedAt = new Date().toISOString();
  }

  function applyDependencyCascade() {
    const shifted = new Set();
    const maxPasses = Math.max(2, state.data.projects.length + 2);
    for (let pass = 0; pass < maxPasses; pass += 1) {
      let changed = false;
      state.data.projects.forEach((project) => {
        sanitizeProjectDates(project);
        const readyStart = getDependencyReadyStart(project);
        if (readyStart && compareDates(project.startDate, readyStart) < 0) {
          shiftProjectToStart(project, readyStart);
          shifted.add(project.name);
          changed = true;
        }
      });
      if (!changed) break;
    }
    return [...shifted];
  }

  function hasDependencyPath(startId, targetId, dependenciesByProject, visited = new Set()) {
    if (startId === targetId) return true;
    if (visited.has(startId)) return false;
    visited.add(startId);
    const dependencies = dependenciesByProject.get(startId) || [];
    return dependencies.some((dependencyId) => hasDependencyPath(dependencyId, targetId, dependenciesByProject, visited));
  }

  function wouldCreateCycle(projectId, nextDependencies) {
    const dependenciesByProject = new Map(state.data.projects.map((project) => [project.id, [...project.dependencies]]));
    dependenciesByProject.set(projectId, nextDependencies);
    return nextDependencies.some((dependencyId) => hasDependencyPath(dependencyId, projectId, dependenciesByProject));
  }

  function getTimelineRange(projects) {
    if (!projects.length) {
      const startDate = todayInput();
      return {
        startDate,
        endDate: addDays(startDate, 27)
      };
    }

    let startDate = projects[0].startDate;
    let endDate = projects[0].endDate;
    projects.forEach((project) => {
      if (compareDates(project.startDate, startDate) < 0) startDate = project.startDate;
      if (compareDates(project.endDate, endDate) > 0) endDate = project.endDate;
    });
    return {
      startDate: addDays(startDate, -2),
      endDate: addDays(endDate, 7)
    };
  }

  function enumerateDays(startDate, endDate) {
    const days = [];
    let current = startDate;
    const guard = Math.min(730, diffDaysInclusive(startDate, endDate));
    for (let index = 0; index < guard; index += 1) {
      days.push(current);
      current = addDays(current, 1);
    }
    return days;
  }

  function isWeekend(value) {
    const date = dateFromInput(value);
    if (!date) return false;
    return date.getDay() === 0 || date.getDay() === 6;
  }

  function projectCoversDate(project, date) {
    return compareDates(project.startDate, date) <= 0 && compareDates(project.endDate, date) >= 0;
  }

  function ownerLoadForDay(owner, date) {
    if (isWeekend(date)) return 0;
    return state.data.projects.reduce((total, project) => {
      if (project.status === "archived" || (project.owner || "Unassigned") !== owner || !projectCoversDate(project, date)) return total;
      return total + getProjectMath(project).dailyPercent;
    }, 0);
  }

  function loadClass(value) {
    if (value > 100) return "over";
    if (value >= 85) return "high";
    if (value > 0) return "active";
    return "";
  }

  function getLoadStats(owner, days) {
    const loads = days.map((day) => ({ day, load: ownerLoadForDay(owner, day) }));
    const activeLoads = loads.filter((item) => item.load > 0);
    const max = loads.reduce((peak, item) => (item.load > peak.load ? item : peak), { day: "", load: 0 });
    const average = activeLoads.length
      ? activeLoads.reduce((sum, item) => sum + item.load, 0) / activeLoads.length
      : 0;
    return {
      loads,
      max,
      average: roundTo(average, 1)
    };
  }

  function render() {
    app.innerHTML = `
      <header class="topbar">
        <div class="brand">
          <div class="brand-mark"><i data-lucide="calendar-range"></i></div>
          <div>
            <h1>Schedule Gantt</h1>
            <p>${escapeHtml(activeProjects().length)} active · ${escapeHtml(archivedProjects().length)} archived · ${escapeHtml(owners().length)} owners</p>
          </div>
        </div>
        <div class="top-actions">
          ${renderSyncPill()}
          <button class="button icon-text" data-action="connect-drive" title="Connect Google Drive">
            <i data-lucide="${state.drive.enabled ? "refresh-cw" : "cloud"}"></i>
            <span>${escapeHtml(driveButtonLabel())}</span>
          </button>
          <button class="button icon-text" data-action="save-drive" title="Save to Drive">
            <i data-lucide="save"></i>
            <span>Save</span>
          </button>
          <button class="button icon-text" data-action="pull-drive" title="Pull from Drive">
            <i data-lucide="download-cloud"></i>
            <span>Pull</span>
          </button>
        </div>
      </header>
      <main class="layout">
        <nav class="tabs" aria-label="Views">
          ${renderTab("gantt", "chart-gantt", "Gantt")}
          ${renderTab("projects", "table-2", "Projects")}
          ${renderTab("settings", "settings", "Settings")}
        </nav>
        ${state.activeTab === "gantt" ? renderGanttView() : ""}
        ${state.activeTab === "projects" ? renderProjectsView() : ""}
        ${state.activeTab === "settings" ? renderSettingsView() : ""}
      </main>
      <div id="toast" class="toast" role="status" aria-live="polite"></div>
    `;

    refreshIcons();
    updateFormMath(document.getElementById("project-form"));
  }

  function refreshIcons() {
    if (window.lucide && typeof window.lucide.createIcons === "function") {
      window.lucide.createIcons();
    }
  }

  function renderTab(id, icon, label) {
    return `
      <button class="tab ${state.activeTab === id ? "active" : ""}" data-tab="${escapeHtml(id)}">
        <i data-lucide="${escapeHtml(icon)}"></i>
        <span>${escapeHtml(label)}</span>
      </button>
    `;
  }

  function renderSyncPill() {
    const statusClass = state.drive.error ? "error" : state.drive.pending ? "pending" : state.drive.connected ? "online" : "";
    return `
      <span class="sync-pill" title="${escapeHtml(state.drive.error || state.drive.message)}">
        <span class="sync-dot ${statusClass}"></span>
        <span>${escapeHtml(state.drive.message)}</span>
      </span>
    `;
  }

  function driveButtonLabel() {
    if (state.drive.pendingAction) return "Resume Drive";
    if (state.drive.enabled) return "Resume Drive";
    return "Connect Drive";
  }

  function renderGanttView() {
    const visibleProjects = getVisibleProjects();
    const range = getTimelineRange(visibleProjects);
    const days = enumerateDays(range.startDate, range.endDate);
    const dayWidth = state.data.settings.dayWidth;
    const totalWidth = days.length * dayWidth;
    const selectedProject = projectById(state.selectedProjectId);
    const shiftedInfo = selectedProject ? getDependencyReadyStart(selectedProject) : "";

    return `
      <section class="view active">
        <div class="summary-strip">
          ${renderMetric("Projects", state.data.projects.length)}
          ${renderMetric("Active", activeProjects().length)}
          ${renderMetric("Archived", archivedProjects().length)}
          ${renderMetric("Owners", owners().length)}
          ${renderMetric("Timeline", `${formatShortDate(range.startDate)} - ${formatShortDate(range.endDate)}`)}
        </div>
        <div class="workbench">
          <aside class="editor-panel">
            <div class="panel-header">
              <div>
                <h2 class="panel-title">${selectedProject ? "Edit project" : "New project"}</h2>
                <p class="panel-subtitle">${shiftedInfo ? `Dependency-ready ${formatLongDate(shiftedInfo)}` : "Schedule details"}</p>
              </div>
              <button class="icon-button ghost" data-action="new-project" title="New project" aria-label="New project">
                <i data-lucide="plus"></i>
              </button>
            </div>
            ${renderProjectForm()}
          </aside>
          <section class="chart-panel">
            <div class="chart-toolbar">
              <label class="select-field">
                <span>Owner</span>
                <select id="owner-filter">
                  <option value="all" ${state.ownerFilter === "all" ? "selected" : ""}>All owners</option>
                  ${owners().map((owner) => `<option value="${escapeHtml(owner)}" ${state.ownerFilter === owner ? "selected" : ""}>${escapeHtml(owner)}</option>`).join("")}
                </select>
              </label>
              <label class="toggle-field" title="Show completed projects on the Gantt timeline">
                <input id="show-archived-toggle" type="checkbox" ${state.data.settings.showArchivedOnTimeline ? "checked" : ""} />
                <span>Show completed</span>
              </label>
              <div class="toolbar-group" aria-label="Timeline zoom">
                <button class="icon-button" data-action="zoom-out" title="Zoom out" aria-label="Zoom out">
                  <i data-lucide="minus"></i>
                </button>
                <button class="icon-button" data-action="zoom-in" title="Zoom in" aria-label="Zoom in">
                  <i data-lucide="plus"></i>
                </button>
              </div>
            </div>
            ${renderGanttChart(visibleProjects, days, totalWidth, dayWidth)}
            ${renderUtilization(days, totalWidth, dayWidth)}
          </section>
        </div>
      </section>
    `;
  }

  function renderMetric(label, value) {
    return `
      <div class="metric">
        <span>${escapeHtml(label)}</span>
        <strong>${escapeHtml(value)}</strong>
      </div>
    `;
  }

  function renderProjectForm() {
    const project = getEditorProject();
    const math = getProjectMath(project);
    const isArchived = project.status === "archived";
    const dependencyOptions = sortedProjects()
      .filter((candidate) => candidate.id !== project.id)
      .map((candidate) => {
        const selected = project.dependencies.includes(candidate.id) ? "selected" : "";
        return `<option value="${escapeHtml(candidate.id)}" ${selected}>${escapeHtml(candidate.name)} (${escapeHtml(candidate.owner)})</option>`;
      })
      .join("");

    return `
      <form id="project-form" class="project-form">
        <input type="hidden" name="projectId" value="${escapeHtml(project.id)}" />
        <label class="field">
          <span>Project</span>
          <input name="name" value="${escapeHtml(project.name)}" placeholder="Implementation launch" required />
        </label>
        <label class="field">
          <span>Owner</span>
          <input name="owner" value="${escapeHtml(project.owner)}" list="owner-list" placeholder="Owner name" required />
        </label>
        <datalist id="owner-list">
          ${owners().map((owner) => `<option value="${escapeHtml(owner)}"></option>`).join("")}
        </datalist>
        <fieldset class="color-palette">
          <legend>Color</legend>
          ${renderColorPalette(project.color)}
        </fieldset>
        <div class="form-row">
          <label class="field">
            <span>Start</span>
            <input type="date" name="startDate" value="${escapeHtml(project.startDate)}" required />
          </label>
          <label class="field">
            <span>End</span>
            <input type="date" name="endDate" value="${escapeHtml(project.endDate)}" required />
          </label>
        </div>
        <fieldset class="segmented">
          <legend>Plan by</legend>
          <label>
            <input type="radio" name="mode" value="effort" ${project.mode === "effort" ? "checked" : ""} />
            <span>Total days</span>
          </label>
          <label>
            <input type="radio" name="mode" value="allocation" ${project.mode === "allocation" ? "checked" : ""} />
            <span>Daily %</span>
          </label>
        </fieldset>
        <div class="form-row">
          <label class="field">
            <span>Effort days</span>
            <input type="number" name="effortDays" min="0" step="0.25" value="${escapeHtml(project.mode === "effort" ? project.effortDays : math.effortDays)}" />
          </label>
          <label class="field">
            <span>Daily load %</span>
            <input type="number" name="dailyPercent" min="0" step="1" value="${escapeHtml(project.mode === "allocation" ? project.dailyPercent : math.dailyPercent)}" />
          </label>
        </div>
        <label class="field">
          <span>Dependencies</span>
          <select name="dependencies" multiple size="${Math.min(5, Math.max(2, state.data.projects.length))}" ${dependencyOptions ? "" : "disabled"}>
            ${dependencyOptions || "<option>No other projects</option>"}
          </select>
        </label>
        <label class="field">
          <span>Notes</span>
          <textarea name="notes" rows="3" placeholder="Scope, assumptions, handoff notes">${escapeHtml(project.notes)}</textarea>
        </label>
        <div class="computed-box" id="form-math">
          ${escapeHtml(math.duration)} calendar days · ${escapeHtml(math.workDays)} workdays · ${escapeHtml(math.effortDays)} effort days · ${escapeHtml(math.dailyPercent)}% per workday
        </div>
        <div class="form-actions">
          <button class="button primary" type="submit">
            <i data-lucide="check"></i>
            <span>${project.id ? "Update" : "Add"}</span>
          </button>
          ${project.id ? `
            <button class="button archive" type="button" data-action="${isArchived ? "restore-project" : "archive-project"}">
              <i data-lucide="${isArchived ? "archive-restore" : "archive"}"></i>
              <span>${isArchived ? "Restore" : "Complete"}</span>
            </button>
            <button class="button danger" type="button" data-action="delete-project">
              <i data-lucide="trash-2"></i>
              <span>Delete</span>
            </button>
          ` : ""}
        </div>
      </form>
    `;
  }

  function renderColorPalette(selectedColor) {
    const selected = normalizeColor(selectedColor);
    const colors = paletteColors();
    const customColumn = colors.includes(selected)
      ? ""
      : `
        <div class="palette-column" aria-label="Current">
          <label class="color-swatch" title="Current color">
            <input type="radio" name="color" value="${escapeHtml(selected)}" checked />
            <span style="background:${escapeHtml(selected)}"></span>
          </label>
        </div>
      `;
    return `${customColumn}${OFFICE_COLOR_PALETTE.map((group) => `
      <div class="palette-column" aria-label="${escapeHtml(group.name)}">
        ${group.colors.map((color, index) => {
          const normalized = normalizeColor(color);
          return `
            <label class="color-swatch" title="${escapeHtml(group.name)} ${index + 1}">
              <input type="radio" name="color" value="${escapeHtml(normalized)}" ${normalized === selected ? "checked" : ""} />
              <span style="background:${escapeHtml(normalized)}"></span>
            </label>
          `;
        }).join("")}
      </div>
    `).join("")}`;
  }

  function renderGanttChart(projects, days, totalWidth, dayWidth) {
    const rangeStart = days[0] || todayInput();
    const rows = projects.map((project) => renderGanttRow(project, rangeStart, totalWidth, dayWidth)).join("");
    return `
      <div class="gantt-scroll" aria-label="Gantt chart">
        <div class="gantt-table" style="--timeline-width:${totalWidth}px; --day-width:${dayWidth}px;">
          <div class="gantt-head">
            <div class="gantt-label head-label">Project</div>
            <div class="date-grid" style="width:${totalWidth}px;">
              ${days.map((day, index) => renderDayHead(day, index)).join("")}
            </div>
          </div>
          <div class="gantt-body">
            ${rows || renderEmptyChart(totalWidth)}
          </div>
        </div>
      </div>
    `;
  }

  function renderDayHead(day, index) {
    const date = dateFromInput(day);
    const showMonth = index === 0 || date.getDate() === 1;
    return `
      <div class="date-cell ${isWeekend(day) ? "weekend" : ""}">
        <span>${showMonth ? escapeHtml(formatMonth(day)) : ""}</span>
        <strong>${escapeHtml(date.getDate())}</strong>
      </div>
    `;
  }

  function renderGanttRow(project, rangeStart, totalWidth, dayWidth) {
    const math = getProjectMath(project);
    const offset = Math.max(0, diffDays(rangeStart, project.startDate));
    const left = offset * dayWidth;
    const width = Math.max(20, math.duration * dayWidth - 8);
    const dependencyNames = project.dependencies.map(projectName).join(", ");
    const selected = state.selectedProjectId === project.id ? "selected" : "";
    const readyStart = getDependencyReadyStart(project);
    const blocked = readyStart && compareDates(project.startDate, readyStart) < 0;
    const archived = project.status === "archived";
    const reorderTarget = state.rowDrag && state.rowDrag.beforeId === project.id ? "reorder-target" : "";
    const reorderDragging = state.rowDrag && state.rowDrag.projectId === project.id ? "reordering" : "";

    return `
      <div class="gantt-row ${selected} ${archived ? "archived" : ""} ${reorderTarget} ${reorderDragging}" data-row-id="${escapeHtml(project.id)}">
        <div class="gantt-label row-label">
          <button class="row-drag-handle" data-row-drag="${escapeHtml(project.id)}" title="Reorder project" aria-label="Reorder project">
            <i data-lucide="grip-vertical"></i>
            <span class="grip-dots" aria-hidden="true">⋮</span>
          </button>
          <button class="row-select" data-select-project="${escapeHtml(project.id)}">
          <span class="color-dot" style="background:${escapeHtml(project.color)}"></span>
          <span class="row-main">
            <strong>${escapeHtml(project.name)}</strong>
            <small>${escapeHtml(project.owner)} · ${escapeHtml(math.dailyPercent)}%/workday ${archived ? "· archived" : ""}</small>
          </span>
          </button>
        </div>
        <div class="row-track" style="width:${totalWidth}px;">
          <div
            class="gantt-bar ${selected} ${blocked ? "blocked" : ""} ${archived ? "archived" : ""}"
            data-project-id="${escapeHtml(project.id)}"
            data-drag-mode="move"
            role="button"
            tabindex="0"
            title="${escapeHtml(project.name)}: ${escapeHtml(formatLongDate(project.startDate))} to ${escapeHtml(formatLongDate(project.endDate))}${dependencyNames ? ` · after ${dependencyNames}` : ""}${archived ? " · archived" : ""}"
            style="left:${left}px; width:${width}px; --project-color:${escapeHtml(project.color)};"
          >
            <span class="resize-handle left" data-drag-mode="resize-start" title="Change start"></span>
            <span class="bar-title">${escapeHtml(project.name)}</span>
            <span class="bar-meta">${archived ? "Archived" : `${escapeHtml(math.effortDays)}d · ${escapeHtml(math.dailyPercent)}%`}</span>
            <span class="resize-handle right" data-drag-mode="resize-end" title="Change end"></span>
          </div>
        </div>
      </div>
    `;
  }

  function renderEmptyChart(totalWidth) {
    return `
      <div class="empty-row">
        <div class="gantt-label"></div>
        <div class="empty-track" style="width:${totalWidth}px;">
          <button class="button primary" data-action="new-project">
            <i data-lucide="plus"></i>
            <span>Add project</span>
          </button>
        </div>
      </div>
    `;
  }

  function renderUtilization(days, totalWidth, dayWidth) {
    const activeOwners = state.ownerFilter === "all" ? owners() : [state.ownerFilter];
    const rows = activeOwners.map((owner) => renderUtilizationRow(owner, days, totalWidth)).join("");
    const title = state.ownerFilter === "all" ? "Owner utilization" : `${state.ownerFilter} utilization`;
    const stats = state.ownerFilter === "all" ? "" : renderOwnerStats(state.ownerFilter, days);
    return `
      <section class="util-panel">
        <div class="panel-header compact">
          <div>
            <h2 class="panel-title">${escapeHtml(title)}</h2>
            <p class="panel-subtitle">${escapeHtml(days.length)} calendar days · weekends = 0% · full day = 100%</p>
          </div>
          ${stats}
        </div>
        <div class="util-scroll">
          <div class="util-table" style="--timeline-width:${totalWidth}px; --day-width:${dayWidth}px;">
            ${rows || `<div class="empty-util">No owner load in this view.</div>`}
          </div>
        </div>
      </section>
    `;
  }

  function renderOwnerStats(owner, days) {
    const stats = getLoadStats(owner, days);
    return `
      <div class="stat-cluster">
        <span>Avg ${escapeHtml(stats.average)}%</span>
        <span>Peak ${escapeHtml(roundTo(stats.max.load, 1))}%${stats.max.day ? ` ${escapeHtml(formatShortDate(stats.max.day))}` : ""}</span>
      </div>
    `;
  }

  function renderUtilizationRow(owner, days) {
    const stats = getLoadStats(owner, days);
    return `
      <div class="util-row">
        <div class="util-owner">
          <strong>${escapeHtml(owner)}</strong>
          <small>Peak ${escapeHtml(roundTo(stats.max.load, 1))}%</small>
        </div>
        <div class="util-cells">
          ${stats.loads.map((item) => renderLoadCell(item.day, item.load)).join("")}
        </div>
      </div>
    `;
  }

  function renderLoadCell(day, load) {
    const level = Math.min(100, Math.max(0, load));
    return `
      <div
        class="load-cell ${loadClass(load)} ${isWeekend(day) ? "weekend" : ""}"
        title="${escapeHtml(formatLongDate(day))}: ${escapeHtml(roundTo(load, 1))}%"
      >
        <span style="height:${level}%"></span>
        <em>${load > 0 ? escapeHtml(roundTo(load, 0)) : ""}</em>
      </div>
    `;
  }

  function renderProjectsView() {
    const projects = sortedProjects();
    return `
      <section class="view active">
        <div class="table-panel">
          <div class="panel-header">
            <div>
              <h2 class="panel-title">Projects</h2>
              <p class="panel-subtitle">${escapeHtml(projects.length)} scheduled items</p>
            </div>
            <button class="button primary" data-action="new-project">
              <i data-lucide="plus"></i>
              <span>Add project</span>
            </button>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Project</th>
                  <th>Owner</th>
                  <th>Status</th>
                  <th>Dates</th>
                  <th>Effort</th>
                  <th>Daily</th>
                  <th>Dependencies</th>
                </tr>
              </thead>
              <tbody>
                ${projects.map(renderProjectRow).join("") || `
                  <tr>
                    <td colspan="7" class="empty-table">No projects yet.</td>
                  </tr>
                `}
              </tbody>
            </table>
          </div>
        </div>
      </section>
    `;
  }

  function renderProjectRow(project) {
    const math = getProjectMath(project);
    const archived = project.status === "archived";
    return `
      <tr class="${archived ? "archived-row" : ""}" data-select-project="${escapeHtml(project.id)}">
        <td>
          <span class="table-title">
            <span class="color-dot" style="background:${escapeHtml(project.color)}"></span>
            ${escapeHtml(project.name)}
          </span>
        </td>
        <td>${escapeHtml(project.owner)}</td>
        <td><span class="status-pill ${archived ? "archived" : "active"}">${archived ? "Archived" : "Active"}</span></td>
        <td>${escapeHtml(formatShortDate(project.startDate))} - ${escapeHtml(formatShortDate(project.endDate))}</td>
        <td>${escapeHtml(math.effortDays)} days</td>
        <td>${escapeHtml(math.dailyPercent)}%</td>
        <td>${escapeHtml(project.dependencies.map(projectName).join(", ") || "None")}</td>
      </tr>
    `;
  }

  function renderSettingsView() {
    return `
      <section class="view active">
        <div class="settings-grid">
          <section class="settings-panel">
            <div class="panel-header">
              <div>
                <h2 class="panel-title">Google Drive</h2>
                <p class="panel-subtitle">${escapeHtml(state.drive.modifiedTime ? `Drive file ${formatDateTime(state.drive.modifiedTime)}` : "appDataFolder")}</p>
              </div>
              ${renderSyncPill()}
            </div>
            <div class="panel-body">
              <div class="drive-actions">
                <button class="button primary" data-action="connect-drive">
                  <i data-lucide="cloud"></i>
                  <span>${escapeHtml(driveButtonLabel())}</span>
                </button>
                <button class="button" data-action="save-drive">
                  <i data-lucide="save"></i>
                  <span>Save</span>
                </button>
                <button class="button" data-action="pull-drive">
                  <i data-lucide="download-cloud"></i>
                  <span>Pull</span>
                </button>
              </div>
              ${state.drive.error ? `<p class="error-text">${escapeHtml(state.drive.error)}</p>` : ""}
              <p class="settings-note">Data file: ${escapeHtml(DRIVE_FILE_NAME)}</p>
            </div>
          </section>
          <section class="settings-panel">
            <div class="panel-header">
              <div>
                <h2 class="panel-title">OAuth client</h2>
                <p class="panel-subtitle">Same client ID key as Project Tracker</p>
              </div>
            </div>
            <form class="panel-body form-grid" id="client-form">
              <label class="field">
                <span>Client ID</span>
                <input name="clientId" value="${escapeHtml(getClientId())}" placeholder="000000000000-example.apps.googleusercontent.com" />
              </label>
              <button class="button primary" type="submit">
                <i data-lucide="check"></i>
                <span>Save client</span>
              </button>
            </form>
          </section>
        </div>
      </section>
    `;
  }

  function handleSubmit(event) {
    if (event.target.id === "project-form") {
      event.preventDefault();
      saveProjectFromForm(event.target);
    }
    if (event.target.id === "client-form") {
      event.preventDefault();
      saveClientId(event);
    }
  }

  function saveProjectFromForm(form) {
    const formData = new FormData(form);
    const projectId = String(formData.get("projectId") || "").trim();
    const isExisting = Boolean(projectId);
    const project = isExisting ? projectById(projectId) : normalizeProject({}, state.data.projects.length);
    if (!project) return;

    const dependencies = formData.getAll("dependencies").map(String).filter(Boolean);
    if (wouldCreateCycle(project.id, dependencies)) {
      showToast("Dependency cycle blocked.");
      return;
    }

    project.name = String(formData.get("name") || "Untitled project").trim() || "Untitled project";
    project.owner = String(formData.get("owner") || "Unassigned").trim() || "Unassigned";
    project.startDate = String(formData.get("startDate") || todayInput());
    project.endDate = String(formData.get("endDate") || project.startDate);
    project.mode = formData.get("mode") === "allocation" ? "allocation" : "effort";
    project.effortDays = Math.max(0, Number(formData.get("effortDays") || 0));
    project.dailyPercent = Math.max(0, Number(formData.get("dailyPercent") || 0));
    project.dependencies = dependencies;
    project.notes = String(formData.get("notes") || "").trim();
    project.color = normalizeColor(formData.get("color") || project.color, state.data.projects.length);
    project.updatedAt = new Date().toISOString();
    sanitizeProjectDates(project);
    if (project.mode === "allocation") {
      project.effortDays = roundTo((workDaysInclusive(project.startDate, project.endDate) * project.dailyPercent) / 100, 2);
    } else {
      const workDays = workDaysInclusive(project.startDate, project.endDate);
      project.dailyPercent = roundTo(workDays ? (project.effortDays / workDays) * 100 : 0, 2);
    }

    if (!isExisting) {
      project.createdAt = new Date().toISOString();
      project.status = "active";
      project.completedAt = "";
      project.order = nextProjectOrder();
      state.data.projects.push(project);
    }

    const shifted = applyDependencyCascade();
    state.selectedProjectId = isExisting ? project.id : "";
    saveLocal();
    showToast(shifted.length ? `Saved. Shifted ${shifted.length} dependent project${shifted.length === 1 ? "" : "s"}.` : isExisting ? "Project saved." : "Project added.");
  }

  function handleClick(event) {
    const tab = event.target.closest("[data-tab]");
    if (tab) {
      state.activeTab = tab.dataset.tab;
      render();
      return;
    }

    const projectTarget = event.target.closest("[data-select-project]");
    if (projectTarget) {
      const projectId = projectTarget.dataset.selectProject;
      state.selectedProjectId = projectId;
      state.activeTab = "gantt";
      render();
      return;
    }

    const actionTarget = event.target.closest("[data-action]");
    if (!actionTarget) return;
    const action = actionTarget.dataset.action;
    handleAction(action);
  }

  function handleAction(action) {
    if (action === "new-project") {
      state.selectedProjectId = "";
      state.activeTab = "gantt";
      render();
    }
    if (action === "delete-project") deleteSelectedProject();
    if (action === "archive-project") setSelectedProjectStatus("archived");
    if (action === "restore-project") setSelectedProjectStatus("active");
    if (action === "zoom-in") setDayWidth(state.data.settings.dayWidth + 6);
    if (action === "zoom-out") setDayWidth(state.data.settings.dayWidth - 6);
    if (action === "connect-drive") connectDrive({ afterAuth: state.drive.pendingAction || "pull" });
    if (action === "save-drive") saveToDrive({ manual: true });
    if (action === "pull-drive") loadFromDrive({ manual: true });
  }

  function setDayWidth(value) {
    state.data.settings.dayWidth = clampNumber(value, MIN_DAY_WIDTH, MAX_DAY_WIDTH);
    saveLocal();
  }

  function deleteSelectedProject() {
    const project = projectById(state.selectedProjectId);
    if (!project) return;
    const confirmed = window.confirm(`Delete ${project.name}? Dependent projects will keep their dates but lose this dependency.`);
    if (!confirmed) return;
    state.data.projects = state.data.projects
      .filter((candidate) => candidate.id !== project.id)
      .map((candidate) => ({
        ...candidate,
        dependencies: candidate.dependencies.filter((dependencyId) => dependencyId !== project.id)
      }));
    normalizeProjectOrder();
    state.selectedProjectId = "";
    saveLocal();
    showToast("Project deleted.");
  }

  function setSelectedProjectStatus(status) {
    const project = projectById(state.selectedProjectId);
    if (!project) return;
    project.status = status === "archived" ? "archived" : "active";
    project.completedAt = project.status === "archived" ? new Date().toISOString() : "";
    project.updatedAt = new Date().toISOString();
    saveLocal();
    showToast(project.status === "archived" ? "Project archived." : "Project restored.");
  }

  function handleChange(event) {
    if (event.target.id === "owner-filter") {
      state.ownerFilter = event.target.value;
      render();
      return;
    }

    if (event.target.id === "show-archived-toggle") {
      state.data.settings.showArchivedOnTimeline = event.target.checked;
      saveLocal();
      showToast(event.target.checked ? "Completed projects shown." : "Completed projects hidden from timeline.");
      return;
    }

    if (event.target.closest("#project-form")) {
      updateFormMath(event.target.closest("#project-form"));
    }
  }

  function handleInput(event) {
    if (event.target.closest("#project-form")) {
      updateFormMath(event.target.closest("#project-form"));
    }
  }

  function updateFormMath(form) {
    if (!form) return;
    const output = form.querySelector("#form-math");
    if (!output) return;
    const formData = new FormData(form);
    const startDate = String(formData.get("startDate") || todayInput());
    const rawEndDate = String(formData.get("endDate") || startDate);
    const endDate = compareDates(rawEndDate, startDate) < 0 ? startDate : rawEndDate;
    const duration = diffDaysInclusive(startDate, endDate);
    const workDays = workDaysInclusive(startDate, endDate);
    const mode = formData.get("mode") === "allocation" ? "allocation" : "effort";
    const effortInput = form.querySelector("[name='effortDays']");
    const dailyInput = form.querySelector("[name='dailyPercent']");

    if (mode === "allocation") {
      const dailyPercent = Math.max(0, Number(formData.get("dailyPercent") || 0));
      const effortDays = roundTo((workDays * dailyPercent) / 100, 2);
      if (effortInput) effortInput.value = effortDays;
      if (dailyInput) dailyInput.disabled = false;
      if (effortInput) effortInput.disabled = true;
      output.textContent = `${duration} calendar days · ${workDays} workdays · ${effortDays} effort days · ${roundTo(dailyPercent, 1)}% per workday`;
      return;
    }

    const effortDays = Math.max(0, Number(formData.get("effortDays") || 0));
    const dailyPercent = roundTo(workDays ? (effortDays / workDays) * 100 : 0, 1);
    if (dailyInput) dailyInput.value = dailyPercent;
    if (effortInput) effortInput.disabled = false;
    if (dailyInput) dailyInput.disabled = true;
    output.textContent = `${duration} calendar days · ${workDays} workdays · ${roundTo(effortDays, 2)} effort days · ${dailyPercent}% per workday`;
  }

  function getRowDropBeforeId(clientY) {
    const rows = [...document.querySelectorAll(".gantt-row[data-row-id]")];
    for (const row of rows) {
      const bounds = row.getBoundingClientRect();
      if (clientY < bounds.top + bounds.height / 2) return row.dataset.rowId || "";
    }
    return "";
  }

  function reorderProjectBefore(projectId, beforeId) {
    if (beforeId === projectId) return false;
    const orderedIds = sortedProjects(state.data.projects).map((project) => project.id);
    const currentIndex = orderedIds.indexOf(projectId);
    if (currentIndex < 0) return false;

    orderedIds.splice(currentIndex, 1);
    const insertIndex = beforeId && beforeId !== projectId ? orderedIds.indexOf(beforeId) : orderedIds.length;
    orderedIds.splice(insertIndex < 0 ? orderedIds.length : insertIndex, 0, projectId);

    const byId = new Map(state.data.projects.map((project) => [project.id, project]));
    orderedIds.forEach((id, index) => {
      const project = byId.get(id);
      if (project) {
        project.order = index;
        project.updatedAt = new Date().toISOString();
      }
    });
    return true;
  }

  function updateRowDragTarget(clientY, moved) {
    if (!state.rowDrag) return;
    const beforeId = getRowDropBeforeId(clientY);
    const nextMoved = Boolean(moved);
    if (beforeId !== state.rowDrag.beforeId || nextMoved !== state.rowDrag.moved) {
      state.rowDrag.beforeId = beforeId;
      state.rowDrag.moved = nextMoved;
      render();
    }
  }

  function getRowAutoScrollVelocity(clientY) {
    const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
    if (!viewportHeight) return 0;
    if (clientY < ROW_AUTOSCROLL_EDGE) {
      return -Math.ceil(((ROW_AUTOSCROLL_EDGE - clientY) / ROW_AUTOSCROLL_EDGE) * ROW_AUTOSCROLL_MAX_SPEED);
    }
    if (clientY > viewportHeight - ROW_AUTOSCROLL_EDGE) {
      return Math.ceil(((clientY - (viewportHeight - ROW_AUTOSCROLL_EDGE)) / ROW_AUTOSCROLL_EDGE) * ROW_AUTOSCROLL_MAX_SPEED);
    }
    return 0;
  }

  function startRowAutoScroll() {
    stopRowAutoScroll();
    const tick = () => {
      if (!state.rowDrag) {
        rowAutoScrollFrame = 0;
        return;
      }

      const velocity = state.rowDrag.moved ? getRowAutoScrollVelocity(state.rowDrag.clientY) : 0;
      if (velocity) {
        const previousScrollY = window.scrollY;
        window.scrollBy(0, velocity);
        if (window.scrollY !== previousScrollY) {
          updateRowDragTarget(state.rowDrag.clientY, true);
        }
      }

      rowAutoScrollFrame = window.requestAnimationFrame(tick);
    };
    rowAutoScrollFrame = window.requestAnimationFrame(tick);
  }

  function stopRowAutoScroll() {
    if (!rowAutoScrollFrame) return;
    window.cancelAnimationFrame(rowAutoScrollFrame);
    rowAutoScrollFrame = 0;
  }

  function handlePointerDown(event) {
    const rowDragTarget = event.target.closest("[data-row-drag]");
    if (rowDragTarget) {
      const project = projectById(rowDragTarget.dataset.rowDrag);
      if (!project) return;
      event.preventDefault();
      state.selectedProjectId = project.id;
      state.rowDrag = {
        projectId: project.id,
        startY: event.clientY,
        clientY: event.clientY,
        beforeId: project.id,
        moved: false
      };
      document.body.classList.add("row-dragging");
      startRowAutoScroll();
      render();
      return;
    }

    const dragTarget = event.target.closest("[data-drag-mode]");
    if (!dragTarget) return;
    const bar = dragTarget.closest(".gantt-bar");
    if (!bar) return;
    const project = projectById(bar.dataset.projectId);
    if (!project) return;

    event.preventDefault();
    state.selectedProjectId = project.id;
    state.drag = {
      projectId: project.id,
      mode: dragTarget.dataset.dragMode,
      startX: event.clientX,
      initialStart: project.startDate,
      initialEnd: project.endDate,
      lastDeltaDays: 0,
      moved: false
    };
    document.body.classList.add("dragging");
    render();
  }

  function handlePointerMove(event) {
    if (state.rowDrag) {
      state.rowDrag.clientY = event.clientY;
      const moved = Math.abs(event.clientY - state.rowDrag.startY) > 3;
      updateRowDragTarget(event.clientY, moved);
      return;
    }

    if (!state.drag) return;
    const dayWidth = state.data.settings.dayWidth;
    const deltaDays = Math.round((event.clientX - state.drag.startX) / dayWidth);
    if (deltaDays === state.drag.lastDeltaDays) return;

    const project = projectById(state.drag.projectId);
    if (!project) return;

    let nextStart = state.drag.initialStart;
    let nextEnd = state.drag.initialEnd;
    if (state.drag.mode === "move") {
      nextStart = addDays(state.drag.initialStart, deltaDays);
      nextEnd = addDays(state.drag.initialEnd, deltaDays);
    } else if (state.drag.mode === "resize-start") {
      nextStart = addDays(state.drag.initialStart, deltaDays);
      if (compareDates(nextStart, nextEnd) > 0) nextStart = nextEnd;
    } else if (state.drag.mode === "resize-end") {
      nextEnd = addDays(state.drag.initialEnd, deltaDays);
      if (compareDates(nextEnd, nextStart) < 0) nextEnd = nextStart;
    }

    project.startDate = nextStart;
    project.endDate = nextEnd;
    project.updatedAt = new Date().toISOString();
    applyDependencyCascade();
    state.drag.lastDeltaDays = deltaDays;
    state.drag.moved = true;
    render();
  }

  function handlePointerUp() {
    if (state.rowDrag) {
      const { projectId, beforeId, moved } = state.rowDrag;
      state.rowDrag = null;
      stopRowAutoScroll();
      document.body.classList.remove("row-dragging");
      if (moved && reorderProjectBefore(projectId, beforeId)) {
        saveLocal();
        showToast("Project order updated.");
      } else {
        render();
      }
      return;
    }

    if (!state.drag) return;
    const moved = state.drag.moved;
    state.drag = null;
    document.body.classList.remove("dragging");
    if (moved) {
      saveLocal();
      showToast("Schedule updated.");
    } else {
      render();
    }
  }

  function handlePointerCancel() {
    if (state.rowDrag) {
      state.rowDrag = null;
      stopRowAutoScroll();
      document.body.classList.remove("row-dragging");
      render();
    }
    if (state.drag) {
      state.drag = null;
      document.body.classList.remove("dragging");
      render();
    }
  }

  function saveClientId(event) {
    const form = new FormData(event.currentTarget);
    const clientId = String(form.get("clientId") || "").trim();
    localStorage.setItem(CLIENT_STORAGE_KEY, clientId);
    state.tokenClient = null;
    showToast("Client ID saved.");
    render();
  }

  function getClientId() {
    return String(window.SCHEDULE_GOOGLE_CLIENT_ID || window.PM_GOOGLE_CLIENT_ID || localStorage.getItem(CLIENT_STORAGE_KEY) || "").trim();
  }

  function ensureTokenClient() {
    const clientId = getClientId();
    if (!clientId) {
      state.activeTab = "settings";
      render();
      showToast("Add a Google OAuth client ID in Settings.");
      return false;
    }
    if (!window.google || !window.google.accounts || !window.google.accounts.oauth2) {
      showToast("Google Identity Services is not loaded.");
      return false;
    }
    if (!state.tokenClient) {
      state.tokenClient = window.google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: DRIVE_SCOPE,
        callback: async (response) => {
          if (response.error) {
            state.drive.error = response.error;
            state.drive.pending = false;
            state.drive.message = "Drive auth failed";
            render();
            return;
          }
          enableDriveSync();
          state.drive.accessToken = response.access_token;
          state.drive.tokenExpiresAt = Date.now() + Number(response.expires_in || 3600) * 1000;
          state.drive.connected = true;
          state.drive.pending = false;
          state.drive.error = "";
          state.drive.message = "Drive connected";
          scheduleTokenExpiryNotice();
          render();

          const action = state.drive.pendingAction || "pull";
          state.drive.pendingAction = "";
          if (action === "save") {
            await saveToDrive({ fromAuth: true });
          } else if (action === "pull") {
            await loadFromDrive({ fromAuth: true });
          }
        }
      });
    }
    return true;
  }

  function connectDrive({ afterAuth = "pull" } = {}) {
    if (!ensureTokenClient()) return;
    const hadDriveGrant = state.drive.enabled;
    enableDriveSync();
    state.drive.pending = true;
    state.drive.pendingAction = afterAuth;
    state.drive.message = "Waiting for Google";
    render();
    state.tokenClient.requestAccessToken({
      prompt: hadDriveGrant ? "" : "consent"
    });
  }

  function scheduleDriveSave() {
    window.clearTimeout(state.saveTimer);
    state.drive.pending = true;
    state.drive.message = "Save queued";
    state.saveTimer = window.setTimeout(() => saveToDrive(), 900);
  }

  async function driveFetch(url, options = {}) {
    if (!hasValidAccessToken()) {
      markDriveNeedsAuth(state.drive.pendingAction || "save", "Reconnect Drive");
      throw makeAuthError("Google Drive needs a fresh access token.");
    }
    const response = await fetch(url, {
      ...options,
      headers: {
        ...(options.headers || {}),
        Authorization: `Bearer ${state.drive.accessToken}`
      }
    });
    if (response.status === 401) {
      markDriveNeedsAuth(state.drive.pendingAction || "save", "Reconnect Drive");
      render();
      throw makeAuthError("Google Drive token expired.");
    }
    if (!response.ok) {
      const text = await response.text();
      throw new Error(text || `Drive request failed with ${response.status}`);
    }
    return response;
  }

  async function findDriveFile() {
    const query = encodeURIComponent(`name='${DRIVE_FILE_NAME}' and trashed=false`);
    const url = `https://www.googleapis.com/drive/v3/files?spaces=appDataFolder&q=${query}&fields=files(id,name,modifiedTime)`;
    const response = await driveFetch(url);
    const body = await response.json();
    const file = Array.isArray(body.files) ? body.files[0] : null;
    if (file) {
      state.drive.fileId = file.id;
      state.drive.modifiedTime = file.modifiedTime || "";
    }
    return file;
  }

  async function loadFromDrive({ manual = false } = {}) {
    if (!hasValidAccessToken()) {
      if (manual) {
        connectDrive({ afterAuth: "pull" });
      } else {
        markDriveNeedsAuth("pull", "Reconnect to pull");
        render();
      }
      return;
    }
    try {
      state.drive.pending = true;
      state.drive.pendingAction = "pull";
      state.drive.message = "Pulling Drive";
      render();

      const file = await findDriveFile();
      if (!file) {
        state.drive.message = "No Drive file";
        state.drive.pending = false;
        render();
        if (hasUserData(state.data)) await saveToDrive();
        return;
      }

      const response = await driveFetch(`https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`);
      const cloudData = normalizeData(await response.json());
      const localData = state.data;
      const cloudTime = new Date(cloudData.updatedAt).getTime();
      const localTime = new Date(localData.updatedAt).getTime();

      if (hasUserData(localData) && localTime > cloudTime && localData.updatedAt !== cloudData.updatedAt) {
        const keepLocal = window.confirm("Local edits are newer than the Drive copy. Keep local changes and save them to Drive?");
        if (keepLocal) {
          state.drive.pending = false;
          await saveToDrive();
          return;
        }
      }

      state.data = cloudData;
      localStorage.setItem(APP_STORAGE_KEY, JSON.stringify(state.data));
      state.drive.modifiedTime = file.modifiedTime || "";
      state.drive.pending = false;
      state.drive.pendingAction = "";
      state.drive.message = "Drive synced";
      render();
      showToast("Loaded from Drive.");
    } catch (error) {
      if (isAuthError(error)) {
        markDriveNeedsAuth("pull", "Reconnect to pull");
        state.drive.error = "";
        render();
        showToast("Reconnect Drive to pull updates.");
        return;
      }
      state.drive.pending = false;
      state.drive.error = String(error.message || error);
      state.drive.message = "Drive error";
      render();
      showToast("Drive sync failed.");
      console.error(error);
    }
  }

  async function saveToDrive({ manual = false } = {}) {
    if (!hasValidAccessToken()) {
      if (manual) {
        connectDrive({ afterAuth: "save" });
      } else {
        markDriveNeedsAuth("save", "Reconnect to save");
        render();
      }
      return;
    }
    try {
      state.drive.pending = true;
      state.drive.pendingAction = "save";
      state.drive.message = "Saving Drive";
      render();

      if (!state.drive.fileId) await findDriveFile();

      const json = JSON.stringify(state.data, null, 2);
      let result;
      if (state.drive.fileId) {
        const response = await driveFetch(`https://www.googleapis.com/upload/drive/v3/files/${state.drive.fileId}?uploadType=media&fields=id,modifiedTime`, {
          method: "PATCH",
          headers: {
            "Content-Type": "application/json"
          },
          body: json
        });
        result = await response.json();
      } else {
        result = await createDriveFile(json);
      }

      state.drive.fileId = result.id || state.drive.fileId;
      state.drive.modifiedTime = result.modifiedTime || state.drive.modifiedTime;
      state.drive.pending = false;
      state.drive.pendingAction = "";
      state.drive.error = "";
      state.drive.message = "Drive saved";
      render();
      if (manual) showToast("Saved to Drive.");
    } catch (error) {
      if (isAuthError(error)) {
        markDriveNeedsAuth("save", "Reconnect to save");
        state.drive.error = "";
        render();
        showToast("Reconnect Drive to finish saving.");
        return;
      }
      state.drive.pending = false;
      state.drive.error = String(error.message || error);
      state.drive.message = "Drive error";
      render();
      showToast("Drive save failed.");
      console.error(error);
    }
  }

  async function createDriveFile(json) {
    const boundary = `schedule-gantt-${Date.now()}`;
    const metadata = {
      name: DRIVE_FILE_NAME,
      parents: ["appDataFolder"],
      mimeType: "application/json"
    };
    const body = [
      `--${boundary}`,
      "Content-Type: application/json; charset=UTF-8",
      "",
      JSON.stringify(metadata),
      `--${boundary}`,
      "Content-Type: application/json; charset=UTF-8",
      "",
      json,
      `--${boundary}--`
    ].join("\r\n");

    const response = await driveFetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,modifiedTime", {
      method: "POST",
      headers: {
        "Content-Type": `multipart/related; boundary=${boundary}`
      },
      body
    });
    return response.json();
  }

  function hasUserData(data) {
    return Boolean(data && Array.isArray(data.projects) && data.projects.length);
  }

  function showToast(message) {
    window.clearTimeout(state.toastTimer);
    const toast = document.getElementById("toast");
    if (!toast) return;
    toast.textContent = message;
    toast.classList.add("show");
    state.toastTimer = window.setTimeout(() => {
      toast.classList.remove("show");
    }, 2400);
  }

  app.addEventListener("click", handleClick);
  app.addEventListener("submit", handleSubmit);
  app.addEventListener("change", handleChange);
  app.addEventListener("input", handleInput);
  app.addEventListener("pointerdown", handlePointerDown);
  window.addEventListener("pointermove", handlePointerMove);
  window.addEventListener("pointerup", handlePointerUp);
  window.addEventListener("pointercancel", handlePointerCancel);

  render();
})();
