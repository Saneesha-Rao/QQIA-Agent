/**
 * QQIA Agent — Analytics Dashboard
 * Provides burndown charts, workstream health, timeline, and change feed.
 * Loaded as a separate module from the main chat UI.
 */
(function() {
  'use strict';

  const API_BASE = '';
  let currentTrack = 'Corp';
  let filterWorkstream = null; // click-to-filter state
  let refreshTimer = null;

  // ---- Init ----
  window.initAnalytics = function() {
    const container = document.getElementById('analytics-container');
    if (!container) return;
    container.innerHTML = getAnalyticsHTML();
    setupTrackToggle();
    loadAll();
    startAutoRefresh();
  };

  // ---- Auto-refresh (5 min) ----
  function startAutoRefresh() {
    if (refreshTimer) clearInterval(refreshTimer);
    refreshTimer = setInterval(() => {
      const container = document.getElementById('analytics-container');
      if (container && container.style.display !== 'none') {
        loadAll();
        updateRefreshIndicator();
        if (window.showToast) window.showToast('Dashboard auto-refreshed', 'info');
      }
    }, 5 * 60 * 1000);
  }

  function updateRefreshIndicator() {
    const el = document.getElementById('last-refresh');
    if (el) el.textContent = 'Last refresh: ' + new Date().toLocaleTimeString();
  }

  function setupTrackToggle() {
    document.querySelectorAll('.track-toggle-btn').forEach(btn => {
      btn.addEventListener('click', function() {
        document.querySelectorAll('.track-toggle-btn').forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        currentTrack = this.dataset.track;
        filterWorkstream = null;
        loadAll();
      });
    });
  }

  function loadAll() {
    loadOverallHealth();
    loadWorkstreamHealth();
    loadTimeline();
    loadChanges();
    updateRefreshIndicator();
  }

  // ---- Filter helpers ----
  function setWorkstreamFilter(ws) {
    if (filterWorkstream === ws) {
      filterWorkstream = null; // toggle off
      if (window.showToast) window.showToast('Filter cleared', 'info');
    } else {
      filterWorkstream = ws;
      if (window.showToast) window.showToast('Filtered to: ' + ws, 'info');
    }
    // Re-render timeline with filter
    loadTimeline();
    // Update health card highlights
    document.querySelectorAll('.health-card').forEach(card => {
      const name = card.dataset.workstream;
      if (filterWorkstream && name !== filterWorkstream) {
        card.style.opacity = '0.4';
      } else {
        card.style.opacity = '1';
      }
    });
    // Update filter indicator
    const ind = document.getElementById('filter-indicator');
    if (ind) {
      if (filterWorkstream) {
        ind.innerHTML = `<span style="background:#ede9fe;color:#7c3aed;padding:4px 12px;border-radius:12px;font-size:0.8rem;">Filtered: <strong>${filterWorkstream}</strong> <span style="cursor:pointer;margin-left:4px;" onclick="window._clearFilter()">✕</span></span>`;
      } else {
        ind.innerHTML = '';
      }
    }
  }
  window._setWsFilter = setWorkstreamFilter;
  window._clearFilter = function() { setWorkstreamFilter(null); };

  // ---- Overall Health ----
  async function loadOverallHealth() {
    const el = document.getElementById('overall-health');
    if (!el) return;
    el.innerHTML = '<div class="loading">Loading...</div>';
    try {
      const res = await fetch(`${API_BASE}/api/analytics/workstream-health?track=${currentTrack}`);
      const data = await res.json();
      el.innerHTML = renderOverallHealth(data.workstreams);
    } catch (e) {
      el.innerHTML = '<div class="error">Failed to load overall health</div>';
    }
  }

  function renderOverallHealth(workstreams) {
    if (!workstreams || workstreams.length === 0) return '<div class="empty">No data</div>';

    let completed = 0, inProgress = 0, blocked = 0, notStarted = 0, overdue = 0, total = 0;
    workstreams.forEach(ws => {
      completed += ws.completed;
      inProgress += ws.inProgress;
      blocked += ws.blocked;
      notStarted += ws.notStarted;
      overdue += ws.overdue;
      total += ws.total;
    });
    const pct = total > 0 ? Math.round(completed / total * 100) : 0;
    const greenCount = workstreams.filter(w => w.health === 'green').length;
    const yellowCount = workstreams.filter(w => w.health === 'yellow').length;
    const redCount = workstreams.filter(w => w.health === 'red').length;

    let overallColor, overallEmoji, overallLabel;
    if (redCount > 3 || blocked > 5) {
      overallColor = '#ef4444'; overallEmoji = '🔴'; overallLabel = 'Needs Attention';
    } else if (redCount > 0 || yellowCount > 2) {
      overallColor = '#f59e0b'; overallEmoji = '🟡'; overallLabel = 'Some Risks';
    } else {
      overallColor = '#10b981'; overallEmoji = '🟢'; overallLabel = 'On Track';
    }

    return `
      <div style="background:white;border-radius:12px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,0.08);border-top:4px solid ${overallColor};">
        <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap;">
          <div style="flex:1;min-width:200px;">
            <div style="font-size:2.5rem;font-weight:700;color:#1f2937;">${pct}%</div>
            <div style="font-size:1rem;color:${overallColor};font-weight:600;">${overallEmoji} ${overallLabel}</div>
            <div style="font-size:0.8rem;color:#6b7280;margin-top:4px;">${currentTrack} Track · ${total} total steps</div>
          </div>
          <div style="flex:2;min-width:250px;">
            <div style="display:flex;height:28px;border-radius:6px;overflow:hidden;background:#f3f4f6;">
              <div style="width:${total?completed/total*100:0}%;background:#10b981;" title="Completed: ${completed}"></div>
              <div style="width:${total?inProgress/total*100:0}%;background:#3b82f6;" title="In Progress: ${inProgress}"></div>
              <div style="width:${total?blocked/total*100:0}%;background:#ef4444;" title="Blocked: ${blocked}"></div>
              <div style="width:${total?notStarted/total*100:0}%;background:#d1d5db;" title="Not Started: ${notStarted}"></div>
            </div>
            <div style="display:flex;gap:16px;margin-top:8px;font-size:0.8rem;color:#4b5563;flex-wrap:wrap;">
              <span>✅ Completed: <strong>${completed}</strong></span>
              <span>🔄 In Progress: <strong>${inProgress}</strong></span>
              <span>⛔ Blocked: <strong>${blocked}</strong></span>
              <span>⬜ Not Started: <strong>${notStarted}</strong></span>
              <span>⏰ Overdue: <strong>${overdue}</strong></span>
            </div>
          </div>
          <div style="min-width:160px;text-align:center;">
            <div style="font-size:0.8rem;color:#6b7280;margin-bottom:6px;">Workstream Health</div>
            <div style="display:flex;gap:12px;justify-content:center;">
              <div style="text-align:center;"><div style="font-size:1.5rem;font-weight:700;color:#10b981;">${greenCount}</div><div style="font-size:0.7rem;color:#6b7280;">🟢 Good</div></div>
              <div style="text-align:center;"><div style="font-size:1.5rem;font-weight:700;color:#f59e0b;">${yellowCount}</div><div style="font-size:0.7rem;color:#6b7280;">🟡 Risk</div></div>
              <div style="text-align:center;"><div style="font-size:1.5rem;font-weight:700;color:#ef4444;">${redCount}</div><div style="font-size:0.7rem;color:#6b7280;">🔴 Critical</div></div>
            </div>
          </div>
        </div>
      </div>`;
  }

  // ---- Workstream Health ----
  async function loadWorkstreamHealth() {
    const el = document.getElementById('health-cards');
    if (!el) return;
    el.innerHTML = '<div class="loading">Loading...</div>';
    try {
      const res = await fetch(`${API_BASE}/api/analytics/workstream-health?track=${currentTrack}`);
      const data = await res.json();
      el.innerHTML = renderHealthCards(data.workstreams);
    } catch (e) {
      el.innerHTML = '<div class="error">Failed to load workstream health</div>';
    }
  }

  function renderHealthCards(workstreams) {
    if (!workstreams || workstreams.length === 0) return '<div class="empty">No data</div>';
    const emoji = { green: '🟢', yellow: '🟡', red: '🔴' };
    const colors = { green: '#10b981', yellow: '#f59e0b', red: '#ef4444' };
    return '<div class="health-grid">' + workstreams.map(ws => {
      const isActive = filterWorkstream === ws.workstream;
      const opacity = filterWorkstream && !isActive ? '0.4' : '1';
      const outline = isActive ? 'outline:2px solid #7c3aed;outline-offset:2px;' : '';
      return `
      <div class="health-card health-${ws.health}" data-workstream="${ws.workstream}" 
           onclick="window._setWsFilter('${ws.workstream.replace(/'/g, "\\'")}')"
           style="border-left:4px solid ${colors[ws.health]};cursor:pointer;opacity:${opacity};${outline}transition:all 0.2s;">
        <div class="health-header">
          <span class="health-emoji">${emoji[ws.health]}</span>
          <span class="health-name">${ws.workstream}</span>
        </div>
        <div class="health-pct">${ws.completionPct}%</div>
        <div class="health-bar">
          <div class="health-bar-fill" style="width:${ws.completionPct}%;background:${colors[ws.health]}"></div>
        </div>
        <div class="health-stats">
          <span title="Completed">✅ ${ws.completed}</span>
          <span title="In Progress">🔄 ${ws.inProgress}</span>
          <span title="Blocked">⛔ ${ws.blocked}</span>
          <span title="Overdue">⏰ ${ws.overdue}</span>
        </div>
        <div class="health-meta">${ws.total} total steps · ⚠️ ${ws.riskScore} issue${ws.riskScore !== 1 ? 's' : ''}</div>
      </div>`;
    }).join('') + '</div>';
  }

  // ---- Timeline / Gantt ----
  async function loadTimeline() {
    const el = document.getElementById('timeline-container');
    if (!el) return;
    el.innerHTML = '<div class="loading">Loading...</div>';
    try {
      const res = await fetch(`${API_BASE}/api/analytics/timeline?track=${currentTrack}`);
      const data = await res.json();
      el.innerHTML = renderTimeline(data.steps, data.milestones);
    } catch (e) {
      el.innerHTML = '<div class="error">Failed to load timeline</div>';
    }
  }

  function renderTimeline(steps, milestones) {
    if (!steps || steps.length === 0) return '<div class="empty">No timeline data</div>';

    function parseDate(str) {
      if (!str || typeof str !== 'string') return null;
      const m = str.match(/^(\d{4}-\d{2}-\d{2})/);
      if (!m) return null;
      const d = new Date(m[1] + 'T00:00:00');
      return isNaN(d.getTime()) ? null : d;
    }

    let dated = steps.filter(s => parseDate(s.start) && parseDate(s.end));
    if (dated.length === 0) return '<div class="empty">No dated steps</div>';

    // Apply workstream filter
    if (filterWorkstream) {
      dated = dated.filter(s => s.workstream === filterWorkstream);
      if (dated.length === 0) return '<div class="empty">No dated steps for ' + filterWorkstream + '</div>';
    }

    const allTs = dated.flatMap(s => [parseDate(s.start).getTime(), parseDate(s.end).getTime()]);
    const minTs = Math.min(...allTs);
    const maxTs = Math.max(...allTs);
    const totalMs = maxTs - minTs || 1;
    const todayTs = Date.now();
    const todayPct = Math.min(100, Math.max(0, (todayTs - minTs) / totalMs * 100));

    const statusColors = {
      'Completed': '#10b981', 'In Progress': '#3b82f6',
      'Blocked': '#ef4444', 'Not Started': '#94a3b8'
    };

    const groups = {};
    dated.forEach(s => {
      if (!groups[s.workstream]) groups[s.workstream] = [];
      groups[s.workstream].push(s);
    });

    const months = [];
    const md = new Date(minTs);
    md.setDate(1);
    md.setMonth(md.getMonth() + 1);
    while (md.getTime() <= maxTs + 86400000 * 30) {
      const pct = (md.getTime() - minTs) / totalMs * 100;
      if (pct >= 0 && pct <= 105) months.push({ label: md.toLocaleString('en', { month: 'short', year: '2-digit' }), pct });
      md.setMonth(md.getMonth() + 1);
    }

    const statusCounts = {};
    dated.forEach(s => { statusCounts[s.status] = (statusCounts[s.status] || 0) + 1; });
    const summaryParts = Object.entries(statusCounts).map(([k, v]) => `${k}: ${v}`);

    let html = `<div style="font-size:0.8rem;color:#6b7280;margin-bottom:8px">${dated.length} steps with dates · ${Object.keys(groups).length} workstreams · ${summaryParts.join(' · ')}</div>`;
    html += '<div style="background:white;border-radius:8px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,0.1);overflow-x:auto;">';

    // Month header
    html += '<div style="position:relative;height:24px;border-bottom:1px solid #e5e7eb;margin-bottom:4px;margin-left:220px;">';
    months.forEach(m => {
      html += `<span style="position:absolute;left:${m.pct}%;font-size:0.7rem;color:#9ca3af;border-left:1px dashed #d1d5db;padding-left:3px;top:0;height:100%;">${m.label}</span>`;
    });
    html += `<div style="position:absolute;left:${todayPct}%;top:-4px;height:3000px;width:2px;background:#ef4444;opacity:0.5;z-index:2;" title="Today (${new Date().toLocaleDateString()})"></div>`;
    html += `<span style="position:absolute;left:${todayPct}%;top:-16px;font-size:0.65rem;color:#ef4444;transform:translateX(-50%);">Today</span>`;
    html += '</div>';

    // Workstream groups — collapsible
    for (const [ws, wsSteps] of Object.entries(groups)) {
      const groupId = 'tl-' + ws.replace(/[^a-zA-Z0-9]/g, '-');
      const showAll = filterWorkstream === ws;
      const maxShow = showAll ? wsSteps.length : 8;

      html += `<div style="border-top:1px solid #f3f4f6;">
        <div onclick="window._toggleGroup('${groupId}')" style="font-weight:600;font-size:0.85rem;color:#374151;padding:10px 0 4px;cursor:pointer;display:flex;align-items:center;gap:6px;user-select:none;">
          <span id="${groupId}-arrow" style="font-size:0.7rem;transition:transform 0.2s;display:inline-block;">▶</span>
          ${ws} (${wsSteps.length})
        </div>
        <div id="${groupId}" style="display:none;">`;

      wsSteps.slice(0, maxShow).forEach(s => {
        const sDate = parseDate(s.start);
        const eDate = parseDate(s.end);
        const startPct = (sDate.getTime() - minTs) / totalMs * 100;
        const endPct = (eDate.getTime() - minTs) / totalMs * 100;
        const widthPct = Math.max(0.8, endPct - startPct);
        const color = statusColors[s.status] || '#94a3b8';
        const statusIcon = s.status === 'Completed' ? '✅' : s.status === 'In Progress' ? '🔄' : s.status === 'Blocked' ? '⛔' : '⬜';
        html += `<div style="display:flex;align-items:center;height:24px;margin-bottom:2px;">
          <div style="width:220px;font-size:0.75rem;color:#4b5563;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex-shrink:0;padding-right:8px;" title="${s.name}">${statusIcon} <strong>${s.id}</strong> ${s.name}</div>
          <div style="flex:1;position:relative;height:16px;min-width:400px;">
            <div style="position:absolute;left:${startPct}%;width:${widthPct}%;height:100%;background:${color};border-radius:3px;min-width:4px;" title="${s.id}: ${s.name}&#10;${s.status}&#10;${s.start} → ${s.end}"></div>
          </div>
        </div>`;
      });
      if (wsSteps.length > maxShow) {
        html += `<div style="font-size:0.7rem;color:#9ca3af;padding:2px 0 2px 220px;">... and ${wsSteps.length - maxShow} more steps</div>`;
      }
      html += '</div></div>';
    }

    // Legend
    html += '<div style="font-size:0.75rem;color:#9ca3af;text-align:center;padding-top:12px;border-top:1px solid #f3f4f6;margin-top:8px;">';
    html += '<span style="display:inline-block;width:12px;height:12px;background:#10b981;border-radius:2px;vertical-align:middle;margin-right:2px;"></span> Completed &nbsp;&nbsp;';
    html += '<span style="display:inline-block;width:12px;height:12px;background:#3b82f6;border-radius:2px;vertical-align:middle;margin-right:2px;"></span> In Progress &nbsp;&nbsp;';
    html += '<span style="display:inline-block;width:12px;height:12px;background:#ef4444;border-radius:2px;vertical-align:middle;margin-right:2px;"></span> Blocked &nbsp;&nbsp;';
    html += '<span style="display:inline-block;width:12px;height:12px;background:#94a3b8;border-radius:2px;vertical-align:middle;margin-right:2px;"></span> Not Started &nbsp;&nbsp;';
    html += '<span style="display:inline-block;width:12px;height:2px;background:#ef4444;vertical-align:middle;margin-right:2px;"></span> Today';
    html += '</div>';

    html += '</div>';
    return html;
  }

  // Toggle collapsible group
  window._toggleGroup = function(groupId) {
    const el = document.getElementById(groupId);
    const arrow = document.getElementById(groupId + '-arrow');
    if (!el) return;
    if (el.style.display === 'none') {
      el.style.display = 'block';
      if (arrow) arrow.style.transform = 'rotate(90deg)';
    } else {
      el.style.display = 'none';
      if (arrow) arrow.style.transform = 'rotate(0deg)';
    }
  };

  // ---- Recent Changes ----
  async function loadChanges() {
    const el = document.getElementById('changes-feed');
    if (!el) return;
    el.innerHTML = '<div class="loading">Loading...</div>';
    try {
      const res = await fetch(`${API_BASE}/api/analytics/changes?hours=48`);
      const data = await res.json();
      el.innerHTML = renderChanges(data.changes);
    } catch (e) {
      el.innerHTML = '<div class="error">Failed to load changes</div>';
    }
  }

  function renderChanges(changes) {
    if (!changes || changes.length === 0) return '<div class="empty">📋 No changes in the last 48 hours</div>';
    const byDate = {};
    changes.forEach(c => {
      const d = new Date(c.changedAt).toLocaleDateString('en', { weekday: 'short', month: 'short', day: 'numeric' });
      if (!byDate[d]) byDate[d] = [];
      byDate[d].push(c);
    });

    let html = '';
    for (const [date, entries] of Object.entries(byDate)) {
      html += `<div class="change-date">${date}</div>`;
      entries.forEach(c => {
        const icon = c.newValue === 'Completed' ? '✅' : c.newValue === 'Blocked' ? '⛔' : c.newValue === 'In Progress' ? '🔄' : '📝';
        html += `<div class="change-item">
          <span class="change-icon">${icon}</span>
          <span class="change-step">${c.stepId}</span>
          <span class="change-detail">${c.field}: ${c.previousValue || '—'} → <strong>${c.newValue}</strong></span>
          <span class="change-meta">by ${c.changedBy} via ${c.source}</span>
        </div>`;
      });
    }
    return html;
  }

  // ---- HTML Template ----
  function getAnalyticsHTML() {
    return `
      <div class="analytics-dashboard">
        <div class="analytics-header">
          <h2>📊 Analytics Dashboard</h2>
          <div style="display:flex;align-items:center;gap:12px;">
            <span id="last-refresh" style="font-size:0.75rem;color:#9ca3af;"></span>
            <div class="track-toggle">
              <button class="track-toggle-btn active" data-track="Corp">Corp</button>
              <button class="track-toggle-btn" data-track="Fed">Fed</button>
            </div>
          </div>
        </div>

        <div id="filter-indicator" style="margin-bottom:8px;"></div>

        <div class="analytics-section">
          <h3>📊 Overall Health</h3>
          <div id="overall-health"></div>
        </div>

        <div class="analytics-section">
          <h3>🏥 Workstream Health <span style="font-size:0.75rem;font-weight:normal;color:#9ca3af;">— click a card to filter timeline</span></h3>
          <div id="health-cards"></div>
        </div>

        <div class="analytics-section">
          <h3>📅 Timeline <span style="font-size:0.75rem;font-weight:normal;color:#9ca3af;">— click ▶ to expand</span></h3>
          <div id="timeline-container"></div>
        </div>

        <div class="analytics-section">
          <h3>📋 Recent Changes</h3>
          <div id="changes-feed"></div>
        </div>
      </div>`;
  }
})();
