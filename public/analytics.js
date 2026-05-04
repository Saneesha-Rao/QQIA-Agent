/**
 * QQIA Agent — Analytics Dashboard
 * Provides burndown charts, workstream health, timeline, and change feed.
 * Loaded as a separate module from the main chat UI.
 */
(function() {
  'use strict';

  const API_BASE = '';
  let currentTrack = 'Corp';

  // ---- Init ----
  window.initAnalytics = function() {
    const container = document.getElementById('analytics-container');
    if (!container) return;
    container.innerHTML = getAnalyticsHTML();
    setupTrackToggle();
    loadAll();
  };

  function setupTrackToggle() {
    document.querySelectorAll('.track-toggle-btn').forEach(btn => {
      btn.addEventListener('click', function() {
        document.querySelectorAll('.track-toggle-btn').forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        currentTrack = this.dataset.track;
        loadAll();
      });
    });
  }

  function loadAll() {
    loadWorkstreamHealth();
    loadBurndown();
    loadTimeline();
    loadChanges();
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
    return '<div class="health-grid">' + workstreams.map(ws => `
      <div class="health-card health-${ws.health}" style="border-left: 4px solid ${colors[ws.health]}">
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
        <div class="health-meta">${ws.total} total steps · Risk: ${ws.riskScore}</div>
      </div>
    `).join('') + '</div>';
  }

  // ---- Burndown Chart ----
  async function loadBurndown() {
    const canvas = document.getElementById('burndown-chart');
    if (!canvas) return;
    try {
      const res = await fetch(`${API_BASE}/api/analytics/burndown?track=${currentTrack}`);
      const data = await res.json();
      renderBurndown(canvas, data.data);
    } catch (e) {
      console.error('Burndown error:', e);
    }
  }

  function renderBurndown(canvas, points) {
    if (!points || points.length === 0) {
      canvas.parentElement.querySelector('.chart-empty')?.classList.remove('hidden');
      return;
    }
    // Destroy previous chart if exists
    if (canvas._chart) canvas._chart.destroy();

    const ctx = canvas.getContext('2d');
    canvas._chart = new Chart(ctx, {
      type: 'line',
      data: {
        labels: points.map(p => p.date),
        datasets: [
          {
            label: 'Completed', data: points.map(p => p.completed),
            borderColor: '#10b981', backgroundColor: 'rgba(16,185,129,0.1)',
            fill: true, tension: 0.3
          },
          {
            label: 'In Progress', data: points.map(p => p.inProgress),
            borderColor: '#3b82f6', backgroundColor: 'rgba(59,130,246,0.1)',
            fill: true, tension: 0.3
          },
          {
            label: 'Blocked', data: points.map(p => p.blocked),
            borderColor: '#ef4444', backgroundColor: 'rgba(239,68,68,0.1)',
            fill: true, tension: 0.3
          },
          {
            label: 'Not Started', data: points.map(p => p.notStarted),
            borderColor: '#9ca3af', backgroundColor: 'rgba(156,163,175,0.1)',
            fill: true, tension: 0.3
          },
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom' }, title: { display: true, text: `Burndown — ${currentTrack} Track` } },
        scales: { y: { beginAtZero: true, stacked: false } }
      }
    });
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

    // Find date range
    const dates = steps.filter(s => s.start && s.end).flatMap(s => [new Date(s.start), new Date(s.end)]);
    if (dates.length === 0) return '<div class="empty">No dated steps</div>';
    const minDate = new Date(Math.min(...dates.map(d => d.getTime())));
    const maxDate = new Date(Math.max(...dates.map(d => d.getTime())));
    const totalDays = Math.max(1, (maxDate - minDate) / 86400000);
    const todayPct = Math.min(100, Math.max(0, ((Date.now() - minDate.getTime()) / 86400000 / totalDays) * 100));

    const statusColors = {
      'Completed': '#10b981', 'In Progress': '#3b82f6',
      'Blocked': '#ef4444', 'Not Started': '#d1d5db'
    };

    // Group by workstream
    const groups = {};
    steps.forEach(s => {
      if (!groups[s.workstream]) groups[s.workstream] = [];
      groups[s.workstream].push(s);
    });

    // Month markers
    const months = [];
    let d = new Date(minDate);
    d.setDate(1);
    while (d <= maxDate) {
      const pct = ((d - minDate) / 86400000 / totalDays) * 100;
      if (pct >= 0 && pct <= 100) months.push({ label: d.toLocaleString('en', { month: 'short', year: '2-digit' }), pct });
      d.setMonth(d.getMonth() + 1);
    }

    let html = '<div class="gantt">';
    // Header with month markers
    html += '<div class="gantt-header">';
    months.forEach(m => { html += `<div class="gantt-month" style="left:${m.pct}%">${m.label}</div>`; });
    html += `<div class="gantt-today" style="left:${todayPct}%" title="Today"></div>`;
    html += '</div>';

    // Rows
    for (const [ws, wsSteps] of Object.entries(groups)) {
      html += `<div class="gantt-group-label">${ws}</div>`;
      wsSteps.forEach(s => {
        if (!s.start || !s.end) return;
        const startPct = ((new Date(s.start) - minDate) / 86400000 / totalDays) * 100;
        const widthPct = Math.max(1, ((new Date(s.end) - new Date(s.start)) / 86400000 / totalDays) * 100);
        const color = statusColors[s.status] || '#9ca3af';
        html += `<div class="gantt-row">
          <div class="gantt-label" title="${s.name}">${s.id}</div>
          <div class="gantt-bar-area">
            <div class="gantt-bar" style="left:${startPct}%;width:${widthPct}%;background:${color}"
                 title="${s.id}: ${s.name}\n${s.status} (${s.start} → ${s.end})"></div>
          </div>
        </div>`;
      });
    }

    // Milestone markers
    if (milestones && milestones.length > 0) {
      html += '<div class="gantt-group-label">📌 Milestones</div>';
      milestones.forEach(m => {
        const mDate = currentTrack === 'Corp' ? m.corpDate : m.fedDate;
        if (!mDate) return;
        const pct = ((new Date(mDate) - minDate) / 86400000 / totalDays) * 100;
        if (pct < 0 || pct > 100) return;
        html += `<div class="gantt-row">
          <div class="gantt-label" title="${m.milestone}">${m.id || '◆'}</div>
          <div class="gantt-bar-area">
            <div class="gantt-milestone" style="left:${pct}%" title="${m.milestone}: ${mDate}">◆</div>
          </div>
        </div>`;
      });
    }

    html += '</div>';
    return html;
  }

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
    // Group by date
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
          <div class="track-toggle">
            <button class="track-toggle-btn active" data-track="Corp">Corp</button>
            <button class="track-toggle-btn" data-track="Fed">Fed</button>
          </div>
        </div>

        <div class="analytics-section">
          <h3>🏥 Workstream Health</h3>
          <div id="health-cards"></div>
        </div>

        <div class="analytics-section">
          <h3>📉 Burndown</h3>
          <div class="chart-wrapper">
            <canvas id="burndown-chart" height="300"></canvas>
            <div class="chart-empty hidden">Not enough data yet — burndown will populate over time as snapshots are recorded daily.</div>
          </div>
        </div>

        <div class="analytics-section">
          <h3>📅 Timeline</h3>
          <div id="timeline-container"></div>
        </div>

        <div class="analytics-section">
          <h3>📋 Recent Changes</h3>
          <div id="changes-feed"></div>
        </div>
      </div>`;
  }
})();
