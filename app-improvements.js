/* Targeted hardening for the dashboard's existing browser-only workflow. */
(function () {
  'use strict';

  // Parse common INR/USD-style values, including Indian digit grouping and
  // accounting negatives such as (₹1,20,000). Invalid or ambiguous values
  // remain null so the report does not silently treat them as zero.
  window.parseCurrency = function parseCurrency(value) {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;

    let text = String(value).trim();
    if (!text) return null;

    let negative = false;
    if (/^\(.*\)$/.test(text)) {
      negative = true;
      text = text.slice(1, -1).trim();
    }

    text = text
      .replace(/^(?:INR|USD|EUR|GBP|JPY|CAD|AUD)\s*/i, '')
      .replace(/[₹$€£¥]/g, '')
      .replace(/\s/g, '');

    if (text.startsWith('-')) {
      negative = !negative;
      text = text.slice(1);
    } else if (text.startsWith('+')) {
      text = text.slice(1);
    }

    const westernGrouped = /^\d{1,3}(?:,\d{3})+(?:\.\d+)?$/;
    const indianGrouped = /^\d{1,2}(?:,\d{2})+,\d{3}(?:\.\d+)?$/;
    const ungrouped = /^\d+(?:\.\d+)?$/;
    if (!westernGrouped.test(text) && !indianGrouped.test(text) && !ungrouped.test(text)) return null;

    const amount = Number(text.replace(/,/g, ''));
    if (!Number.isFinite(amount)) return null;
    return negative ? -amount : amount;
  };

  // RFC-4180-style quoting plus spreadsheet formula-injection protection for
  // text cells. Numeric negatives stay numeric and are not rewritten.
  window.sanitizeCsvCell = function sanitizeCsvCell(value) {
    let text = String(value ?? '');
    if (typeof value !== 'number' && /^[\u0000-\u0020]*[=+@-]/.test(text)) {
      text = `'${text}`;
    }
    if (/[",\r\n]/.test(text)) text = `"${text.replace(/"/g, '""')}"`;
    return text;
  };

  // Replace removal-by-filename with a stable per-entry identity, so repeated
  // uploads with the same filename can be removed independently.
  window.renderUploadedFiles = function renderUploadedFiles() {
    const list = document.getElementById('uploadedFilesList');
    if (!list) return;
    list.replaceChildren();

    if (!state.files.length) {
      list.classList.add('empty');
      const item = document.createElement('li');
      item.textContent = 'No files uploaded yet.';
      list.appendChild(item);
      return;
    }

    list.classList.remove('empty');
    state.files.forEach((file) => {
      if (!file._dashboardId) {
        file._dashboardId = (window.crypto?.randomUUID?.() || `${Date.now()}-${Math.random()}`);
      }
      const id = file._dashboardId;
      const item = document.createElement('li');
      const meta = document.createElement('div');
      meta.className = 'file-meta';
      const name = document.createElement('span');
      name.className = 'file-name';
      name.textContent = file.name;
      const size = document.createElement('span');
      size.className = 'file-size';
      size.textContent = `${file.rowsCount} rows • ${formatBytes(file.size)}`;
      meta.append(name, size);

      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'remove-file-btn';
      remove.setAttribute('aria-label', `Remove ${file.name}`);
      remove.textContent = 'Remove';
      remove.addEventListener('click', () => {
        state.files = state.files.filter((entry) => entry._dashboardId !== id);
        state.rows = state.files.flatMap((entry) => entry.rows);
        state.detailSelection = { column: '', value: '', rows: [] };
        if (!state.files.length) {
          state.mapping = { location: '', revenue: [], expense: [] };
        } else {
          buildMappingsFromCurrentData();
        }
        renderAll();
      });

      item.append(meta, remove);
      list.appendChild(item);
    });
  };

  // Ensure optional/missing page sections don't make rendering throw.
  const renderRequirements = {
    renderMappingControls: 'mappingGrid',
    renderPreviewTable: 'previewTableHead',
    renderReportTable: 'reportTableBody',
    renderDetailPanel: 'detailSummary',
    renderChart: 'chartBars'
  };
  Object.entries(renderRequirements).forEach(([name, requiredId]) => {
    const original = window[name];
    if (typeof original !== 'function') return;
    window[name] = function guardedRender(...args) {
      if (!document.getElementById(requiredId)) return;
      return original.apply(this, args);
    };
  });

  const exportReport = window.exportCurrentReport;
  if (typeof exportReport === 'function') {
    window.exportCurrentReport = function exportCurrentReportSafely() {
      if (!state.summary || !state.reportRows.length) {
        setStatus('warning', 'There is no valid report to export yet.');
        return;
      }
      const rows = state.reportRows.filter((row) => !row.isTotal);
      const total = state.reportRows.find((row) => row.isTotal);
      const csvRows = [
        ['Location', 'Revenue', 'Expense', 'EBITDA'],
        ...rows.map((row) => [row.location, row.revenue, row.expense, row.ebitda]),
        ['Total', total?.revenue ?? 0, total?.expense ?? 0, total?.ebitda ?? 0]
      ].map((row) => row.map(window.sanitizeCsvCell).join(','));

      const blob = new Blob(['\uFEFF', csvRows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `misflow-report-${new Date().toISOString().slice(0, 10)}.csv`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 0);
      setStatus('success', 'CSV export created successfully.');
    };
  }

  function connectUploadButton() {
    const button = document.getElementById('uploadBtn');
    const input = document.getElementById('fileInput');
    if (button && input) button.addEventListener('click', () => input.click());
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', connectUploadButton, { once: true });
  } else {
    connectUploadButton();
  }
})();
