const MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024;

const SAMPLE_DATA = [
  { Location: 'Bengaluru', Revenue: 1050000, Expense: 740000 },
  { Location: 'Hyderabad', Revenue: 980000, Expense: 700000 },
  { Location: 'Chennai', Revenue: 860000, Expense: 610000 },
  { Location: 'Bengaluru', Revenue: 1160000, Expense: 810000 },
  { Location: 'Pune', Revenue: 730000, Expense: 520000 },
  { Location: 'Hyderabad', Revenue: 890000, Expense: 670000 },
  { Location: 'Chennai', Revenue: 760000, Expense: 540000 },
  { Location: 'Pune', Revenue: 675000, Expense: 495000 },
  { Location: 'Delhi', Revenue: 930000, Expense: 690000 },
  { Location: 'Delhi', Revenue: 1040000, Expense: 780000 }
];

const SAMPLE_EXCEL_ROWS = [
  ['Location', 'Revenue', 'Expense'],
  ['Bengaluru', 1050000, 740000],
  ['Hyderabad', 980000, 700000],
  ['Chennai', 860000, 610000],
  ['Bengaluru', 1160000, 810000],
  ['Pune', 730000, 520000],
  ['Hyderabad', 890000, 670000],
  ['Chennai', 760000, 540000],
  ['Pune', 675000, 495000],
  ['Delhi', 930000, 690000],
  ['Delhi', 1040000, 780000]
];

const state = {
  files: [],
  rows: [],
  mapping: {
    location: '',
    revenue: [],
    expense: []
  },
  reportRows: [],
  summary: null,
  filterText: '',
  sortValue: 'revenue-desc'
};

const currencyFormatter = new Intl.NumberFormat('en-IN', {
  style: 'currency',
  currency: 'INR',
  maximumFractionDigits: 0
});

const numberFormatter = new Intl.NumberFormat('en-IN', {
  maximumFractionDigits: 0
});

document.addEventListener('DOMContentLoaded', () => {
  initializeApp();
});

function initializeApp() {
  const fileInput = document.getElementById('fileInput');
  const uploadZone = document.getElementById('uploadZone');
  const loadSampleBtn = document.getElementById('loadSampleBtn');
  const clearDataBtn = document.getElementById('clearDataBtn');
  const exportBtn = document.getElementById('exportBtn');
  const searchInput = document.getElementById('reportSearch');
  const sortSelect = document.getElementById('sortBy');
  const downloadSampleExcelBtn = document.getElementById('downloadSampleExcelBtn');

  fileInput.addEventListener('change', async (event) => {
    const files = Array.from(event.target.files || []);
    if (files.length) {
      await handleFiles(files);
      event.target.value = '';
    }
  });

  ['dragenter', 'dragover'].forEach((type) => {
    uploadZone.addEventListener(type, (event) => {
      event.preventDefault();
      uploadZone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach((type) => {
    uploadZone.addEventListener(type, (event) => {
      event.preventDefault();
      uploadZone.classList.remove('dragover');
    });
  });

  uploadZone.addEventListener('drop', async (event) => {
    const files = Array.from(event.dataTransfer?.files || []);
    if (files.length) {
      await handleFiles(files);
    }
  });

  downloadSampleExcelBtn.addEventListener('click', () => {
    downloadSampleExcel();
  });

  loadSampleBtn.addEventListener('click', () => {
    loadSampleData();
  });

  clearDataBtn.addEventListener('click', () => {
    clearAllData();
  });

  exportBtn.addEventListener('click', () => {
    exportCurrentReport();
  });

  searchInput.addEventListener('input', (event) => {
    state.filterText = event.target.value.trim().toLowerCase();
    renderReportTable();
  });

  sortSelect.addEventListener('change', (event) => {
    state.sortValue = event.target.value;
    renderReportTable();
  });

  if (typeof XLSX === 'undefined') {
    setStatus('error', 'SheetJS failed to load. Please refresh the page or check the CDN connection.');
  }

  renderUploadedFiles();
  renderMappingControls();
  renderPreviewTable();
  renderReportTable();
  updateKpiCards();
}

function normalizeColumnName(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function normalizeSelectedColumns(value) {
  if (Array.isArray(value)) {
    return value.filter(Boolean);
  }

  if (typeof value === 'string' && value) {
    return [value];
  }

  return [];
}

function isRevenueLikeColumn(columnName) {
  const normalized = normalizeColumnName(columnName);
  const revenueKeywords = [
    'revenue', 'sales', 'income', 'turnover', 'net revenue', 'gross revenue', 'total revenue',
    'sell', 'amount received', 'sales revenue', 'total sales'
  ];

  return revenueKeywords.some((keyword) => normalized.includes(keyword));
}

function isExpenseLikeColumn(columnName) {
  const normalized = normalizeColumnName(columnName);
  const expenseKeywords = [
    'expense', 'expenses', 'cost', 'costs', 'spend', 'operating expense', 'total expense',
    'expense amount', 'cost of goods', 'marketing', 'salary', 'rent', 'payroll', 'overhead'
  ];

  return expenseKeywords.some((keyword) => normalized.includes(keyword));
}

function preferMapping(columns) {
  const rules = {
    location: ['branch', 'location', 'store', 'site', 'region', 'city', 'office', 'area'],
    revenue: ['revenue', 'sales', 'income', 'turnover', 'amount', 'net sales', 'total revenue'],
    expense: ['expense', 'expenses', 'cost', 'costs', 'spend', 'operating expense', 'total expense']
  };

  const matches = {};

  Object.entries(rules).forEach(([key, keywords]) => {
    const target = columns.filter((column) => {
      const normalized = normalizeColumnName(column);
      return keywords.some((keyword) => normalized.includes(keyword));
    });

    matches[key] = key === 'location' ? (target[0] || '') : target;
  });

  return matches;
}

function isSupportedFile(file) {
  const supported = ['csv', 'xls', 'xlsx'];
  const extension = file.name.split('.').pop()?.toLowerCase();
  return supported.includes(extension || '');
}

function generateDummyWorkbook() {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(SAMPLE_EXCEL_ROWS);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'MISFlow_Data');
  return workbook;
}

function downloadSampleExcel() {
  if (typeof XLSX === 'undefined') {
    setStatus('error', 'SheetJS did not load correctly. The dummy Excel file could not be generated.');
    return;
  }

  const workbook = generateDummyWorkbook();
  const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });

  const link = document.createElement('a');
  const dateStamp = new Date().toISOString().slice(0, 10);
  link.href = URL.createObjectURL(blob);
  link.download = `misflow-sample-report-${dateStamp}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);

  setStatus('success', 'Dummy Excel file downloaded. Re-upload it to test the import and report workflow.');
}

async function handleFiles(files) {
  const validFiles = [];
  const errors = [];

  for (const file of files) {
    if (!isSupportedFile(file)) {
      errors.push(`${file.name} is not a supported CSV, XLS, or XLSX file.`);
      continue;
    }

    if (file.size > MAX_FILE_SIZE_BYTES) {
      errors.push(`${file.name} exceeds the 10 MB limit.`);
      continue;
    }

    validFiles.push(file);
  }

  if (!validFiles.length) {
    if (errors.length) {
      setStatus('error', errors.join(' '));
    }
    return;
  }

  const addedFiles = [];
  const fileErrors = [];

  for (const file of validFiles) {
    try {
      const rows = await parseWorkbook(file);
      if (!rows.length) {
        fileErrors.push(`${file.name} is empty or has no usable rows.`);
        continue;
      }

      addedFiles.push({
        name: file.name,
        size: file.size,
        rows,
        rowsCount: rows.length
      });
    } catch (error) {
      fileErrors.push(`${file.name} could not be read. ${error.message || 'Please check the file format.'}`);
    }
  }

  if (!addedFiles.length) {
    setStatus('error', fileErrors.join(' '));
    return;
  }

  state.files.push(...addedFiles);
  state.rows = state.files.flatMap((item) => item.rows);
  buildMappingsFromCurrentData();

  if (fileErrors.length) {
    setStatus('warning', `${fileErrors.join(' ')} Some files were skipped.`);
  } else {
    setStatus('success', `${addedFiles.length} file(s) processed successfully.`);
  }

  renderAll();
}

async function parseWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', raw: false, cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  const firstSheet = workbook.Sheets[firstSheetName];

  if (!firstSheet) {
    throw new Error('The workbook does not contain any sheets.');
  }

  const rawRows = XLSX.utils.sheet_to_json(firstSheet, {
    raw: false,
    defval: '',
    blankrows: false
  });

  if (!rawRows.length) {
    return [];
  }

  if (Array.isArray(rawRows[0])) {
    const headers = rawRows[0].map((header, idx) => String(header || `Column ${idx + 1}`));
    return rawRows.slice(1).map((row) => {
      const record = {};
      headers.forEach((header, idx) => {
        record[header] = row[idx] ?? '';
      });
      return record;
    });
  }

  return rawRows;
}

function loadSampleData() {
  const rows = SAMPLE_DATA.map((row) => ({
    ...row,
    Location: row.Location || 'Unassigned'
  }));

  state.files = [{ name: 'sample-business-report.csv', size: 0, rows, rowsCount: rows.length }];
  state.rows = rows;
  buildMappingsFromCurrentData();
  setStatus('success', 'Sample business data loaded. Review the mapped fields and generate the report.');
  renderAll();
}

function clearAllData() {
  state.files = [];
  state.rows = [];
  state.mapping = {
    location: '',
    revenue: [],
    expense: []
  };
  state.reportRows = [];
  state.summary = null;
  state.filterText = '';
  state.sortValue = 'revenue-desc';

  document.getElementById('reportSearch').value = '';
  document.getElementById('sortBy').value = 'revenue-desc';
  document.getElementById('exportBtn').disabled = true;

  setStatus('info', 'All uploaded data and demo content were cleared.');
  renderAll();
}

function buildMappingsFromCurrentData() {
  const columns = getAllColumns(state.rows);
  const mapping = preferMapping(columns);

  const inferredRevenue = Array.isArray(mapping.revenue) ? mapping.revenue : normalizeSelectedColumns(mapping.revenue);
  const inferredExpense = Array.isArray(mapping.expense) ? mapping.expense : normalizeSelectedColumns(mapping.expense);

  const allRevenueColumns = columns.filter((column) => isRevenueLikeColumn(column));
  const allExpenseColumns = columns.filter((column) => isExpenseLikeColumn(column));

  state.mapping = {
    location: mapping.location || '',
    revenue: inferredRevenue.length ? inferredRevenue : allRevenueColumns.slice(0, 3),
    expense: inferredExpense.length ? inferredExpense : allExpenseColumns.slice(0, 3)
  };

  renderMappingControls();
}

function getAllColumns(rows) {
  const columns = new Set();
  rows.forEach((row) => {
    Object.keys(row || {}).forEach((key) => columns.add(key));
  });
  return [...columns];
}

function renderMappingControls() {
  const mappingGrid = document.getElementById('mappingGrid');
  const columns = getAllColumns(state.rows);

  if (!columns.length) {
    mappingGrid.innerHTML = '<p class="empty-copy">Upload data to map location, revenue, and expense columns.</p>';
    return;
  }

  const labels = [
    { key: 'location', label: 'Location / Branch', multi: false },
    { key: 'revenue', label: 'Revenue columns', multi: true },
    { key: 'expense', label: 'Expense columns', multi: true }
  ];

  mappingGrid.innerHTML = '';

  labels.forEach(({ key, label, multi }) => {
    const field = document.createElement('div');
    field.className = 'map-field';

    const labelEl = document.createElement('label');
    labelEl.setAttribute('for', `mapping-${key}`);
    labelEl.textContent = label;

    const select = document.createElement('select');
    select.id = `mapping-${key}`;
    select.name = key;
    select.setAttribute('aria-label', label);

    if (multi) {
      select.multiple = true;
      select.size = Math.min(columns.length + 1, 6);
    }

    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = multi ? 'Select one or more columns' : 'Select a column';
    if (!multi) {
      select.appendChild(placeholder);
    }

    columns.forEach((column) => {
      const option = document.createElement('option');
      option.value = column;
      option.textContent = column;

      const selectedValues = normalizeSelectedColumns(state.mapping[key]);
      if (selectedValues.includes(column)) {
        option.selected = true;
      }

      select.appendChild(option);
    });

    select.addEventListener('change', (event) => {
      if (!multi) {
        state.mapping[key] = event.target.value;
      } else {
        state.mapping[key] = Array.from(event.target.selectedOptions).map((option) => option.value);
      }

      renderAll();
    });

    field.appendChild(labelEl);
    field.appendChild(select);
    mappingGrid.appendChild(field);
  });
}

function renderUploadedFiles() {
  const list = document.getElementById('uploadedFilesList');
  list.innerHTML = '';

  if (!state.files.length) {
    list.classList.add('empty');
    const li = document.createElement('li');
    li.textContent = 'No files uploaded yet.';
    list.appendChild(li);
    return;
  }

  list.classList.remove('empty');

  state.files.forEach((item) => {
    const li = document.createElement('li');

    const meta = document.createElement('div');
    meta.className = 'file-meta';

    const name = document.createElement('span');
    name.className = 'file-name';
    name.textContent = item.name;

    const size = document.createElement('span');
    size.className = 'file-size';
    size.textContent = `${item.rowsCount} rows • ${formatBytes(item.size)}`;

    meta.appendChild(name);
    meta.appendChild(size);

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'remove-file-btn';
    removeBtn.setAttribute('aria-label', `Remove ${item.name}`);
    removeBtn.textContent = 'Remove';
    removeBtn.addEventListener('click', () => {
      const index = state.files.findIndex((entry) => entry.name === item.name);
      if (index >= 0) {
        state.files.splice(index, 1);
        state.rows = state.files.flatMap((entry) => entry.rows);
        if (!state.files.length) {
          state.mapping = {
            location: '',
            revenue: [],
            expense: []
          };
        } else {
          buildMappingsFromCurrentData();
        }
        renderAll();
      }
    });

    li.appendChild(meta);
    li.appendChild(removeBtn);
    list.appendChild(li);
  });
}

function renderPreviewTable() {
  const head = document.getElementById('previewTableHead');
  const body = document.getElementById('previewTableBody');

  head.innerHTML = '';
  body.innerHTML = '';

  const columns = getAllColumns(state.rows);
  if (!columns.length) {
    head.innerHTML = '<th>No data</th>';
    body.innerHTML = '<tr><td>No preview available.</td></tr>';
    return;
  }

  columns.slice(0, 8).forEach((column) => {
    const th = document.createElement('th');
    th.textContent = column;
    head.appendChild(th);
  });

  const previewRows = state.rows.slice(0, 8);
  previewRows.forEach((row) => {
    const tr = document.createElement('tr');
    columns.slice(0, 8).forEach((column) => {
      const td = document.createElement('td');
      td.textContent = formatCellValue(row[column]);
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
}

function renderReportTable() {
  const tableBody = document.getElementById('reportTableBody');
  tableBody.innerHTML = '';

  const locationColumn = state.mapping.location || '';
  const revenueColumns = normalizeSelectedColumns(state.mapping.revenue);
  const expenseColumns = normalizeSelectedColumns(state.mapping.expense);

  if (!state.rows.length || !locationColumn || !revenueColumns.length || !expenseColumns.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 4;
    cell.textContent = 'Complete the column mapping to generate the report.';
    row.appendChild(cell);
    tableBody.appendChild(row);
    document.getElementById('exportBtn').disabled = true;
    document.getElementById('chartBars').innerHTML = '';
    updateKpiCards();
    return;
  }

  const report = buildReport();
  state.reportRows = report.rows;
  state.summary = report.summary;
  document.getElementById('exportBtn').disabled = false;

  const searchQuery = state.filterText;
  const filteredRows = [...state.reportRows].filter((row) => {
    if (!searchQuery) {
      return true;
    }
    return String(row.location).toLowerCase().includes(searchQuery);
  });

  const sortedRows = applySorting(filteredRows, state.sortValue);

  if (!sortedRows.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 4;
    cell.textContent = 'No matching locations found for your search.';
    row.appendChild(cell);
    tableBody.appendChild(row);
    renderChart([]);
    updateKpiCards();
    return;
  }

  sortedRows.forEach((row) => {
    const tr = document.createElement('tr');
    if (row.isTotal) {
      tr.classList.add('total-row');
    }

    const location = document.createElement('td');
    location.textContent = row.location;

    const revenue = document.createElement('td');
    revenue.textContent = formatCurrency(row.revenue);

    const expense = document.createElement('td');
    expense.textContent = formatCurrency(row.expense);

    const ebitda = document.createElement('td');
    ebitda.textContent = formatCurrency(row.ebitda);
    ebitda.classList.toggle('negative', row.ebitda < 0);
    ebitda.classList.toggle('positive', row.ebitda >= 0);

    tr.appendChild(location);
    tr.appendChild(revenue);
    tr.appendChild(expense);
    tr.appendChild(ebitda);
    tableBody.appendChild(tr);
  });

  renderChart(sortedRows.filter((row) => !row.isTotal));
  updateKpiCards();
}

function applySorting(rows, value) {
  const sorted = [...rows];

  switch (value) {
    case 'revenue-asc':
      sorted.sort((a, b) => a.revenue - b.revenue);
      break;
    case 'ebitda-desc':
      sorted.sort((a, b) => b.ebitda - a.ebitda);
      break;
    case 'ebitda-asc':
      sorted.sort((a, b) => a.ebitda - b.ebitda);
      break;
    case 'location-asc':
      sorted.sort((a, b) => a.location.localeCompare(b.location));
      break;
    case 'revenue-desc':
    default:
      sorted.sort((a, b) => b.revenue - a.revenue);
      break;
  }

  return sorted;
}

function sumSelectedColumns(row, selectedColumns) {
  const numericValues = selectedColumns
    .map((column) => parseCurrency(row[column]))
    .filter((value) => value !== null && Number.isFinite(value));

  return numericValues.reduce((sum, value) => sum + value, 0);
}

function buildReport() {
  const locationColumn = state.mapping.location;
  const revenueColumns = normalizeSelectedColumns(state.mapping.revenue);
  const expenseColumns = normalizeSelectedColumns(state.mapping.expense);
  const totals = new Map();
  const warnings = [];

  state.rows.forEach((row, index) => {
    const rawLocation = row[locationColumn];
    const revenue = sumSelectedColumns(row, revenueColumns);
    const expense = sumSelectedColumns(row, expenseColumns);

    const location = String(rawLocation ?? '').trim() || 'Unassigned';

    if (!revenueColumns.length || !expenseColumns.length) {
      warnings.push(`Row ${index + 1}: please select at least one revenue column and one expense column.`);
      return;
    }

    if (!totals.has(location)) {
      totals.set(location, { location, revenue: 0, expense: 0 });
    }

    const rowEntry = totals.get(location);
    rowEntry.revenue += revenue;
    rowEntry.expense += expense;
  });

  if (warnings.length) {
    setStatus('warning', warnings.join(' '));
  }

  const rows = [...totals.values()].map((entry) => ({
    location: entry.location,
    revenue: entry.revenue,
    expense: entry.expense,
    ebitda: entry.revenue - entry.expense,
    isTotal: false
  }));

  const totalRevenue = rows.reduce((sum, entry) => sum + entry.revenue, 0);
  const totalExpense = rows.reduce((sum, entry) => sum + entry.expense, 0);
  const totalEbitda = totalRevenue - totalExpense;

  rows.push({
    location: 'Total',
    revenue: totalRevenue,
    expense: totalExpense,
    ebitda: totalEbitda,
    isTotal: true
  });

  return {
    rows,
    summary: {
      totalRevenue,
      totalExpense,
      ebitda: totalEbitda,
      locationCount: new Set(rows.filter((row) => !row.isTotal).map((row) => row.location)).size
    }
  };
}

function renderChart(rows) {
  const chartPanel = document.getElementById('chartBars');
  chartPanel.innerHTML = '';

  if (!rows.length) {
    chartPanel.innerHTML = '<p class="empty-copy">No chart data available.</p>';
    return;
  }

  const maxValue = Math.max(...rows.map((row) => Math.max(Math.abs(row.revenue), Math.abs(row.expense), 1)), 1);

  rows.slice(0, 6).forEach((row) => {
    const rowWrap = document.createElement('div');
    rowWrap.className = 'chart-row';

    const top = document.createElement('div');
    top.className = 'chart-row-top';

    const label = document.createElement('span');
    label.textContent = row.location;

    const value = document.createElement('span');
    value.textContent = formatCurrency(row.revenue);

    const track = document.createElement('div');
    track.className = 'bar-track';

    const fill = document.createElement('div');
    fill.className = 'bar-fill';
    const width = Math.max((row.revenue / maxValue) * 100, 6);
    fill.style.width = `${width}%`;

    top.appendChild(label);
    top.appendChild(value);
    track.appendChild(fill);
    rowWrap.appendChild(top);
    rowWrap.appendChild(track);
    chartPanel.appendChild(rowWrap);
  });
}

function updateKpiCards() {
  const summary = state.summary;
  const revenueCard = document.getElementById('kpiRevenue');
  const expenseCard = document.getElementById('kpiExpense');
  const ebitdaCard = document.getElementById('kpiEbitda');
  const locationsCard = document.getElementById('kpiLocations');

  if (!summary) {
    revenueCard.textContent = '₹0';
    expenseCard.textContent = '₹0';
    ebitdaCard.textContent = '₹0';
    locationsCard.textContent = '0';
    return;
  }

  revenueCard.textContent = formatCurrency(summary.totalRevenue);
  expenseCard.textContent = formatCurrency(summary.totalExpense);
  ebitdaCard.textContent = formatCurrency(summary.ebitda);
  ebitdaCard.classList.toggle('negative', summary.ebitda < 0);
  locationsCard.textContent = String(summary.locationCount);
}

function parseCurrency(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  const normalized = String(value)
    .trim()
    .replace(/[$,]/g, '')
    .replace(/₹/g, '')
    .replace(/\s+/g, '');

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function formatCurrency(value) {
  return currencyFormatter.format(Number(value || 0));
}

function formatBytes(sizeInBytes) {
  if (!sizeInBytes) return '0 KB';
  const units = ['B', 'KB', 'MB', 'GB'];
  let sizeValue = sizeInBytes;
  let unitIndex = 0;

  while (sizeValue >= 1024 && unitIndex < units.length - 1) {
    sizeValue /= 1024;
    unitIndex += 1;
  }

  return `${numberFormatter.format(sizeValue)} ${units[unitIndex]}`;
}

function formatCellValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}

function sanitizeCsvCell(value) {
  const stringValue = String(value ?? '');
  if (/^(?:[=+\-@])/.test(stringValue)) {
    return `'${stringValue}`;
  }
  return stringValue;
}

function exportCurrentReport() {
  if (!state.summary || !state.reportRows.length) {
    setStatus('warning', 'There is no valid report to export yet.');
    return;
  }

  const rows = state.reportRows.filter((row) => !row.isTotal);
  const totalRow = state.reportRows.find((row) => row.isTotal);

  const csvRows = [
    ['Location', 'Revenue', 'Expense', 'EBITDA'],
    ...rows.map((row) => [row.location, row.revenue, row.expense, row.ebitda]),
    ['Total', totalRow?.revenue || 0, totalRow?.expense || 0, totalRow?.ebitda || 0]
  ].map((row) => row.map((cell) => sanitizeCsvCell(cell)).join(','));

  const content = csvRows.join('\n');
  const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  const dateStamp = new Date().toISOString().slice(0, 10);
  anchor.href = url;
  anchor.download = `misflow-report-${dateStamp}.csv`;
  anchor.click();
  URL.revokeObjectURL(url);

  setStatus('success', 'CSV export created successfully.');
}

function setStatus(type, message) {
  const statusBanner = document.getElementById('statusBanner');
  statusBanner.className = `status-banner ${type}`;
  statusBanner.textContent = message;
}

function renderAll() {
  renderUploadedFiles();
  renderMappingControls();
  renderPreviewTable();
  renderReportTable();
}

window.addEventListener('beforeunload', () => {
  // Intentionally do not persist uploaded financial data in browser storage.
});

