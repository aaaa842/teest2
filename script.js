const DEFAULT_DATA_URL = 'data.xlsx';
const DEFAULT_REFRESH_INTERVAL_MS = 60_000;

const SELECTORS = {
  tableBody: [
    '[data-role="data-table-body"]',
    '#data-table-body',
    '#data-table tbody',
    '#table-body'
  ],
  lastUpdated: [
    '[data-role="last-updated"]',
    '#last-updated',
    '#lastUpdated',
    '.last-updated'
  ],
  error: [
    '[data-role="error"]',
    '#error-message',
    '#errorMessage',
    '.error-message'
  ],
  refreshButton: [
    '[data-action="refresh-now"]',
    '#refresh-now',
    '#refreshNow',
    '#refresh-button',
    '.refresh-now'
  ],
  dataSource: [
    '[data-source-url]',
    '[data-data-url]',
    'meta[name="data-source"]'
  ],
  refreshInterval: [
    '[data-refresh-interval]',
    'meta[name="refresh-interval"]'
  ]
};

function queryElement(candidates) {
  for (const selector of candidates) {
    const element = document.querySelector(selector);
    if (element) {
      return element;
    }
  }
  return null;
}

function resolveDataSourceUrl() {
  const element = queryElement(SELECTORS.dataSource);
  if (element) {
    if (element.dataset?.sourceUrl) {
      return element.dataset.sourceUrl;
    }
    if (element.dataset?.dataUrl) {
      return element.dataset.dataUrl;
    }
    if (element.getAttribute) {
      const attrUrl = element.getAttribute('data-source-url') || element.getAttribute('data-data-url') || element.getAttribute('content') || element.getAttribute('href');
      if (attrUrl) {
        return attrUrl;
      }
    }
  }
  return DEFAULT_DATA_URL;
}

function resolveRefreshInterval() {
  const element = queryElement(SELECTORS.refreshInterval);
  let value;
  if (element) {
    if (element.dataset?.refreshInterval) {
      value = element.dataset.refreshInterval;
    } else if (element.getAttribute) {
      value = element.getAttribute('data-refresh-interval') || element.getAttribute('content');
    }
  }

  const parsed = Number.parseInt(value, 10);
  if (!Number.isNaN(parsed) && parsed > 0) {
    return parsed;
  }

  return DEFAULT_REFRESH_INTERVAL_MS;
}

function updateLastUpdatedLabel(date = new Date()) {
  const label = queryElement(SELECTORS.lastUpdated);
  if (!label) {
    return;
  }

  const formatted = date.toLocaleString();
  if ('textContent' in label) {
    label.textContent = formatted;
  }
  if ('dateTime' in label) {
    label.dateTime = date.toISOString();
  }
}

function showError(error) {
  const container = queryElement(SELECTORS.error);
  if (container) {
    const message = error instanceof Error ? error.message : String(error);
    container.textContent = message;
    container.hidden = false;
  }
  console.error(error);
}

function clearError() {
  const container = queryElement(SELECTORS.error);
  if (container) {
    container.textContent = '';
    container.hidden = true;
  }
}

function renderRows(rows) {
  const tableBody = queryElement(SELECTORS.tableBody);
  if (!tableBody || !Array.isArray(rows)) {
    return;
  }

  tableBody.innerHTML = '';

  const fragment = document.createDocumentFragment();

  rows.forEach((row, rowIndex) => {
    if (!Array.isArray(row)) {
      return;
    }

    const tr = document.createElement('tr');
    row.forEach((cell) => {
      const cellElement = document.createElement(rowIndex === 0 ? 'th' : 'td');
      cellElement.textContent = cell == null ? '' : cell;
      tr.appendChild(cellElement);
    });

    fragment.appendChild(tr);
  });

  tableBody.appendChild(fragment);
}

async function loadData(_forceNoCache = false) {
  const dataSource = resolveDataSourceUrl();
  const url = new URL(dataSource, window.location.href);

  // أضف معلمة طابع زمني في كل طلب للتغلب على التخزين المؤقت للمتصفح.
  // نبقي العلم في التوقيع للاتساق مع الاستدعاءات الحالية،
  // لكنه أصبح الآن غير ضروري تقريباً لأننا نضيف المعامل دائماً.
  url.searchParams.set('ts', Date.now().toString());

  try {
    const response = await fetch(url.toString(), { cache: 'no-store' });
    if (!response.ok) {
      if (response.status === 404) {
        throw new Error('تعذر العثور على ملف البيانات. تأكد من توفير "data.xlsx" أو تحديث رابط المصدر.');
      }
      throw new Error(`تعذر جلب البيانات (HTTP ${response.status})`);
    }

    const arrayBuffer = await response.arrayBuffer();

    if (window.XLSX) {
      const workbook = window.XLSX.read(arrayBuffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames?.[0];

      if (firstSheetName) {
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = window.XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        renderRows(rows);
      }
    }

    updateLastUpdatedLabel();
    clearError();
  } catch (error) {
    showError(error);
  }
}

function initialize() {
  const refreshButton = queryElement(SELECTORS.refreshButton);
  if (refreshButton) {
    refreshButton.addEventListener('click', () => loadData(true));
  }

  loadData(true);

  const interval = resolveRefreshInterval();
  setInterval(() => loadData(true), interval);
}

document.addEventListener('DOMContentLoaded', initialize);
