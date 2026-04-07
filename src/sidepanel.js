/* ═══════════════════════════════════════════════════════════
   F&SCM Entity Browser — Full-Page Application
   v1.1 — Two-column layout, schema fix, 100-entity paging
   ═══════════════════════════════════════════════════════════ */

const App = {
  tabId:               null,
  baseUrl:             null,
  entityList:          [],
  filteredList:        [],
  displayCount:        100,
  selectedEntity:      null,
  activeView:          null,
  metadata:            null,
  metadataDoc:         null,
  _metadataDocSource:  null,
  entityTypeCache:     {},
  lastQueryData:       null,
  lastSearchQuery:     null,
  aliasHint:           null,
  selectedLegalEntity: null,
  colFilters:          {},   // col-header search state � survives grid re-renders
  _colLabels:          {},   // display name map: apiFieldName ? readable label
  _schemaForEntity:    null,
  _schemaTypeMap:      null,
  pendingFilterQuery:  false,
  queryInFlight:       false,
  lastGridScroll:      { top: 0, left: 0 },
  filterMode:          'compatible',
  filterRowSeq:        0,
  lastFallbackInfo:    null,
};

const CACHE_KEY_PREFIX = 'entityList_';
const CACHE_TTL_MS     = 24 * 60 * 60 * 1000; // 24 h
const COLUMN_FILTER_DEBOUNCE_MS = 900;

// ── Bootstrap ──────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  bindStaticEvents();
  bindEnvPicker();
  init();
});

async function init() {
  showLoading('Connecting to F\u0026SCM\u2026');

  let tabId = null, tabUrl = null;
  try {
    const stored = await chrome.storage.session.get(['activeFscmTabId', 'activeFscmTabUrl']);
    if (stored.activeFscmTabId && stored.activeFscmTabUrl?.includes('.operations.dynamics.com')) {
      tabId = stored.activeFscmTabId;
      tabUrl = stored.activeFscmTabUrl;
    } else {
      const tabs = await chrome.tabs.query({ url: 'https://*.operations.dynamics.com/*' });
      if (tabs.length > 0) {
        tabId  = tabs[0].id;
        tabUrl = tabs[0].url;
        await chrome.storage.session.set({ activeFscmTabId: tabId, activeFscmTabUrl: tabUrl });
      }
    }
  } catch (_) {}

  if (!tabId || !tabUrl) {
    hideLoading();
    renderNoTab();
    return;
  }

  App.tabId   = tabId;
  App.baseUrl = new URL(tabUrl).origin;

  const envName = App.baseUrl.replace('https://', '').split('.')[0];
  document.getElementById('envBadge').textContent = envName + ' \u25be';

  try {
    await loadEntityList();
  } catch (e) {
    hideLoading();
    showError('Could not load entity list: ' + e.message);
    renderNoTab();
  }
}

// ── Static event bindings ──────────────────────────────────

function bindStaticEvents() {
  const searchInput = document.getElementById('searchInput');
  const clearSearch = document.getElementById('clearSearch');

  searchInput.addEventListener('input', () => {
    const q = searchInput.value.trim();
    clearSearch.classList.toggle('hidden', q === '');
    App.displayCount = 100;
    filterEntityList(q);
  });

  clearSearch.addEventListener('click', () => {
    searchInput.value = '';
    clearSearch.classList.add('hidden');
    App.displayCount = 100;
    filterEntityList('');
  });

  document.getElementById('refreshEntities').addEventListener('click', async () => {
    const entityCacheKey   = CACHE_KEY_PREFIX + App.baseUrl;
    const metadataCacheKey = 'metadata_' + App.baseUrl;
    await chrome.storage.local.remove([entityCacheKey, metadataCacheKey]);
    App.metadata = null;
    App.metadataDoc = null;
    App._metadataDocSource = null;
    App.entityTypeCache = {};
    App._schemaForEntity = null;
    App._schemaTypeMap = null;
    await loadEntityList(true);
  });

  document.getElementById('closeError').addEventListener('click', () => {
    document.getElementById('errorBar').classList.add('hidden');
  });
}

// ── Environment Picker ─────────────────────────────────────

function bindEnvPicker() {
  const badge      = document.getElementById('envBadge');
  const picker     = document.getElementById('envPicker');
  const customUrl  = document.getElementById('envCustomUrl');
  const connectBtn = document.getElementById('envCustomConnect');

  badge.addEventListener('click', async (e) => {
    e.stopPropagation();
    if (!picker.classList.contains('hidden')) {
      picker.classList.add('hidden');
    } else {
      await renderEnvPickerList();
      picker.classList.remove('hidden');
    }
  });

  document.addEventListener('click', (e) => {
    if (!picker.contains(e.target) && e.target !== badge) picker.classList.add('hidden');
  });

  connectBtn.addEventListener('click', () => {
    const raw = customUrl.value.trim().replace(/\/$/, '');
    if (!raw.startsWith('https://') || !raw.includes('.operations.dynamics.com')) {
      showError('URL must be https://*.operations.dynamics.com');
      return;
    }
    let origin;
    try {
      origin = new URL(raw).origin;
    } catch (_) {
      showError('Invalid URL format. Example: https://myenv.operations.dynamics.com');
      return;
    }
    picker.classList.add('hidden');
    switchEnvironment(null, origin);
  });

  customUrl.addEventListener('keydown', e => { if (e.key === 'Enter') connectBtn.click(); });
}

async function renderEnvPickerList() {
  const listEl = document.getElementById('envPickerList');
  listEl.innerHTML = '<div style="padding:8px 14px;font-size:12px;color:var(--text-muted)">Scanning\u2026</div>';

  let openTabs = [];
  try { openTabs = await chrome.tabs.query({ url: 'https://*.operations.dynamics.com/*' }); } catch (_) {}

  if (openTabs.length === 0) {
    listEl.innerHTML = '<div style="padding:8px 14px;font-size:12px;color:var(--text-muted)">No F\u0026SCM tabs open. Use the URL below.</div>';
    return;
  }

  const seen = new Map();
  openTabs.forEach(tab => {
    try { const o = new URL(tab.url).origin; if (!seen.has(o)) seen.set(o, tab); } catch (_) {}
  });

  listEl.innerHTML = [...seen.entries()].map(([origin, tab]) => {
    const host   = origin.replace('https://', '');
    const envKey = host.split('.')[0];
    const active = App.baseUrl === origin;
    return `
      <div class="env-picker-item ${active ? 'active' : ''}" data-origin="${escHtml(origin)}" data-tabid="${tab.id}">
        <div class="env-picker-item-dot"></div>
        <div class="env-picker-item-info">
          <div class="env-picker-item-name">${escHtml(envKey)}</div>
          <div class="env-picker-item-url">${escHtml(host)}</div>
        </div>
        <span class="env-picker-item-tag">Open tab</span>
      </div>`;
  }).join('');

  listEl.querySelectorAll('.env-picker-item').forEach(el => {
    el.addEventListener('click', () => {
      document.getElementById('envPicker').classList.add('hidden');
      switchEnvironment(parseInt(el.dataset.tabid, 10), el.dataset.origin);
    });
  });
}

async function switchEnvironment(tabId, baseUrl) {
  App.tabId          = tabId;
  App.baseUrl        = baseUrl;
  App.entityList     = [];
  App.filteredList   = [];
  App.selectedEntity = null;
  App.activeView     = null;
  App.metadata       = null;
  App.metadataDoc    = null;
  App._metadataDocSource = null;
  App.entityTypeCache = {};
  App._schemaForEntity = null;
  App._schemaTypeMap = null;
  App.lastQueryData  = null;
  App.displayCount   = 100;

  if (!App.tabId) {
    try {
      const tabs = await chrome.tabs.query({ url: `${baseUrl}/*` });
      if (tabs.length > 0) App.tabId = tabs[0].id;
    } catch (_) {}
  }

  if (!App.tabId) {
    showError('No open tab for this environment. Navigate to it in Edge first.');
    return;
  }

  await chrome.storage.session.set({ activeFscmTabId: App.tabId, activeFscmTabUrl: baseUrl });

  const envKey = baseUrl.replace('https://', '').split('.')[0];
  document.getElementById('envBadge').textContent = envKey + ' \u25be';

  // Reset UI
  document.getElementById('entityList').innerHTML   = '';
  document.getElementById('sidebarMeta').textContent = '';
  document.getElementById('loadMoreContainer').innerHTML = '';
  document.getElementById('detailPane').innerHTML = welcomeHTML();

  try { await loadEntityList(); } catch (e) { showError('Could not connect: ' + e.message); }
}

// ── Entity List ────────────────────────────────────────────

async function loadEntityList(forceRefresh = false) {
  const cacheKey = CACHE_KEY_PREFIX + App.baseUrl;

  if (!forceRefresh) {
    try {
      const stored = await chrome.storage.local.get(cacheKey);
      const entry  = stored[cacheKey];
      if (entry && (Date.now() - entry.ts) < CACHE_TTL_MS) {
        App.entityList   = entry.list;
        App.filteredList = [...App.entityList];
        hideLoading();
        renderEntityList(entry.ts);
        if (Date.now() - entry.ts > 60 * 60 * 1000) {
          fetchAndCacheEntityList(cacheKey, true);
        }
        return;
      }
    } catch (_) {}
  }

  await fetchAndCacheEntityList(cacheKey, false);
}

async function fetchAndCacheEntityList(cacheKey, silent) {
  if (!silent) {
    showLoading('Loading entity list (first load, please wait)\u2026');
  } else {
    const btn = document.getElementById('refreshEntities');
    if (btn) btn.innerHTML = '<span class="spinning">\u21bb</span>';
  }

  try {
    const result = await apiCall('/data');
    if (result.type !== 'json' || !Array.isArray(result.data?.value)) {
      throw new Error('Unexpected response from /data endpoint');
    }

    const list = result.data.value
      .filter(e => e.kind === 'EntitySet')
      .map(e => e.url || e.name)
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b));

    await chrome.storage.local.set({ [cacheKey]: { list, ts: Date.now() } });
    App.entityList   = list;
    App.filteredList = [...App.entityList];
    if (!silent) hideLoading();
    renderEntityList(Date.now());
  } catch (e) {
    if (!silent) { hideLoading(); throw e; }
  } finally {
    const btn = document.getElementById('refreshEntities');
    if (btn) btn.innerHTML = '\u21bb';
  }
}

// AOT table name ? OData entity keyword mappings
// Lets users type familiar table names like CustTable, VendTable, SalesTable
const AOT_ALIASES = {
  custtable:          ['customer'],
  vendtable:          ['vendor'],
  salestable:         ['salesorderheader', 'salesorder'],
  salesline:          ['salesorderline'],
  purchtable:         ['purchaseorderheader', 'purchaseorder'],
  purchline:          ['purchaseorderline'],
  inventtable:        ['releasedproduct', 'ecoresproduct'],
  inventtrans:        ['inventorytransaction'],
  inventdim:          ['inventorydimension'],
  inventsum:          ['inventoryon'],
  hcmworker:          ['worker'],
  dirpartytable:      ['globalpart', 'party'],
  ledgerjournaltable: ['ledgerjournalheader'],
  ledgerjournaltrans: ['ledgerjournalline'],
  custtrans:          ['customertransaction'],
  vendtrans:          ['vendortransaction'],
  banktable:          ['bankaccount'],
  taxgroup:           ['taxgroup'],
  invroutetable:      ['productionroute'],
  bomtable:           ['bom'],
  prodtable:          ['productionorder'],
  wmsorder:           ['warehouseorder'],
  wmslocation:        ['warehouselocation'],
  wmsinventcont:      ['licenseplat'],
  whsworktable:       ['warehousework'],
  retailtransaction:  ['retailtransaction'],
  ecorescategory:     ['procurementcategory', 'ecorescategory'],
  unitofmeasure:      ['unitofmeasuretranslation', 'unitofmeasure'],
  currtable:          ['currenc'],
  logisticsaddress:   ['logisticspostaladdress'],
  contactperson:      ['contact']
};

function resolveAliasKeywords(query) {
  const q = query.toLowerCase().replace(/[^a-z0-9]/g, '');
  return AOT_ALIASES[q] || null;
}

function filterEntityList(query) {
  App.lastSearchQuery = query;
  if (!query) {
    App.filteredList = [...App.entityList];
    App.aliasHint = null;
    renderEntityList();
    return;
  }

  const q = query.toLowerCase();

  // Direct name match first
  const direct = App.entityList.filter(n => n.toLowerCase().includes(q));

  if (direct.length > 0) {
    App.filteredList = direct;
    App.aliasHint = null;
  } else {
    // Try AOT alias expansion
    const keywords = resolveAliasKeywords(query);
    if (keywords) {
      App.filteredList = App.entityList.filter(n =>
        keywords.some(kw => n.toLowerCase().includes(kw))
      );
      App.aliasHint = keywords;
    } else {
      App.filteredList = [];
      App.aliasHint = null;
    }
  }

  renderEntityList();
}

function renderEntityList(cacheTs) {
  const listEl  = document.getElementById('entityList');
  const metaEl  = document.getElementById('sidebarMeta');
  const moreEl  = document.getElementById('loadMoreContainer');
  const total   = App.entityList.length;
  const filtered = App.filteredList.length;
  const visible  = App.filteredList.slice(0, App.displayCount);
  const hasMore  = filtered > App.displayCount;
  const searchVal = document.getElementById('searchInput').value.trim();
  const cachedAgo = cacheTs ? formatAge(Date.now() - cacheTs) : null;

  metaEl.innerHTML = total
    ? `<span>${filtered < total ? `${filtered} of ${total}` : total} entities</span>` +
      (cachedAgo && !searchVal ? `<span style="margin-left:auto;color:var(--text-subtle);font-size:11px">cached ${cachedAgo} ago</span>` : '')
    : '';

  // AOT alias banner
  const aliasEl = document.getElementById('aliasHint');
  if (aliasEl) {
    if (App.aliasHint && filtered > 0) {
      aliasEl.innerHTML = `<span class="alias-label">&#9432; AOT alias &rarr; showing entities matching: <em>${App.aliasHint.join(', ')}</em></span>`;
      aliasEl.classList.remove('hidden');
    } else {
      aliasEl.classList.add('hidden');
    }
  }

  if (filtered === 0) {
    const q = App.lastSearchQuery || '';
    const keywords = resolveAliasKeywords(q);
    listEl.innerHTML = `
      <div class="empty-state">
        No OData entities match <strong>${escHtml(q)}</strong>.<br><br>
        ${keywords ? `Tried alias expansion (<em>${keywords.join(', ')}</em>) � no results either.<br>` : ''}
        <small style="color:var(--text-subtle)">Note: AOT tables (CustTable, VendTable�) are exposed<br>as OData data entities with different names.</small>
      </div>`;
    moreEl.innerHTML = '';
    return;
  }

  listEl.innerHTML = visible.map(name => {
    const matchedAlias = App.aliasHint
      ? App.aliasHint.find(kw => name.toLowerCase().includes(kw))
      : null;
    const entityClass = classifyEntity(name);
    return `
    <div class="entity-item ${App.selectedEntity === name ? 'active' : ''}" data-name="${escHtml(name)}" title="${escHtml(name)}">
      <span class="entity-name">${escHtml(name)}</span>
      ${entityClass ? `<span class="entity-type-badge badge-${entityClass}" title="${entityClass === 'master' ? 'Master data entity' : 'Transactional entity'}">${entityClass === 'master' ? 'M' : 'T'}</span>` : ''}
      ${matchedAlias ? `<span class="alias-tag" title="Matched via AOT alias">AOT</span>` : ''}
      <div class="entity-actions">
        <button class="data-btn" data-name="${escHtml(name)}" title="Browse Data">\u25b6</button>
      </div>
    </div>
  `;
  }).join('');

  moreEl.innerHTML = hasMore
    ? `<button class="load-more-btn">Load ${Math.min(100, filtered - App.displayCount)} more\u2026 (${filtered - App.displayCount} remaining)</button>`
    : '';

  listEl.querySelectorAll('.entity-item').forEach(item => {
    item.addEventListener('click', () => showSchemaPanel(item.dataset.name));
  });
  listEl.querySelectorAll('.data-btn').forEach(btn => {
    btn.addEventListener('click', e => { e.stopPropagation(); showDataPanel(btn.dataset.name); });
  });
  moreEl.querySelector('.load-more-btn')?.addEventListener('click', () => {
    App.displayCount += 100;
    renderEntityList(cacheTs);
  });
}

// ── Detail Pane helpers ────────────────────────────────────

function welcomeHTML() {
  return `
    <div class="welcome-state">
      <div class="welcome-icon">&#9783;</div>
      <div class="welcome-title">F&amp;SCM Entity Browser</div>
      <div class="welcome-sub">Select an entity from the list on the left to view its fields and data</div>
    </div>`;
}

function buildDetailShell(entityName, activeTab) {
  return `
    <div class="detail-header">
      <span class="detail-entity-name" title="${escHtml(entityName)}">${escHtml(entityName)}</span>
      <button class="secondary-btn" id="detailSchemaBtn" ${activeTab === 'schema' ? 'style="display:none"' : ''}>View Schema</button>
      <button class="secondary-btn" id="detailDataBtn"   ${activeTab === 'data'   ? 'style="display:none"' : ''}>Browse Data \u25b6</button>
    </div>
    <div class="detail-tabs">
      <button class="detail-tab ${activeTab === 'schema' ? 'active' : ''}" data-view="schema">Schema</button>
      <button class="detail-tab ${activeTab === 'data'   ? 'active' : ''}" data-view="data">Data</button>
    </div>
    <div class="detail-body" id="detailBody"></div>
  `;
}

function bindDetailTabs(entityName) {
  document.getElementById('detailSchemaBtn')?.addEventListener('click', () => showSchemaPanel(entityName));
  document.getElementById('detailDataBtn')?.addEventListener('click',   () => showDataPanel(entityName));
  document.querySelectorAll('.detail-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      if (tab.dataset.view === 'schema') showSchemaPanel(entityName);
      else                               showDataPanel(entityName);
    });
  });
}

function highlightSidebar(entityName) {
  document.querySelectorAll('.entity-item').forEach(el =>
    el.classList.toggle('active', el.dataset.name === entityName)
  );
}

// ── Schema View ────────────────────────────────────────────

async function showSchemaPanel(entityName) {
  App.selectedEntity = entityName;
  App.activeView     = 'schema';
  highlightSidebar(entityName);

  const pane = document.getElementById('detailPane');
  pane.innerHTML = buildDetailShell(entityName, 'schema');
  bindDetailTabs(entityName);

  const body = document.getElementById('detailBody');

  // Check if metadata is already available (in-memory or cached)
  let metadataReady = !!App.metadata;
  if (!metadataReady) {
    try {
      const stored = await chrome.storage.local.get('metadata_' + App.baseUrl);
      const entry  = stored['metadata_' + App.baseUrl];
      if (entry && (Date.now() - entry.ts) < METADATA_CACHE_TTL_MS) {
        App.metadata  = entry.xml;
        metadataReady = true;
      }
    } catch (_) {}
  }

  if (!metadataReady) {
    // Show prompt instead of auto-downloading
    body.innerHTML = `
      <div class="schema-body">
        <div class="metadata-prompt">
          <div class="metadata-prompt-icon">&#8987;</div>
          <div class="metadata-prompt-title">Schema requires metadata download</div>
          <div class="metadata-prompt-desc">
            Downloading <code>$metadata</code> for <strong>${escHtml(App.baseUrl.replace('https://', '').split('.')[0])}</strong>
            takes 10�30 seconds. It will be cached for 7 days.
          </div>
          <button class="primary-btn" id="downloadMetaBtn">Download &amp; Show Schema</button>
          <div class="metadata-prompt-hint">You can also browse entity <strong>Data</strong> without downloading metadata.</div>
        </div>
      </div>`;

    document.getElementById('downloadMetaBtn').addEventListener('click', async () => {
      body.innerHTML = '<div class="schema-body"><div class="loading-inline">Loading schema\u2026</div></div>';
      await loadSchema(entityName, body);
    });
    return;
  }

  // Metadata already available � load immediately
  body.innerHTML = '<div class="schema-body"><div class="loading-inline">Loading schema\u2026</div></div>';
  await loadSchema(entityName, body);
}

async function loadSchema(entityName, body) {
  try {
    const et = await resolveEntityType(entityName);
    body.innerHTML = `<div class="schema-body">${buildSchemaHTML(et)}</div>`;
    body.querySelectorAll('.nav-link').forEach(link => {
      link.addEventListener('click', () => showSchemaPanel(link.dataset.nav));
    });
  } catch (e) {
    body.innerHTML = `<div class="error-inline">${escHtml(e.message)}</div>`;
  }
}

const METADATA_CACHE_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 days

async function resolveEntityType(entitySetName) {
  if (App.entityTypeCache[entitySetName]) return App.entityTypeCache[entitySetName];

  // Check in-memory first, then localStorage cache, then download
  if (!App.metadata) {
    try {
      const stored = await chrome.storage.local.get('metadata_' + App.baseUrl);
      const entry  = stored['metadata_' + App.baseUrl];
      if (entry && (Date.now() - entry.ts) < METADATA_CACHE_TTL_MS) {
        App.metadata = entry.xml;
        App.metadataDoc = null;
        App._metadataDocSource = null;
        App.entityTypeCache = {};
      }
    } catch (_) {}
  }
  if (!App.metadata) {
    showLoading('Downloading schema\u2026');
    try {
      const result = await apiCall('/data/$metadata');
      if (result.type !== 'xml') throw new Error('Expected XML from $metadata');
      App.metadata = result.data;
      App.metadataDoc = null;
      App._metadataDocSource = null;
      App.entityTypeCache = {};
      try {
        await chrome.storage.local.set({ ['metadata_' + App.baseUrl]: { xml: App.metadata, ts: Date.now() } });
      } catch (_) {}
    } finally {
      hideLoading();
    }
  }
  const parsed = parseEntityType(App.metadata, entitySetName);
  App.entityTypeCache[entitySetName] = parsed;
  return parsed;
}

function getMetadataDocument(xmlString) {
  if (App.metadataDoc && App._metadataDocSource === xmlString) {
    return App.metadataDoc;
  }
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, 'text/xml');
  if (doc.querySelector('parsererror')) {
    throw new Error('Could not parse $metadata XML. Try refreshing the entity list.');
  }
  App.metadataDoc = doc;
  App._metadataDocSource = xmlString;
  return doc;
}

// KEY FIX: use getElementsByTagNameNS('*', ...) to handle XML namespaces
function xmlAll(parent, localName) {
  const ns = parent.getElementsByTagNameNS('*', localName);
  return ns.length > 0 ? [...ns] : [...parent.getElementsByTagName(localName)];
}

function parseEntityType(xmlString, entitySetName) {
  const doc = getMetadataDocument(xmlString);

  // Find EntitySet to resolve the EntityType name
  let entityTypeName = entitySetName;
  const entitySetEl = xmlAll(doc, 'EntitySet')
    .find(e => e.getAttribute('Name') === entitySetName);

  if (entitySetEl) {
    const fullType = entitySetEl.getAttribute('EntityType') || '';
    entityTypeName = fullType.split('.').pop() || entitySetName;
  }

  // Find the EntityType element
  const entityTypeEl = xmlAll(doc, 'EntityType')
    .find(e => e.getAttribute('Name') === entityTypeName);

  if (!entityTypeEl) {
    throw new Error(`Schema not found for '${entitySetName}'. The entity may not expose full metadata.`);
  }

  // Key fields
  const keyRefs = xmlAll(entityTypeEl, 'PropertyRef');
  const keys    = new Set(keyRefs.map(k => k.getAttribute('Name')));

  // Build a map of external annotations: { PropName: labelString }
  // OData v4 CSDL uses <Annotations Target="NS.EntityTypeName/PropertyName"> elements
  // at the schema level rather than inline child elements of <Property>.
  const externalLabels = {};
  xmlAll(doc, 'Annotations').forEach(annotationsEl => {
    const target = annotationsEl.getAttribute('Target') || '';
    const slashIdx = target.lastIndexOf('/');
    if (slashIdx === -1) return;
    const typeRef  = target.substring(0, slashIdx);
    const propName = target.substring(slashIdx + 1);
    // match "Namespace.EntityTypeName" or just "EntityTypeName"
    if (!typeRef.endsWith(`.${entityTypeName}`) && typeRef !== entityTypeName) return;
    const displayAnnot = xmlAll(annotationsEl, 'Annotation').find(a => {
      const term = (a.getAttribute('Term') || '').toLowerCase();
      return term.includes('displayname') || term.includes('label') || term.includes('description');
    });
    if (displayAnnot) {
      const val = displayAnnot.getAttribute('String') || displayAnnot.getAttribute('string') || null;
      if (val) externalLabels[propName] = val;
    }
  });

  // Properties � inline annotation first, then external annotation
  const properties = xmlAll(entityTypeEl, 'Property')
    .filter(p => p.parentNode === entityTypeEl)
    .map(p => {
      const pName    = p.getAttribute('Name');
      const inlineEl = xmlAll(p, 'Annotation').find(a => {
        const term = (a.getAttribute('Term') || '').toLowerCase();
        return term.includes('displayname') || term.includes('label') || term.includes('description');
      });
      const label = inlineEl
        ? (inlineEl.getAttribute('String') || inlineEl.getAttribute('string') || null)
        : (externalLabels[pName] || null);
      return {
        name:      pName,
        label:     label,
        type:      (p.getAttribute('Type') || '').replace('Edm.', ''),
        nullable:  p.getAttribute('Nullable') !== 'false',
        maxLength: p.getAttribute('MaxLength') || null,
        precision: p.getAttribute('Precision') || null,
        scale:     p.getAttribute('Scale') || null,
        isKey:     keys.has(p.getAttribute('Name'))
      };
    });

  // Navigation properties
  const navProps = xmlAll(entityTypeEl, 'NavigationProperty')
    .map(n => {
      const targetType = (n.getAttribute('Type') || '')
        .replace('Collection(', '')
        .replace(')', '')
        .split('.')
        .pop();
      const targetEntitySet = xmlAll(doc, 'EntitySet').find(es => {
        const fullType = es.getAttribute('EntityType') || '';
        return fullType.endsWith(`.${targetType}`) || fullType === targetType;
      })?.getAttribute('Name') || null;
      return {
        name: n.getAttribute('Name'),
        type: targetType,
        target: targetEntitySet
      };
    });

  return { name: entitySetName, typeName: entityTypeName, keys: [...keys], properties, navProps };
}

function buildSchemaHTML(et) {
  const sizeHint = p => {
    if (p.maxLength)            return p.maxLength;
    if (p.precision && p.scale) return `${p.precision},${p.scale}`;
    if (p.precision)            return p.precision;
    return '\u2014';
  };

  return `
    <div class="schema-meta-row">
      <span><strong>Keys:</strong> ${et.keys.join(', ') || '\u2014'}</span>
      <span><strong>Fields:</strong> ${et.properties.length}</span>
      <span><strong>Navigation props:</strong> ${et.navProps.length}</span>
    </div>

    <div class="section-title">Fields</div>
    <table class="schema-table">
      <thead>
        <tr>
          <th>Field</th><th>Type</th><th>Max&nbsp;/&nbsp;Precision</th><th title="Required">Req</th>
        </tr>
      </thead>
      <tbody>
        ${et.properties.map(p => `
          <tr class="${p.isKey ? 'key-row' : ''}">
            <td><span class="field-name">${p.isKey ? '<span class="key-badge">KEY</span>' : ''}${escHtml(p.name)}</span></td>
            <td class="type-cell">${escHtml(p.type)}</td>
            <td class="meta-cell">${escHtml(sizeHint(p))}</td>
            <td class="meta-cell" title="${!p.nullable ? 'Required' : 'Optional'}">${!p.nullable ? '\u25cf' : '\u25cb'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>

    ${et.navProps.length > 0 ? `
      <div class="section-title">Navigation Properties</div>
      <table class="schema-table">
        <thead><tr><th>Name</th><th>Target Entity</th></tr></thead>
        <tbody>
          ${et.navProps.map(n => `
            <tr>
              <td>${n.target
                ? `<span class="field-name nav-link" data-nav="${escHtml(n.target)}">${escHtml(n.name)}</span>`
                : `<span class="field-name">${escHtml(n.name)}</span>`}</td>
              <td class="type-cell">${escHtml(n.type)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    ` : ''}
  `;
}

// ── Data View ──────────────────────────────────────────────

function showDataPanel(entityName) {
  App.selectedEntity      = entityName;
  App.activeView          = 'data';
  App.selectedLegalEntity = null;
  App.colFilters          = {};   // reset col-header filters for new entity
  App._schemaForEntity    = null;
  App._schemaTypeMap      = null;
  App.lastGridScroll      = { top: 0, left: 0 };
  App.filterRowSeq        = 0;
  App.lastFallbackInfo    = null;
  highlightSidebar(entityName);

  const pane = document.getElementById('detailPane');
  pane.innerHTML = buildDetailShell(entityName, 'data');
  bindDetailTabs(entityName);

  const body = document.getElementById('detailBody');
  body.innerHTML = `
    <datalist id="colOptionsList"></datalist>
    <div class="data-controls">
      <div class="data-action-row">
        <select id="legalEntitySelect" class="le-select" title="Filter by legal entity">
          <option value="">&#127970; All companies (cross-company)</option>
        </select>
        <button class="primary-btn" id="queryBtn">&#9654; Query</button>
        <button class="icon-action-btn" id="refreshBtn" title="Refresh (re-run last query)">&#8635;</button>
        <button class="icon-action-btn" id="exportBtn" title="Export CSV">&darr; CSV</button>
        <div class="template-wrap">
          <button class="icon-action-btn" id="templateBtn" title="Download import template">&#128203;&nbsp;Template&nbsp;&#9662;</button>
          <div class="template-dropdown hidden" id="templateDropdown">
            <button class="template-opt-btn" id="templateMandatoryBtn">&#9679;&nbsp;Mandatory fields only</button>
            <button class="template-opt-btn" id="templateAllBtn">&#9675;&nbsp;All fields</button>
          </div>
        </div>
      </div>
      <div class="filter-builder" id="filterBuilder">
        <div class="filter-rows" id="filterRows"></div>
        <div class="filter-footer">
          <button class="link-btn" id="addFilterBtn">+ Add filter</button>
          <button class="link-btn faded-btn" id="toggleAdvancedBtn">Advanced&hellip;</button>
        </div>
        <div class="raw-filter-row hidden" id="rawFilterRow">
          <span class="raw-filter-label">OData&nbsp;$filter:</span>
          <input type="text" id="filterInput" placeholder="e.g. Name eq 'test' and Amount gt 100" spellcheck="false">
        </div>
      </div>
      <div class="data-options-row">
        <label class="inline-label">Top <input type="number" id="topInput" value="100" min="1" max="5000"></label>
        <label class="inline-label">Skip <input type="number" id="skipInput" value="0" min="0"></label>
        <label class="inline-label">Mode
          <select id="filterModeSelect" class="mode-select" title="Strict keeps operators as-is. Compatible auto-downgrades unsupported fuzzy clauses.">
            <option value="compatible" ${App.filterMode === 'compatible' ? 'selected' : ''}>Compatible</option>
            <option value="strict" ${App.filterMode === 'strict' ? 'selected' : ''}>Strict</option>
          </select>
        </label>
      </div>
    </div>
    <div class="data-status-bar" id="dataStatusBar"></div>
    <div class="data-grid-area" id="dataGrid"></div>
    <div class="pagination-bar" id="pagination"></div>
  `;

  // Action row
  document.getElementById('queryBtn').addEventListener('click', queryData);
  document.getElementById('refreshBtn').addEventListener('click', queryData);
  document.getElementById('exportBtn').addEventListener('click', exportCSV);
  document.getElementById('filterModeSelect').addEventListener('change', e => {
    App.filterMode = e.target.value === 'strict' ? 'strict' : 'compatible';
    queryData();
  });

  // LE: auto-query on change
  let leDebounce;
  document.getElementById('legalEntitySelect').addEventListener('change', e => {
    App.selectedLegalEntity = e.target.value || null;
    clearTimeout(leDebounce);
    leDebounce = setTimeout(queryData, 150);
  });

  // Delegated listener for data-action buttons rendered inside detailBody
  body.addEventListener('click', e => {
    const action = e.target.closest('[data-action]')?.dataset.action;
    if (action === 'clearFilters') clearAllFilters();
    // template dropdown close
    if (!e.target.closest('.template-wrap')) {
      document.getElementById('templateDropdown')?.classList.add('hidden');
    }
  });

  // Template dropdown
  document.getElementById('templateBtn').addEventListener('click', e => {
    e.stopPropagation();
    document.getElementById('templateDropdown').classList.toggle('hidden');
  });
  document.getElementById('templateMandatoryBtn').addEventListener('click', () => downloadTemplate(true));
  document.getElementById('templateAllBtn').addEventListener('click', () => downloadTemplate(false));

  // Filter builder
  document.getElementById('addFilterBtn').addEventListener('click', () => addFilterRow());
  document.getElementById('toggleAdvancedBtn').addEventListener('click', () => {
    const rawRow = document.getElementById('rawFilterRow');
    const btn    = document.getElementById('toggleAdvancedBtn');
    rawRow.classList.toggle('hidden');
    btn.textContent = rawRow.classList.contains('hidden') ? 'Advanced\u2026' : 'Simple mode';
  });

  // Raw filter: debounced auto-query
  let rawDebounce;
  document.getElementById('filterInput').addEventListener('input', () => {
    clearTimeout(rawDebounce);
    rawDebounce = setTimeout(queryData, 600);
  });
  document.getElementById('filterInput').addEventListener('keydown', e => {
    if (e.key === 'Enter') { clearTimeout(rawDebounce); queryData(); }
  });

  // Populate LE dropdown then auto-query
  populateLegalEntityDropdown().then(() => queryData());
}

async function queryData() {
  if (!App.selectedEntity) return;

  App.queryInFlight = true;
  updateQueryPendingStatus();

  const legalEntity = document.getElementById('legalEntitySelect')?.value.trim() || '';
  const filterCtx   = buildODataFilter();
  const top         = Math.max(1, parseInt(document.getElementById('topInput')?.value) || 100);
  const skip        = Math.max(0, parseInt(document.getElementById('skipInput')?.value) || 0);

  // Build combined $filter: legal entity + filter builder
  const filterParts = [];
  if (legalEntity) filterParts.push(`dataAreaId eq '${legalEntity}'`);
  if (filterCtx.filter) filterParts.push(filterCtx.filter);
  const combinedFilter = filterParts.join(' and ');

  // ALWAYS use cross-company=true so D365 doesn't silently scope to default company
  let endpoint = `/data/${App.selectedEntity}?cross-company=true&$top=${top}&$skip=${skip}&$count=true`;
  if (combinedFilter) endpoint += `&$filter=${encodeURIComponent(combinedFilter)}`;

  const gridEl   = document.getElementById('dataGrid');
  const statusEl = document.getElementById('dataStatusBar');
  const pageEl   = document.getElementById('pagination');

  if (!gridEl) return;

  const prevScroll = { top: gridEl.scrollTop, left: gridEl.scrollLeft };
  App.lastGridScroll = prevScroll;
  clearFilterFallbackBadges();
  App.lastFallbackInfo = null;

  gridEl.innerHTML   = '<div class="loading-inline">Querying\u2026</div>';
  if (pageEl) pageEl.innerHTML = '';

  try {
    let result;
    let fallbackInfo = null;

    try {
      result = await apiCall(endpoint);
    } catch (e) {
      const msg = e.message || '';
      const canFallback = (msg.includes('400') || msg.includes('HTTP 400'))
        && App.filterMode === 'compatible'
        && filterCtx.hasDowngradableFuzzy;
      if (!canFallback) throw e;

      const fallbackResult = await tryCompatibleFallback({
        entityName: App.selectedEntity,
        legalEntity,
        top,
        skip,
        filterCtx
      });
      result = fallbackResult.result;
      fallbackInfo = fallbackResult.info;
    }

    if (result.type !== 'json') throw new Error('Non-JSON response');

    App.lastQueryData = result.data;
    const rows  = result.data.value || [];
    const count = result.data['@odata.count'];

    if (statusEl) {
      const fetchedAt = new Date().toLocaleTimeString();
      const rangeText = rows.length > 0
        ? `rows ${skip + 1}\u2013${skip + rows.length}`
        : 'rows 0\u20130';
      statusEl.innerHTML = `
        <span class="data-entity-label">${escHtml(App.selectedEntity)}</span>
        ${count != null ? `<span class="record-count">${Number(count).toLocaleString()} total</span>` : ''}
        <span class="showing-count">${rangeText}</span>
        ${buildFallbackSummaryHTML(fallbackInfo)}
        <span class="fetch-time" title="Last fetched">&#128337; ${fetchedAt}</span>
        ${hasActiveFilters() ? '<button class="clear-filters-btn" data-action="clearFilters">&#10005; Clear filters</button>' : ''}
      `;
      updateQueryPendingStatus();
    }

    if (fallbackInfo) {
      App.lastFallbackInfo = fallbackInfo;
      applyFilterFallbackBadges(fallbackInfo.downgradedRowIds || []);
    }

    const cols = rows.length > 0 ? Object.keys(rows[0]).filter(k => !k.startsWith('@')) : [];
    // Populate labels BEFORE renderDataGrid so column headers pick them up immediately
    populateColLabels(App.selectedEntity, cols);
    if (cols.length > 0) updateColDatalist(cols);
    renderDataGrid(rows, gridEl, prevScroll);
    renderPagination(skip, top, count, rows.length, pageEl);
  } catch (e) {
    const msg = e.message || '';
    let friendlyMsg = 'Query failed: ' + msg;
    if (msg.includes('400') || msg.includes('HTTP 400')) {
      if (msg.includes('TargetInvocationException') || msg.includes('An error has occurred')) {
        friendlyMsg = 'This entity does not support the filter you used. '
          + 'Try using <strong>eq</strong> (exact match) instead of contains/startswith, '
          + 'or query without filters and export to CSV.';
      } else {
        friendlyMsg = 'Bad request (HTTP 400) — check your filter syntax. ' + msg;
      }
    }
    if (gridEl) gridEl.innerHTML = `<div class="error-inline">${friendlyMsg}${hasActiveFilters() ? '<br><button class="link-btn" data-action="clearFilters">&#10005; Clear all filters and retry</button>' : ''}</div>`;
    showError('Query failed — see details in the data panel.');
  } finally {
    App.queryInFlight = false;
    App.pendingFilterQuery = false;
    updateQueryPendingStatus();
  }
}

function buildEndpoint(entityName, top, skip, legalEntity, filter) {
  const parts = [];
  if (legalEntity) parts.push(`dataAreaId eq '${legalEntity}'`);
  if (filter) parts.push(filter);
  const combined = parts.join(' and ');
  let endpoint = `/data/${entityName}?cross-company=true&$top=${top}&$skip=${skip}&$count=true`;
  if (combined) endpoint += `&$filter=${encodeURIComponent(combined)}`;
  return endpoint;
}

function parseSmartStringFilterInput(rawValue) {
  const raw = String(rawValue || '').trim();
  if (!raw) return null;

  if (raw.startsWith('=')) {
    const exact = raw.slice(1).trim();
    if (!exact) return null;
    return { valueForExpr: exact, valueForEq: exact, fuzzy: false, mode: 'exact' };
  }
  if (raw.startsWith('^')) {
    const start = raw.slice(1).trim();
    if (!start) return null;
    return { valueForExpr: `${start}*`, valueForEq: start, fuzzy: true, mode: 'startswith' };
  }
  if (raw.endsWith('$')) {
    const end = raw.slice(0, -1).trim();
    if (!end) return null;
    return { valueForExpr: `*${end}`, valueForEq: end, fuzzy: true, mode: 'endswith' };
  }
  if (raw.includes('*')) {
    return { valueForExpr: raw, valueForEq: raw.replace(/\*/g, ''), fuzzy: true, mode: 'pattern' };
  }
  return { valueForExpr: `*${raw}*`, valueForEq: raw, fuzzy: true, mode: 'contains' };
}

function parseComparableFilterInput(rawValue, type) {
  const raw = String(rawValue || '').trim();
  if (!raw) return null;

  const m = raw.match(/^(<=|>=|<|>|=)?\s*(.+)$/);
  if (!m) return null;

  const op = m[1] || '=';
  const value = (m[2] || '').trim();
  if (!value) return null;

  if (type === 'number') {
    if (!/^[-+]?\d+(\.\d+)?$/.test(value)) return null;
    return { op, value };
  }

  if (type === 'datetime') {
    // Accept common F&O date/date-time forms; backend validates exact type-specific literal.
    if (!/^\d{4}-\d{2}-\d{2}(?:[T\s].+)?$/.test(value)) return null;
    return { op, value };
  }

  return null;
}

function clausesToFilter(clauses) {
  return clauses.map(c => c.currentExpr || c.expr).filter(Boolean).join(' and ');
}

async function tryCompatibleFallback({ entityName, legalEntity, top, skip, filterCtx }) {
  const fuzzyClauses = filterCtx.clauses.filter(c => c.isFuzzy && c.canDowngrade);
  const maxProbe = 8;

  // Step 1: downgrade all fuzzy-capable clauses to ensure a safe baseline.
  const working = filterCtx.clauses.map(c => ({
    ...c,
    currentExpr: (c.isFuzzy && c.canDowngrade) ? c.eqExpr : c.expr,
    downgraded: !!(c.isFuzzy && c.canDowngrade)
  }));

  let result = await apiCall(buildEndpoint(entityName, top, skip, legalEntity, clausesToFilter(working)));

  // Step 2: restore fuzzy clause-by-clause where backend accepts it.
  const probes = fuzzyClauses.slice(0, maxProbe);
  for (const probe of probes) {
    const trial = working.map(c => ({ ...c }));
    const idx = trial.findIndex(c => c.id === probe.id);
    if (idx === -1) continue;
    trial[idx].currentExpr = trial[idx].expr;
    trial[idx].downgraded = false;
    try {
      const trialResult = await apiCall(buildEndpoint(entityName, top, skip, legalEntity, clausesToFilter(trial)));
      // Restore succeeded, keep it.
      working[idx] = trial[idx];
      result = trialResult;
    } catch (_) {
      // Keep downgraded for this clause.
    }
  }

  const downgradedRows = working
    .filter(c => c.source === 'row' && c.downgraded && c.rowId)
    .map(c => c.rowId);

  return {
    result,
    info: {
      downgradedCount: working.filter(c => c.downgraded).length,
      downgradedRowIds: [...new Set(downgradedRows)],
      testedClauses: probes.length,
      totalFuzzyClauses: fuzzyClauses.length,
      probeLimited: fuzzyClauses.length > maxProbe
    }
  };
}

function buildFallbackSummaryHTML(info) {
  if (!info) return '';
  if (info.downgradedCount <= 0) {
    return '<span class="filter-fallback-summary ok">fuzzy supported</span>';
  }
  const probeNote = info.probeLimited ? ` (tested first ${info.testedClauses})` : '';
  return `<span class="filter-fallback-summary" title="compatible mode fallback">${info.downgradedCount} clause${info.downgradedCount > 1 ? 's' : ''} downgraded${probeNote}</span>`;
}

function clearFilterFallbackBadges() {
  document.querySelectorAll('#filterRows .filter-row').forEach(row => {
    row.classList.remove('is-fallback');
    const badge = row.querySelector('.filter-fallback-badge');
    if (badge) {
      badge.textContent = '';
      badge.classList.add('hidden');
    }
  });
}

function applyFilterFallbackBadges(rowIds) {
  if (!rowIds || rowIds.length === 0) return;
  rowIds.forEach(id => {
    const row = document.querySelector(`#filterRows .filter-row[data-row-id="${CSS.escape(id)}"]`);
    if (!row) return;
    row.classList.add('is-fallback');
    const badge = row.querySelector('.filter-fallback-badge');
    if (!badge) return;
    badge.textContent = 'EQ fallback';
    badge.classList.remove('hidden');
    badge.title = 'This fuzzy clause was downgraded to exact match in Compatible mode.';
  });
}

function renderDataGrid(rows, gridEl, restoreScroll = { top: 0, left: 0 }) {
  if (!rows || rows.length === 0) {
    const msg = hasActiveFilters()
      ? `<div class="empty-state">No records match your filters.<br><button class="link-btn" data-action="clearFilters">&#10005; Clear all filters and re-query</button></div>`
      : '<div class="empty-state">No records returned for this query.</div>';
    gridEl.innerHTML = msg;
    return;
  }

  const cols = Object.keys(rows[0]).filter(k => !k.startsWith('@'));

  // Store raw rows on the grid element for client-side column filtering
  gridEl._rawRows = rows;
  gridEl._cols    = cols;
  gridEl._renderToken = (gridEl._renderToken || 0) + 1;
  const renderToken = gridEl._renderToken;

  gridEl.innerHTML = `
    <table class="data-table" id="dataTableEl">
      <thead>
        <tr>${cols.map(c => {
          const label = App._colLabels?.[c];
          return label && label !== c
            ? `<th title="${escHtml(c)}"><span class="col-label">${escHtml(label)}</span><span class="col-name-sub">${escHtml(c)}</span></th>`
            : `<th title="${escHtml(c)}">${escHtml(c)}</th>`;
        }).join('')}</tr>
        <tr class="col-filter-row">
          ${cols.map(c => {
            const hint = getColFilterHint(c);
            return `
            <th class="col-filter-cell">
              <input class="col-filter-input" data-col="${escHtml(c)}"
                     placeholder="${escHtml(hint.placeholder)}"
                     title="${escHtml(hint.title)}"
                     spellcheck="false"
                     value="${escHtml(App.colFilters[c] || '')}">
            </th>`;
          }).join('')}
        </tr>
      </thead>
      <tbody id="dataTableBody"></tbody>
    </table>
  `;

  renderTableBody(rows, cols, gridEl, renderToken, restoreScroll);

  // Wire col-filter inputs: update App.colFilters (source of truth), then re-query server-side
  let debounceTimer;
  gridEl.querySelectorAll('.col-filter-input').forEach(input => {
    input.addEventListener('input', () => {
      const col = input.dataset.col;
      const val = input.value.trim();
      App.colFilters[col] = val;
      // clear empty entries
      if (!App.colFilters[col]) delete App.colFilters[col];

      clearTimeout(debounceTimer);

      const colType = detectColType(col);
      if (!shouldAutoRunColumnFilter(val, colType)) {
        App.pendingFilterQuery = false;
        updateQueryPendingStatus();
        return;
      }

      App.pendingFilterQuery = true;
      updateQueryPendingStatus();
      debounceTimer = setTimeout(() => {
        const skipEl = document.getElementById('skipInput');
        if (skipEl) skipEl.value = '0';
        queryData();
      }, COLUMN_FILTER_DEBOUNCE_MS);
    });
    input.addEventListener('keydown', e => {
      if (e.key === 'Escape') {
        input.value = '';
        delete App.colFilters[input.dataset.col];
        clearTimeout(debounceTimer);
        App.pendingFilterQuery = true;
        updateQueryPendingStatus();
        queryData();
      }
      if (e.key === 'Enter') {
        clearTimeout(debounceTimer);
        App.pendingFilterQuery = true;
        updateQueryPendingStatus();
        queryData();
      }
    });
  });
}

function renderTableBody(rows, cols, gridEl, renderToken, restoreScroll = { top: 0, left: 0 }) {
  const tbody = gridEl.querySelector('#dataTableBody');
  if (!tbody) return;

  if (rows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="${cols.length}" class="null-val" style="text-align:center;padding:20px">No rows match column filters</td></tr>`;
    return;
  }

  const buildRowHtml = row => `
    <tr>
      ${cols.map(c => {
        const raw = row[c];
        if (raw === null || raw === undefined || raw === '') return '<td class="null-val">null</td>';
        const val = String(raw);
        return `<td title="${escHtml(val)}">${escHtml(truncate(val, 60))}</td>`;
      }).join('')}
    </tr>
  `;

  if (rows.length <= 250) {
    tbody.innerHTML = rows.map(buildRowHtml).join('');
    requestAnimationFrame(() => restoreGridScroll(gridEl, renderToken, restoreScroll));
    return;
  }

  // Render large tables in chunks to keep typing and scrolling responsive.
  tbody.innerHTML = '';
  const chunkSize = 100;
  let index = 0;

  const renderChunk = () => {
    if (gridEl._renderToken !== renderToken) return;
    const slice = rows.slice(index, index + chunkSize);
    tbody.insertAdjacentHTML('beforeend', slice.map(buildRowHtml).join(''));
    if (index === 0) restoreGridScroll(gridEl, renderToken, restoreScroll);
    index += chunkSize;
    if (index < rows.length) requestAnimationFrame(renderChunk);
    else restoreGridScroll(gridEl, renderToken, restoreScroll);
  };

  requestAnimationFrame(renderChunk);
}

function getColFilterHint(col) {
  const type = detectColType(col);
  switch (type) {
    case 'number':
      return {
        placeholder: '100 | >100 | <=250',
        title: `Numeric column: supports =, >, >=, <, <= (examples: 100, >100, <=250). Auto-search runs after ${COLUMN_FILTER_DEBOUNCE_MS}ms pause; press Enter for immediate search.`
      };
    case 'boolean':
      return {
        placeholder: 'true or false',
        title: `Boolean column: use true/false (or 1/0). Auto-search runs after ${COLUMN_FILTER_DEBOUNCE_MS}ms pause; press Enter for immediate search.`
      };
    case 'datetime':
      return {
        placeholder: '2025-01-15 | >=2025-01-01',
        title: `Date/time column: supports =, >, >=, <, <= with YYYY-MM-DD or ISO timestamp (example: >=2025-01-01). Auto-search runs after ${COLUMN_FILTER_DEBOUNCE_MS}ms pause; press Enter for immediate search.`
      };
    default:
      return {
        placeholder: 'text | =exact | ^start | end$ | *pattern*',
        title: `Text-like column: plain text=contains, =exact, ^starts with, end$, or explicit *wildcard*. Auto-search runs after ${COLUMN_FILTER_DEBOUNCE_MS}ms pause; press Enter for immediate search.`
      };
  }
}

function shouldAutoRunColumnFilter(value, type) {
  const raw = String(value || '').trim();
  if (!raw) return true; // clearing filter should rerun query

  if (type === 'number') {
    return !!parseComparableFilterInput(raw, 'number');
  }

  if (type === 'boolean') {
    const normalized = (raw.startsWith('=') ? raw.slice(1) : raw).trim().toLowerCase();
    return ['true', 'false', '1', '0'].includes(normalized);
  }

  if (type === 'datetime') {
    return !!parseComparableFilterInput(raw, 'datetime');
  }

  // For text-like filters, avoid firing on 1-character probes unless explicit syntax is used.
  const normalized = raw.replace(/^=/, '').replace(/^\^/, '').replace(/\$$/, '').replace(/\*/g, '').trim();
  return normalized.length >= 2 || raw.includes('*') || raw.startsWith('=') || raw.startsWith('^') || raw.endsWith('$');
}

function restoreGridScroll(gridEl, renderToken, scrollPos) {
  if (gridEl._renderToken !== renderToken) return;
  gridEl.scrollTop = Math.max(0, scrollPos?.top || 0);
  gridEl.scrollLeft = Math.max(0, scrollPos?.left || 0);
}

function updateQueryPendingStatus() {
  const statusEl = document.getElementById('dataStatusBar');
  if (!statusEl) return;

  const isPending = App.pendingFilterQuery || App.queryInFlight;
  const existing = statusEl.querySelector('.query-pending');

  if (isPending && !existing) {
    statusEl.insertAdjacentHTML('beforeend', '<span class="query-pending">Querying...</span>');
    return;
  }
  if (!isPending && existing) {
    existing.remove();
  }
}

function applyColumnFilters(gridEl) {
  const rows = gridEl._rawRows;
  const cols = gridEl._cols;
  if (!rows || !cols) return;

  // Collect active column filters
  const filters = {};
  gridEl.querySelectorAll('.col-filter-input').forEach(input => {
    const val = input.value.trim();
    if (val) filters[input.dataset.col] = val.toLowerCase();
  });

  const filtered = Object.keys(filters).length === 0
    ? rows
    : rows.filter(row =>
        Object.entries(filters).every(([col, term]) => {
          const cell = row[col];
          if (cell === null || cell === undefined) return false;
          return String(cell).toLowerCase().includes(term);
        })
      );

  // Update match count in status bar
  const statusEl = document.getElementById('dataStatusBar');
  const activeFilters = Object.keys(filters).length;
  if (statusEl && activeFilters > 0) {
    const existing = statusEl.querySelector('.col-filter-status');
    const msg = `� <strong>${filtered.length}</strong> matching column filter${activeFilters > 1 ? 's' : ''}`;
    if (existing) existing.innerHTML = msg;
    else statusEl.insertAdjacentHTML('beforeend', `<span class="col-filter-status">${msg}</span>`);
  } else {
    statusEl?.querySelector('.col-filter-status')?.remove();
  }

  renderTableBody(filtered, cols, gridEl);
}

function renderPagination(skip, top, count, rowCount, pageEl) {
  if (!pageEl || count == null || Number(count) <= top) {
    if (pageEl) pageEl.innerHTML = '';
    return;
  }

  const total   = Number(count);
  const page    = Math.floor(skip / top) + 1;
  const pages   = Math.ceil(total / top);
  const hasPrev = skip > 0;
  const hasNext = skip + rowCount < total;

  pageEl.innerHTML = `
    <button class="page-btn" id="prevBtn" ${hasPrev ? '' : 'disabled'}>\u2190 Prev</button>
    <span>Page ${page} / ${pages}</span>
    <button class="page-btn" id="nextBtn" ${hasNext ? '' : 'disabled'}>Next \u2192</button>
  `;

  pageEl.querySelector('#prevBtn')?.addEventListener('click', () => {
    document.getElementById('skipInput').value = Math.max(0, skip - top);
    queryData();
  });
  pageEl.querySelector('#nextBtn')?.addEventListener('click', () => {
    document.getElementById('skipInput').value = skip + top;
    queryData();
  });
}

// ── CSV Export ─────────────────────────────────────────────


// -- Filter Builder ---------------------------------------------------------

function buildODataFilter() {
  const disableFuzzy = !!arguments[0]?.disableFuzzy;
  const parts = [];
  const clauses = [];
  let hasFuzzy = false;
  document.querySelectorAll('#filterRows .filter-row').forEach((row, idx) => {
    const rowId = row.dataset.rowId || String(idx + 1);
    const col = row.querySelector('.filter-col-input')?.value.trim() || '';
    const op  = row.querySelector('.filter-op-select')?.value || 'eq';
    const val = row.querySelector('.filter-val-input')?.value.trim() ?? '';
    if (!col) return;
    const esc  = val.replace(/'/g, "''");
    const type = detectColType(col);
    // For eq/ne, omit quotes for non-string types so D365 gets the right literal
    const eqLiteral = (v) => {
      if (type === 'number')  { const n = Number(v); return isNaN(n) ? `'${v.replace(/'/g,"''")}'` : String(n); }
      if (type === 'boolean') { const b = v.toLowerCase(); return (b === 'true' || b === '1') ? 'true' : 'false'; }
      if (type === 'datetime') return v;
      return `'${v.replace(/'/g, "''")}'`;
    };
    switch (op) {
      case 'eq': {
        const expr = `${col} eq ${eqLiteral(val)}`;
        parts.push(expr);
        clauses.push({ id: `row:${rowId}`, source: 'row', rowId, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        break;
      }
      case 'ne': {
        const expr = `${col} ne ${eqLiteral(val)}`;
        parts.push(expr);
        clauses.push({ id: `row:${rowId}`, source: 'row', rowId, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        break;
      }
      case 'contains':
        hasFuzzy = true;
        {
          const expr = `${col} eq '*${esc}*'`;
          const eqExpr = `${col} eq ${eqLiteral(val)}`;
          const currentExpr = disableFuzzy ? eqExpr : expr;
          parts.push(currentExpr);
          clauses.push({ id: `row:${rowId}`, source: 'row', rowId, isFuzzy: true, canDowngrade: true, expr, eqExpr });
        }
        break;
      case 'startswith':
        hasFuzzy = true;
        {
          const expr = `${col} eq '${esc}*'`;
          const eqExpr = `${col} eq ${eqLiteral(val)}`;
          const currentExpr = disableFuzzy ? eqExpr : expr;
          parts.push(currentExpr);
          clauses.push({ id: `row:${rowId}`, source: 'row', rowId, isFuzzy: true, canDowngrade: true, expr, eqExpr });
        }
        break;
      case 'gt': {
        const expr = `${col} gt ${val}`;
        parts.push(expr);
        clauses.push({ id: `row:${rowId}`, source: 'row', rowId, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        break;
      }
      case 'lt': {
        const expr = `${col} lt ${val}`;
        parts.push(expr);
        clauses.push({ id: `row:${rowId}`, source: 'row', rowId, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        break;
      }
      case 'ge': {
        const expr = `${col} ge ${val}`;
        parts.push(expr);
        clauses.push({ id: `row:${rowId}`, source: 'row', rowId, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        break;
      }
      case 'le': {
        const expr = `${col} le ${val}`;
        parts.push(expr);
        clauses.push({ id: `row:${rowId}`, source: 'row', rowId, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        break;
      }
      case 'null': {
        const expr = `${col} eq null`;
        parts.push(expr);
        clauses.push({ id: `row:${rowId}`, source: 'row', rowId, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        break;
      }
      case 'notnull': {
        const expr = `${col} ne null`;
        parts.push(expr);
        clauses.push({ id: `row:${rowId}`, source: 'row', rowId, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        break;
      }
    }
  });
  // Column-header search inputs � use eq (exact match) � universally safe in D365 OData
  Object.entries(App.colFilters).forEach(([col, val]) => {
    if (!val) return;
    const type = detectColType(col);
    switch (type) {
      case 'number': {
        const parsed = parseComparableFilterInput(val, 'number');
        if (parsed) {
          const op = parsed.op === '=' ? 'eq' : parsed.op === '>' ? 'gt' : parsed.op === '<' ? 'lt' : parsed.op === '>=' ? 'ge' : 'le';
          const expr = `${col} ${op} ${parsed.value}`;
          parts.push(expr);
          clauses.push({ id: `col:${col}`, source: 'col', rowId: null, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        }
        break;
      }
      case 'boolean': {
        const b = (val.startsWith('=') ? val.slice(1) : val).trim().toLowerCase();
        if (b === 'true' || b === '1') {
          const expr = `${col} eq true`;
          parts.push(expr);
          clauses.push({ id: `col:${col}`, source: 'col', rowId: null, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        } else if (b === 'false' || b === '0') {
          const expr = `${col} eq false`;
          parts.push(expr);
          clauses.push({ id: `col:${col}`, source: 'col', rowId: null, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        }
        break;
      }
      case 'datetime': {
        const parsed = parseComparableFilterInput(val, 'datetime');
        if (parsed) {
          const op = parsed.op === '=' ? 'eq' : parsed.op === '>' ? 'gt' : parsed.op === '<' ? 'lt' : parsed.op === '>=' ? 'ge' : 'le';
          const expr = `${col} ${op} ${parsed.value}`;
          parts.push(expr);
          clauses.push({ id: `col:${col}`, source: 'col', rowId: null, isFuzzy: false, canDowngrade: false, expr, eqExpr: expr });
        }
        break;
      }
      default:
        // String/other fields: smart syntax with wildcard-eq default for contains behavior.
        {
          const parsed = parseSmartStringFilterInput(val);
          if (!parsed) break;
          const exprVal = parsed.valueForExpr.replace(/'/g, "''");
          const eqVal   = parsed.valueForEq.replace(/'/g, "''");
          const expr = `${col} eq '${exprVal}'`;
          const eqExpr = `${col} eq '${eqVal}'`;
          const currentExpr = (disableFuzzy && parsed.fuzzy) ? eqExpr : expr;
          if (parsed.fuzzy) hasFuzzy = true;
          parts.push(currentExpr);
          clauses.push({
            id: `col:${col}`,
            source: 'col',
            rowId: null,
            isFuzzy: parsed.fuzzy,
            canDowngrade: parsed.fuzzy,
            expr,
            eqExpr
          });
        }
        break;
    }
  });
  const rawRow = document.getElementById('rawFilterRow');
  const rawVal = document.getElementById('filterInput')?.value.trim() || '';
  if (rawRow && !rawRow.classList.contains('hidden') && rawVal) {
    parts.push(rawVal);
    const rawHasFuzzy = /(contains\s*\(|startswith\s*\()/i.test(rawVal);
    if (rawHasFuzzy) hasFuzzy = true;
    clauses.push({ id: 'raw:1', source: 'raw', rowId: null, isFuzzy: rawHasFuzzy, canDowngrade: false, expr: rawVal, eqExpr: rawVal });
  }
  return {
    filter: parts.join(' and '),
    clauses,
    hasFuzzy,
    hasDowngradableFuzzy: clauses.some(c => c.isFuzzy && c.canDowngrade)
  };
}

function hasActiveFilters() {
  if (Object.keys(App.colFilters).length > 0) return true;
  const rows = document.querySelectorAll('#filterRows .filter-row');
  if (rows.length > 0) return true;
  const rawRow = document.getElementById('rawFilterRow');
  const rawVal = document.getElementById('filterInput')?.value.trim();
  if (rawRow && !rawRow.classList.contains('hidden') && rawVal) return true;
  return false;
}

function clearAllFilters() {
  App.colFilters = {};
  const fr = document.getElementById('filterRows');
  if (fr) fr.innerHTML = '';
  const fi = document.getElementById('filterInput');
  if (fi) fi.value = '';
  clearFilterFallbackBadges();
  document.querySelectorAll('.col-filter-input').forEach(i => { i.value = ''; });
  const skipEl = document.getElementById('skipInput');
  if (skipEl) skipEl.value = '0';
  queryData();
}

function detectColType(col) {
  // 1. Check cached schema metadata for the definitive OData type
  if (App.metadata && App.selectedEntity) {
    try {
      if (App._schemaForEntity !== App.selectedEntity || !App._schemaTypeMap) {
        const et = App.entityTypeCache[App.selectedEntity] || parseEntityType(App.metadata, App.selectedEntity);
        if (!App.entityTypeCache[App.selectedEntity]) App.entityTypeCache[App.selectedEntity] = et;
        App._schemaForEntity = App.selectedEntity;
        App._schemaTypeMap = {};
        et.properties.forEach(p => { App._schemaTypeMap[p.name] = p.type; });
      }
      const propType = App._schemaTypeMap[col];
      if (propType) {
        const t = propType.toLowerCase();
        if (t === 'string')                          return 'string';
        if (['int16','int32','int64','decimal','double','single','byte'].includes(t)) return 'number';
        if (t === 'boolean')                         return 'boolean';
        if (t.includes('date') || t.includes('time')) return 'datetime';
        return 'other'; // guid, enum, binary, etc. � contains() will fail
      }
    } catch (_) {}
  }
  // 2. Fallback: infer from JS value in current data
  const rows = App.lastQueryData?.value;
  if (!rows?.length) return 'unknown';
  const sample = rows.find(r => r[col] !== null && r[col] !== undefined && r[col] !== '');
  if (!sample) return 'unknown';
  const v = sample[col];
  if (typeof v === 'number')  return 'number';
  if (typeof v === 'boolean') return 'boolean';
  if (typeof v === 'string') {
    if (/^\d{4}-\d{2}-\d{2}/.test(v))              return 'datetime';
    if (/^[0-9a-f]{8}-[0-9a-f]{4}-/i.test(v))      return 'other'; // guid
    return 'string';
  }
  return 'unknown';
}

function addFilterRow() {
  const container = document.getElementById('filterRows');
  if (!container) return;

  const rowId = String(++App.filterRowSeq);

  const row = document.createElement('div');
  row.className = 'filter-row';
  row.dataset.rowId = rowId;
  row.innerHTML = `
    <input type="text" class="filter-col-input" list="colOptionsList"
           placeholder="column" spellcheck="false" autocomplete="off">
    <select class="filter-op-select">
      <option value="eq">= equals</option>
      <option value="ne">&ne; not equals</option>
      <option value="contains">&#8715; contains</option>
      <option value="startswith">starts&nbsp;with</option>
      <option value="gt">&gt; greater&nbsp;than</option>
      <option value="lt">&lt; less&nbsp;than</option>
      <option value="ge">&ge;</option>
      <option value="le">&le;</option>
      <option value="null">is null</option>
      <option value="notnull">is not null</option>
    </select>
    <input type="text" class="filter-val-input" placeholder="value" spellcheck="false">
    <span class="filter-fallback-badge hidden"></span>
    <button class="filter-remove-btn" title="Remove">&#215;</button>
  `;
  container.appendChild(row);

  const colInput = row.querySelector('.filter-col-input');
  const opSel    = row.querySelector('.filter-op-select');
  const valInput = row.querySelector('.filter-val-input');

  let debounce;
  const scheduleQuery = () => { clearTimeout(debounce); debounce = setTimeout(queryData, 600); };

  opSel.addEventListener('change', () => {
    valInput.style.visibility = ['null', 'notnull'].includes(opSel.value) ? 'hidden' : '';
    scheduleQuery();
  });
  row.querySelector('.filter-remove-btn').addEventListener('click', () => { row.remove(); queryData(); });
  colInput.addEventListener('change', scheduleQuery);
  valInput.addEventListener('input', scheduleQuery);
  valInput.addEventListener('keydown', e => { if (e.key === 'Enter') { clearTimeout(debounce); queryData(); } });

  colInput.focus();
}

function updateColDatalist(cols) {
  const dl = document.getElementById('colOptionsList');
  if (!dl) return;
  dl.innerHTML = cols.map(c => `<option value="${escHtml(c)}">`).join('');
}

function splitPascalCase(name) {
  return name
    .replace(/^_+/, '')                        // strip leading underscores
    .replace(/([a-z])([A-Z])/g, '$1 $2')       // camelCase boundary
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2') // consecutive caps (e.g. XMLParser)
    .replace(/_/g, ' ')                         // remaining underscores
    .replace(/^[a-z]/, c => c.toUpperCase())    // capitalise first letter
    .trim();
}

function populateColLabels(entityName, cols) {
  App._colLabels = {};
  if (!cols || cols.length === 0) return;

  // Parse metadata ONCE (not once per column � metadata XML can be very large)
  let schemaProps = null;
  if (App.metadata && entityName) {
    try {
      schemaProps = parseEntityType(App.metadata, entityName).properties;
    } catch (_) {}
  }

  cols.forEach(col => {
    if (schemaProps) {
      const prop = schemaProps.find(p => p.name === col);
      if (prop?.label) { App._colLabels[col] = prop.label; return; }
    }
    // Fallback: split PascalCase/camelCase property name into readable words
    App._colLabels[col] = splitPascalCase(col);
  });
  console.log('[EntityBrowser] populateColLabels:', cols.length, 'cols, sample:', App._colLabels[cols[0]]);
}

async function downloadTemplate(mandatoryOnly) {
  document.getElementById('templateDropdown')?.classList.add('hidden');
  if (!App.selectedEntity) return;
  try {
    showLoading('Loading schema\u2026');
    const et = await resolveEntityType(App.selectedEntity);
    hideLoading();

    // For all-fields: mandatory columns first, then optional
    let props;
    if (mandatoryOnly) {
      props = et.properties.filter(p => !p.nullable);
    } else {
      const mandatory = et.properties.filter(p => !p.nullable);
      const optional  = et.properties.filter(p =>  p.nullable);
      props = [...mandatory, ...optional];
    }
    if (!props.length) { showError('No fields found in schema.'); return; }

    // Build Excel SpreadsheetML (opens natively in Excel, no library needed)
    const xmlRows = props.map(p => {
      const displayName = p.label || p.name;
      const fieldName   = p.name;
      const required    = !p.nullable;
      return [
        `        <Row>`,
        `          <Cell ss:StyleID="${required ? 'mandatory' : 'optional'}"><Data ss:Type="String">${xmlEsc(displayName)}</Data></Cell>`,
        `          <Cell ss:StyleID="sub"><Data ss:Type="String">${xmlEsc(fieldName)}</Data></Cell>`,
        `          <Cell ss:StyleID="sub"><Data ss:Type="String">${xmlEsc(p.type)}</Data></Cell>`,
        `          <Cell ss:StyleID="sub"><Data ss:Type="String">${required ? 'Required' : 'Optional'}</Data></Cell>`,
        `        </Row>`,
      ].join('\n');
    }).join('\n');

    const xml = [
      `<?xml version="1.0" encoding="UTF-8"?>`,
      `<?mso-application progid="Excel.Sheet"?>`,
      `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"`,
      `  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"`,
      `  xmlns:x="urn:schemas-microsoft-com:office:excel">`,
      `  <Styles>`,
      `    <Style ss:ID="header"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#0078D4" ss:Pattern="Solid"/></Style>`,
      `    <Style ss:ID="mandatory"><Font ss:Bold="1"/><Interior ss:Color="#FFF4CE" ss:Pattern="Solid"/></Style>`,
      `    <Style ss:ID="optional"></Style>`,
      `    <Style ss:ID="sub"><Font ss:Color="#605E5C" ss:Size="9"/></Style>`,
      `  </Styles>`,
      `  <Worksheet ss:Name="${xmlEsc(App.selectedEntity)}">`,
      `    <Table>`,
      `      <Column ss:Width="180"/>`,
      `      <Column ss:Width="140"/>`,
      `      <Column ss:Width="80"/>`,
      `      <Column ss:Width="70"/>`,
      `      <Row>`,
      `        <Cell ss:StyleID="header"><Data ss:Type="String">Display Name</Data></Cell>`,
      `        <Cell ss:StyleID="header"><Data ss:Type="String">Field Name (API)</Data></Cell>`,
      `        <Cell ss:StyleID="header"><Data ss:Type="String">Type</Data></Cell>`,
      `        <Cell ss:StyleID="header"><Data ss:Type="String">Required</Data></Cell>`,
      `      </Row>`,
      xmlRows,
      `    </Table>`,
      `  </Worksheet>`,
      `  <Worksheet ss:Name="Import Data">`,
      `    <Table>`,
      `      <Row>`,
      props.map(p => `        <Cell ss:StyleID="${!p.nullable ? 'mandatory' : 'optional'}"><Data ss:Type="String">${xmlEsc(p.name)}</Data></Cell>`).join('\n'),
      `      </Row>`,
      `    </Table>`,
      `  </Worksheet>`,
      `</Workbook>`,
    ].join('\n');

    const blob = new Blob([xml], { type: 'application/vnd.ms-excel;charset=utf-8;' });
    const a    = document.createElement('a');
    a.href     = URL.createObjectURL(blob);
    a.download = `${App.selectedEntity}_template_${mandatoryOnly ? 'mandatory' : 'all'}.xls`;
    a.click();
    URL.revokeObjectURL(a.href);
  } catch (e) {
    hideLoading();
    showError('Template failed: ' + e.message);
  }
}

// -- Legal Entity Helpers ---------------------------------------------------

async function loadLegalEntities() {
  const cacheKey = 'legalEntities_' + App.baseUrl;
  try {
    const stored = await chrome.storage.local.get(cacheKey);
    const entry  = stored[cacheKey];
    if (entry && (Date.now() - entry.ts) < 24 * 60 * 60 * 1000) return entry.list;
  } catch (_) {}

  const candidates = [
    { entity: 'LegalEntities',     idField: 'LegalEntityId', nameField: 'Name' },
    { entity: 'DataAreaEntity',    idField: 'DataArea',      nameField: 'Name' },
    { entity: 'DataAreas',         idField: 'DataAreaId',    nameField: 'Name' },
    { entity: 'CompanyInfoEntity', idField: 'DataArea',      nameField: 'Name' },
  ];

  for (const { entity, idField, nameField } of candidates) {
    try {
      const url = '/data/' + entity + '?$select=' + idField + ',' + nameField + '&$top=200';
      const result = await apiCall(url);
      if (result.type === 'json' && Array.isArray(result.data?.value) && result.data.value.length > 0) {
        const sample        = result.data.value[0];
        const actualIdField = Object.keys(sample).find(k => k.toLowerCase() === idField.toLowerCase());
        const actualNmField = Object.keys(sample).find(k => k.toLowerCase() === nameField.toLowerCase());
        if (!actualIdField) continue;
        const list = result.data.value
          .map(r => ({ id: r[actualIdField], name: actualNmField ? r[actualNmField] : '' }))
          .filter(r => r.id)
          .sort((a, b) => a.id.localeCompare(b.id));
        if (list.length > 0) {
          try { await chrome.storage.local.set({ [cacheKey]: { list, ts: Date.now() } }); } catch (_) {}
          return list;
        }
      }
    } catch (_) { /* try next candidate */ }
  }
  return [];
}

async function populateLegalEntityDropdown() {
  const sel = document.getElementById('legalEntitySelect');
  if (!sel) return;
  try {
    const list = await loadLegalEntities();
    const existing = new Set([...sel.options].map(o => o.value));
    list.forEach(({ id, name }) => {
      if (!existing.has(id)) {
        const opt = document.createElement('option');
        opt.value = id;
        opt.text  = name ? id + ' � ' + name : id;
        sel.appendChild(opt);
      }
    });
  } catch (_) { /* silent */ }
}

function exportCSV() {
  const rows = App.lastQueryData?.value;
  if (!rows?.length) { showError('No data to export. Run a query first.'); return; }

  const cols = Object.keys(rows[0]).filter(k => !k.startsWith('@'));
  const csv  = [
    cols.map(csvEsc).join(','),
    ...rows.map(row => cols.map(c => csvEsc(String(row[c] ?? ''))).join(','))
  ].join('\r\n');

  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
  a.download = `${App.selectedEntity}_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ── API proxy (via content script) ────────────────────────

async function apiCall(endpoint, method = 'GET', body = null) {
  const url = endpoint.startsWith('http') ? endpoint : `${App.baseUrl}${endpoint}`;

  try {
    return await sendToContentScript(url, method, body);
  } catch (err) {
    const isConnErr = err.message?.includes('Receiving end does not exist') ||
                      err.message?.includes('Could not establish connection');
    if (!isConnErr) throw err;
  }

  // Auto-inject content script and retry
  try {
    await chrome.scripting.executeScript({
      target: { tabId: App.tabId },
      files:  ['src/content.js']
    });
    await new Promise(r => setTimeout(r, 300));
  } catch (injectErr) {
    throw new Error('Could not inject into tab. Make sure the F\u0026SCM URL is open in Edge. (' + injectErr.message + ')');
  }

  return sendToContentScript(url, method, body);
}

function sendToContentScript(url, method, body) {
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(App.tabId, { type: 'FSCM_API_REQUEST', url, method, body }, (response) => {
      if (chrome.runtime.lastError) { reject(new Error(chrome.runtime.lastError.message)); return; }
      if (!response)                { reject(new Error('No response from page')); return; }
      if (!response.success)        { reject(new Error(response.error || 'API call failed')); return; }
      resolve(response);
    });
  });
}

// ── UI helpers ─────────────────────────────────────────────

function showLoading(msg) {
  document.getElementById('statusMessage').textContent = msg || 'Loading\u2026';
  document.getElementById('statusOverlay').classList.remove('hidden');
}
function hideLoading() {
  document.getElementById('statusOverlay').classList.add('hidden');
}
function showError(msg) {
  document.getElementById('errorText').textContent = msg;
  const bar = document.getElementById('errorBar');
  bar.classList.remove('hidden');
  setTimeout(() => bar.classList.add('hidden'), 10_000);
}

function renderNoTab() {
  document.getElementById('entityList').innerHTML = `
    <div class="no-tab-msg">
      <div class="no-tab-icon">&#127760;</div>
      <p>Open a D365 F&amp;SCM page in Edge, then click the extension icon.</p>
      <p class="hint">Supported: <code>*.operations.dynamics.com</code></p>
    </div>
  `;
}

// ── Utilities ──────────────────────────────────────────────

function formatAge(ms) {
  const s = Math.floor(ms / 1000);
  if (s < 60)  return `${s}s`;
  const m = Math.floor(s / 60);
  if (m < 60)  return `${m}m`;
  const h = Math.floor(m / 60);
  if (h < 24)  return `${h}h`;
  return `${Math.floor(h / 24)}d`;
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function xmlEsc(str) {
  return String(str || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function classifyEntity(name) {
  const n = name.toLowerCase();
  const txPatterns = ['trans','journal','ledger','invoice','order','line','header','posting',
    'movement','requisition','request','quotation','receipt','transfer','pick','pack','ship',
    'voucher','subledger','settlement','payment','remittance','purch','salesline','purchline'];
  const masterPatterns = ['customer','vendor','item','product','employee','worker','asset',
    'resource','group','setup','param','config','address','contact','category','price',
    'attribute','dimension','warehouse','location','table','master','reference','code'];
  if (txPatterns.some(p => n.includes(p))) return 'transaction';
  if (masterPatterns.some(p => n.includes(p))) return 'master';
  return null;
}

function truncate(str, max) {
  return str.length > max ? str.slice(0, max) + '\u2026' : str;
}

function csvEsc(val) {
  if (val.includes(',') || val.includes('"') || val.includes('\n')) {
    return `"${val.replace(/"/g, '""')}"`;
  }
  return val;
}

