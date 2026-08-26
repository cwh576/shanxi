(function () {
  const STORAGE_KEY = 'shanxi-electric-dashboard-v2';
  const DATA_STORAGE_KEY = `${STORAGE_KEY}-data`;
  const UI_STORAGE_KEY = `${STORAGE_KEY}-ui`;
  const defaults = window.APP_DATA || {};
  const state = {
    tab: 'longTerm',
    ui: {
      activeTab: 'longTerm',
      longMonths: (defaults.longTermWeightedClearing?.months || []).slice(-1).map((m) => m.month),
      longMonthOpen: false,
      agentMonthOpen: false,
      agentMonths: (defaults.agentPurchase?.months || []).slice(-1).map((m) => m.month),
      voltageLevelId: '',
      loadUserOpen: false,
      longK: 0,
      longSplit: false,
      agentMonth: defaults.agentPurchase?.month || '2026-07',
      agentSplit: false,
      splitMonth: defaults.marketSplit?.month || defaults.agentPurchase?.month || '2026-07',
      splitVoltageLevelId: '',
      agentImagePreview: [],
      splitImagePreview: [],
      loadUsers: [],
      loadDateFrom: '',
      loadDateTo: '',
      loadMode: '96',
      loadLegendHidden: [],
      longLegendHidden: [],
      agentLegendHidden: [],
      importStatus: '',
      importDetail: '',
      importPreview: '',
      loadManageUser: '',
      loadManageDateFrom: '',
      loadManageDateTo: '',
      loadManageMode: '96',
      compareMonth: (defaults.longTermWeightedClearing?.months || []).slice(-1)[0]?.month || defaults.agentPurchase?.month || '2026-07',
      compareSplit: false,
      compareK: 0,
      compareLoadMode: '24',
      compareRows: []
    },
    data: clone(defaults),
    settingsOpen: false,
    settingsUnlocked: false,
    settingsTab: 'load'
  };
  const cache = {
    loadRecordsRef: null,
    loadRecordsLength: -1,
    loadAggregated: [],
    splitMonthsRef: null,
    splitCacheKey: '',
    splitRows: []
  };
  let renderQueued = false;
  let persistTimer = null;
  let persistDataTimer = null;
  let storageWarned = false;
  let lastExportUrl = '';
  let ocrWorkerPromise = null;
  let splitMonthDragIndex = null;
  let remoteSaveTimer = null;
  let remoteSyncing = false;
  let remoteReady = false;
  let remoteStatus = '正在连接云端数据…';
  let remoteAccessToken = readRemoteAccessToken();
  const TABLE_PREVIEW_LIMIT = 1000;
  const CHART_SERIES_LIMIT = 80;

  hydrate();
  ensureDataShape();
  ensureSelections();
  persistStartupMigration();
  bindGlobalEvents();
  window.addEventListener('beforeunload', flushPersist);
  render();
  syncRemoteData();

  function clone(v) {
    return JSON.parse(JSON.stringify(v || {}));
  }

  function hydrate() {
    const raw = localStorage.getItem(STORAGE_KEY);
    try {
      if (raw) {
        const saved = JSON.parse(raw);
        if (saved.data) state.data = merge(state.data, saved.data);
        if (saved.ui) state.ui = { ...state.ui, ...saved.ui };
      }
      const rawData = localStorage.getItem(DATA_STORAGE_KEY);
      const rawUi = localStorage.getItem(UI_STORAGE_KEY);
      if (rawData) state.data = merge(state.data, JSON.parse(rawData));
      if (rawUi) state.ui = { ...state.ui, ...JSON.parse(rawUi) };
      if (['longTerm', 'agent', 'split', 'load', 'compare'].includes(state.ui.activeTab)) state.tab = state.ui.activeTab;
    } catch (_) {}
  }

  function merge(target, source) {
    if (Array.isArray(target) && Array.isArray(source)) return source;
    if (!target || !source || typeof target !== 'object' || typeof source !== 'object') return source === undefined ? target : source;
    const out = Array.isArray(target) ? [] : { ...target };
    for (const key of Object.keys(source)) out[key] = merge(target[key], source[key]);
    return out;
  }

  function persist() {
    clearTimeout(persistTimer);
    persistTimer = setTimeout(persistNow, 250);
  }

  function persistNow() {
    safeSetItem(UI_STORAGE_KEY, JSON.stringify(state.ui));
  }

  function persistData() {
    clearTimeout(persistTimer);
    clearTimeout(persistDataTimer);
    persistDataTimer = setTimeout(persistDataNow, 250);
  }

  function persistDataNow() {
    safeSetItem(UI_STORAGE_KEY, JSON.stringify(state.ui));
    safeSetItem(DATA_STORAGE_KEY, JSON.stringify(state.data));
    queueRemoteSave();
  }

  function flushPersist() {
    clearTimeout(persistTimer);
    clearTimeout(persistDataTimer);
    persistNow();
    if (persistDataTimer) persistDataNow();
  }

  function safeSetItem(key, value) {
    try {
      localStorage.setItem(key, value);
    } catch (err) {
      if (!storageWarned) {
        storageWarned = true;
        alert('数据量过大，浏览器本地存储空间不足。本次页面仍可查看和导出，建议及时导出Excel备份。');
      }
      console.warn(err);
    }
  }

  function supabaseConfig() {
    const config = window.SUPABASE_CONFIG || {};
    if (!config.url || !config.publishableKey || !config.table || !config.rowId) return null;
    return config;
  }

  function readRemoteAccessToken() {
    try {
      return sessionStorage.getItem('shanxi-supabase-access-token') || '';
    } catch (_) {
      return '';
    }
  }

  function writeRemoteAccessToken(value) {
    try {
      if (value) sessionStorage.setItem('shanxi-supabase-access-token', value);
      else sessionStorage.removeItem('shanxi-supabase-access-token');
    } catch (_) {}
  }

  function remoteHeaders(withJson = false) {
    const config = supabaseConfig();
    const token = remoteAccessToken || config?.publishableKey || '';
    return {
      apikey: config?.publishableKey || '',
      Authorization: `Bearer ${token}`,
      ...(withJson ? { 'Content-Type': 'application/json' } : {})
    };
  }

  async function supabaseFetch(path, options = {}) {
    const config = supabaseConfig();
    if (!config) throw new Error('未配置 Supabase');
    const response = await fetch(`${config.url}${path}`, {
      ...options,
      headers: { ...remoteHeaders(Boolean(options.body)), ...(options.headers || {}) }
    });
    if (!response.ok) {
      const detail = await response.text();
      throw new Error(detail || `云端请求失败（${response.status}）`);
    }
    const text = await response.text();
    if (!text.trim()) return null;
    try {
      return JSON.parse(text);
    } catch (err) {
      throw new Error(`云端返回了无法解析的数据：${err?.message || err}`);
    }
  }

  async function syncRemoteData() {
    const config = supabaseConfig();
    if (!config) {
      remoteReady = true;
      remoteStatus = '未配置云端数据';
      return;
    }
    remoteSyncing = true;
    remoteStatus = '正在读取云端最新数据…';
    scheduleRender();
    try {
      const rows = await supabaseFetch(`/rest/v1/${encodeURIComponent(config.table)}?id=eq.${encodeURIComponent(config.rowId)}&select=id,payload,updated_at`);
      const row = Array.isArray(rows) ? rows[0] : null;
      if (row?.payload && typeof row.payload === 'object') {
        state.data = merge(state.data, row.payload);
        ensureDataShape();
        ensureSelections();
        invalidateDataCache();
        remoteStatus = `已同步云端数据${row.updated_at ? `（${formatRemoteTime(row.updated_at)}）` : ''}`;
      } else {
        remoteStatus = '云端还没有共享数据，当前使用内置数据';
      }
      remoteReady = true;
    } catch (err) {
      remoteReady = true;
      remoteStatus = `云端读取失败：${err?.message || err}`;
      console.warn(err);
    } finally {
      remoteSyncing = false;
      scheduleRender();
    }
  }

  function queueRemoteSave() {
    if (!remoteReady || remoteSyncing) return;
    clearTimeout(remoteSaveTimer);
    remoteSaveTimer = setTimeout(() => saveRemoteData().catch((err) => {
      remoteStatus = `云端保存失败：${err?.message || err}`;
      scheduleRender();
    }), 800);
  }

  async function saveRemoteData() {
    const config = supabaseConfig();
    if (!config) return;
    if (!remoteAccessToken) {
      remoteStatus = '本地已保存；请在设置中登录云端后同步';
      scheduleRender();
      return;
    }
    remoteSyncing = true;
    remoteStatus = '正在保存到云端…';
    scheduleRender();
    try {
      await supabaseFetch(`/rest/v1/${encodeURIComponent(config.table)}?on_conflict=id`, {
        method: 'POST',
        headers: { Prefer: 'resolution=merge-duplicates,return=minimal' },
        body: JSON.stringify({
          id: config.rowId,
          payload: state.data,
          updated_at: new Date().toISOString()
        })
      });
      remoteStatus = `云端已更新（${formatRemoteTime(new Date().toISOString())}）`;
    } finally {
      remoteSyncing = false;
      scheduleRender();
    }
  }

  function formatRemoteTime(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value || '');
    return date.toLocaleString('zh-CN', { hour12: false });
  }

  async function loginRemote() {
    const email = byId('remoteEmail')?.value.trim();
    const password = byId('remotePassword')?.value || '';
    if (!email || !password) {
      alert('请输入云端账号和密码');
      return;
    }
    const config = supabaseConfig();
    if (!config) return;
    const response = await fetch(`${config.url}/auth/v1/token?grant_type=password`, {
      method: 'POST',
      headers: { apikey: config.publishableKey, 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password })
    });
    const result = await response.json();
    if (!response.ok || !result.access_token) throw new Error(result.error_description || result.msg || '云端登录失败');
    remoteAccessToken = result.access_token;
    writeRemoteAccessToken(remoteAccessToken);
    remoteStatus = `云端已登录：${email}`;
    scheduleRender();
  }

  function logoutRemote() {
    remoteAccessToken = '';
    writeRemoteAccessToken('');
    remoteStatus = '已退出云端登录';
    scheduleRender();
  }

  function scheduleRender() {
    if (renderQueued) return;
    renderQueued = true;
    requestAnimationFrame(() => {
      renderQueued = false;
      render();
    });
  }

  function invalidateDataCache() {
    cache.loadRecordsRef = null;
    cache.loadRecordsLength = -1;
    cache.splitMonthsRef = null;
    cache.splitCacheKey = '';
  }

  function bindGlobalEvents() {
    document.addEventListener('click', async (e) => {
      const openFileBtn = e.target.closest?.('[data-open-file]');
      if (openFileBtn) {
        e.preventDefault();
        const input = byId(openFileBtn.dataset.openFile);
        if (input) {
          setImportFeedback('已打开文件选择器', '选择图片或Excel后会自动识别。', '导入后这里会显示识别原文、解析结果和失败原因。');
          if (typeof input.showPicker === 'function') {
            input.showPicker();
          } else {
            input.click();
          }
        }
        return;
      }
      const settingsTabBtn = e.target.closest?.('[data-settings-tab]');
      if (settingsTabBtn) {
        e.preventDefault();
        if (state.settingsOpen && state.settingsUnlocked) {
          state.settingsTab = settingsTabBtn.dataset.settingsTab;
          scheduleRender();
        }
        return;
      }
      const loadModeBtn = e.target.closest?.('[data-load-manage-mode]');
      if (loadModeBtn) {
        e.preventDefault();
        state.ui.loadManageMode = loadModeBtn.dataset.loadManageMode;
        persist();
        scheduleRender();
        return;
      }
      if (e.target.closest?.('#exportDataBtn')) {
        e.preventDefault();
        await runSettingsExport();
        return;
      }
      const exportBtn = e.target.closest?.('[data-export]');
      if (exportBtn) {
        e.preventDefault();
        e.stopPropagation();
        await runExportButton(exportBtn);
        return;
      }
      if (e.target.closest?.('#closeSettingsBtn')) {
        e.preventDefault();
        closeSettings();
        return;
      }
      if (e.target.closest?.('#unlockSettingsBtn')) {
        e.preventDefault();
        unlockSettings();
        return;
      }
      const backdrop = e.target.closest?.('#settingsBackdrop');
      if (backdrop && e.target === backdrop) closeSettings();
    });
    document.addEventListener('change', (e) => {
      const t = e.target;
      if (t?.matches?.('#importLoadFile')) return handleImport(t);
      if (t?.matches?.('#importLongFile')) return handleImportLong(t);
      if (t?.matches?.('#importAgentImageFile')) return handleImportAgentImage(t);
      if (t?.matches?.('#importAgentExcelFile')) return handleImportAgentExcel(t);
      if (t?.matches?.('#importSplitImageFile')) return handleImportSplitImage(t);
      if (t?.matches?.('#importSplitExcelFile')) return handleImportSplitExcel(t);
      if (t?.matches?.('#compareLoadFile')) return handleCompareImport(t);
      if (!state.settingsOpen || !state.settingsUnlocked) return;
      if (!t || !t.matches) return;
      if (t.matches('#loadManageUser')) {
        state.ui.loadManageUser = t.value || '';
        persist();
        scheduleRender();
      } else if (t.matches('#loadManageDateFrom')) {
        state.ui.loadManageDateFrom = t.value || '';
        persist();
        scheduleRender();
      } else if (t.matches('#loadManageDateTo')) {
        state.ui.loadManageDateTo = t.value || '';
        persist();
        scheduleRender();
      }
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && state.settingsOpen) closeSettings();
      if (e.key === 'Enter' && e.target?.id === 'settingsPwd') unlockSettings();
    });
  }

  async function runExportButton(btn) {
    if (!btn || btn.disabled) return;
    btn.disabled = true;
    const oldText = btn.textContent;
    btn.textContent = '正在导出...';
    try {
      await exportTable(btn.dataset.export, btn);
    } catch (err) {
      alert('导出失败：' + (err?.message || err));
    } finally {
      btn.disabled = false;
      btn.textContent = oldText;
    }
  }

  async function runSettingsExport() {
    const btn = byId('exportDataBtn');
    if (!btn || btn.disabled) return;
    btn.disabled = true;
    const oldText = btn.textContent;
    btn.textContent = '正在导出...';
    try {
      await exportData();
    } catch (err) {
      alert('导出失败：' + (err?.message || err));
    } finally {
      btn.disabled = false;
      btn.textContent = oldText;
    }
  }

  function unlockSettings() {
    const pwd = byId('settingsPwd');
    if (!pwd) return;
    if (pwd.value.trim() === '741852') {
      state.settingsUnlocked = true;
      render();
    } else {
      alert('密码不对');
    }
  }

  function ensureSelections() {
    const months = state.data.longTermWeightedClearing?.months || [];
    state.ui.longMonths = (state.ui.longMonths || []).filter((m) => months.some((x) => x.month === m));
    if (!('longMonths' in state.ui) && months.length) state.ui.longMonths = [months[months.length - 1].month];
    const agentMonths = state.data.agentPurchase?.months || [];
    state.ui.agentMonths = (state.ui.agentMonths || []).filter((m) => agentMonths.some((x) => x.month === m));
    if (!('agentMonths' in state.ui) && agentMonths.length) state.ui.agentMonths = [agentMonths[agentMonths.length - 1].month];
    if (!state.ui.compareMonth && months.length) state.ui.compareMonth = months[months.length - 1].month;
    if (!state.ui.splitMonth && months.length) state.ui.splitMonth = months[months.length - 1].month;
    const splitOptions = splitVoltageLevelOptions();
    const selectedSplit = findVoltageOptionStrict(splitOptions, state.ui.splitVoltageLevelId) || findVoltageOptionStrict(splitOptions, state.ui.voltageLevelId);
    if (selectedSplit) state.ui.splitVoltageLevelId = selectedSplit.id;
    if (!state.ui.splitVoltageLevelId && splitOptions.length) state.ui.splitVoltageLevelId = splitOptions[0].id;
    const voltageOptions = voltageLevelOptions();
    const selectedVoltage = findVoltageOptionStrict(voltageOptions, state.ui.voltageLevelId);
    if (selectedVoltage) state.ui.voltageLevelId = selectedVoltage.id;
    if (!state.ui.voltageLevelId && voltageOptions.length) state.ui.voltageLevelId = voltageOptions[0].id;
    const loadRows = aggregateLoad(state.data.userLoad?.records || []);
    const loadUsers = [...new Set(loadRows.map((r) => r.userName).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'));
    if (state.ui.loadManageUser && !loadUsers.includes(state.ui.loadManageUser)) state.ui.loadManageUser = '';
  }

  function persistStartupMigration() {
    const version = '20260826-voltage-levels-3';
    if (state.ui.storageSchemaVersion === version) return;
    state.ui.storageSchemaVersion = version;
    persistDataNow();
  }

  function setImportFeedback(status = '', detail = '', preview = '') {
    state.ui.importStatus = status;
    state.ui.importDetail = detail;
    state.ui.importPreview = preview;
  }

  function summarizeOcrPreview(text, limit = 20) {
    const lines = String(text || '').split(/\n+/).map((s) => s.trim()).filter(Boolean);
    if (!lines.length) return '';
    return lines.slice(0, limit).join('\n');
  }

  function buildImportPreview({ fileName = '', kind = '', rawText = '', parsed = null, error = '' }) {
    const parts = [];
    if (fileName) parts.push(`文件：${fileName}`);
    if (kind) parts.push(`类型：${kind}`);
    if (error) parts.push(`结果：${error}`);
    if (parsed) parts.push(describeParsedImport(parsed));
    if (rawText) {
      parts.push('原始识别：');
      parts.push(summarizeOcrPreview(rawText, 40) || '（空）');
    }
    return parts.join('\n');
  }

  function describeParsedImport(parsed) {
    if (!parsed || typeof parsed !== 'object') return '';
    if (Object.prototype.hasOwnProperty.call(parsed, 'flatPriceKwh')) {
      return [
        `月份：${parsed.month || ''}`,
        `平段价：${formatNumber(parsed.flatPriceKwh)}`,
        `当月平均：${formatNumber(parsed.averagePurchaseKwh)}`,
        `历史偏差：${formatNumber(parsed.historyDeviationKwh)}`,
        `电压等级：${(parsed.voltageLevels || []).length} 条`,
        ...(parsed.voltageLevels || []).map((v) => `${v.label || ''}｜线损 ${formatNumber(v.lineLossKwh)}｜输配 ${formatNumber(v.transmissionKwh)}｜基金 ${formatNumber(v.fundKwh)}｜系统 ${formatNumber(v.systemKwh)}`)
      ].filter(Boolean).join('\n');
    }
    if (Object.prototype.hasOwnProperty.call(parsed, 'voltageLevels')) {
      return [
        `月份：${parsed.month || ''}`,
        `电压等级：${(parsed.voltageLevels || []).length} 条`,
        ...(parsed.voltageLevels || []).map((v) => `${v.label || ''}｜线损 ${formatNumber(v.lineLossKwh)}｜输配 ${formatNumber(v.transmissionKwh)}｜基金 ${formatNumber(v.fundKwh)}｜系统 ${formatNumber(v.systemKwh)}`)
      ].filter(Boolean).join('\n');
    }
    return JSON.stringify(parsed, null, 2);
  }

  function withTimeout(promise, ms, label) {
    let timer = null;
    const timeout = new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error(label || `超时 ${ms}ms`)), ms);
    });
    return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
  }

  function ensureDataShape() {
    const agent = state.data.agentPurchase || {};
    if (!agent.months) agent.months = [{ month: agent.month || '2026-07', flatPriceKwh: agent.flatPriceKwh || 0, averagePurchaseKwh: agent.averagePurchaseKwh || 0, historyDeviationKwh: agent.historyDeviationKwh || 0 }];
    if (!agent.voltageLevels && state.data.marketSplit?.voltageLevels) agent.voltageLevels = clone(state.data.marketSplit.voltageLevels);
    (agent.months || []).forEach((item) => normalizeAgentMonthData(item));
    state.data.agentPurchase = agent;
    const split = state.data.marketSplit || {};
    if (Array.isArray(split.voltageLevels)) split.voltageLevels = normalizeSplitVoltageLevels(split.voltageLevels);
    if (!split.months) split.months = [{ month: split.month || agent.month || '2026-07', voltageLevels: clone(split.voltageLevels || agent.voltageLevels || []) }];
    split.months = (split.months || []).map((item) => ({
      ...item,
      voltageLevels: normalizeSplitVoltageLevels(item?.voltageLevels || split.voltageLevels || agent.voltageLevels || [])
    }));
    state.data.marketSplit = split;
  }

  function byId(id) { return document.getElementById(id); }
  function esc(v) { return String(v ?? '').replace(/[&<>]/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[m])); }
  function fmt(v, d = 3) { return v === null || v === undefined || Number.isNaN(v) ? '—' : Number(v).toFixed(d); }
  function round(v, d = 3) { const p = 10 ** d; return Math.round(v * p) / p; }
  function monthLabel(m) { const [y, mm] = String(m).split('-'); return `${y}年${Number(mm)}月`; }
  function monthIndex(m) { return Number(String(m).split('-')[1] || 0) - 1; }
  function normalizeVoltageKey(value, { omitType = false } = {}) {
    let s = String(value ?? '').trim().toLowerCase();
    s = s
      .replace(/\s+/g, '')
      .replace(/两部制|dual|two/g, 'dual')
      .replace(/单一制|single/g, 'single')
      .replace(/不满|以下|under/g, 'under')
      .replace(/及以上|以上|up/g, 'up')
      .replace(/千伏|kv/g, '')
      .replace(/[_-]\d+$/g, '')
      .replace(/[^a-z0-9\u4e00-\u9fa5]+/g, '');
    if (omitType) s = s.replace(/^(dual|single)/, '');
    return s;
  }
  function voltageFullKey(v) { return normalizeVoltageKey(v?.label || v?.name || v?.voltage || v?.id || v); }
  function voltageBandKey(v) { return normalizeVoltageKey(v?.label || v?.name || v?.voltage || v?.id || v, { omitType: true }); }
  function findVoltageOption(options, ref) {
    const list = options || [];
    if (!list.length) return null;
    const refId = typeof ref === 'object' ? String(ref?.id || '') : String(ref || '');
    const refLabel = typeof ref === 'object' ? String(ref?.label || '') : '';
    const exact = list.find((v) => String(v.id || '') === refId) || (refLabel ? list.find((v) => String(v.label || '') === refLabel) : null);
    if (exact) return exact;
    const fullKey = voltageFullKey(ref);
    if (fullKey) {
      const full = list.find((v) => voltageFullKey(v) === fullKey || normalizeVoltageKey(v.id) === fullKey);
      if (full) return full;
    }
    const bandKey = voltageBandKey(ref);
    if (bandKey) {
      const bandMatches = list.filter((v) => voltageBandKey(v) === bandKey || normalizeVoltageKey(v.id, { omitType: true }) === bandKey);
      if (bandMatches.length === 1) return bandMatches[0];
    }
    return null;
  }
  function findVoltageOptionStrict(options, ref) {
    const list = options || [];
    if (!list.length) return null;
    const refId = typeof ref === 'object' ? String(ref?.id || '') : String(ref || '');
    const refLabel = typeof ref === 'object' ? String(ref?.label || '') : '';
    const exact = list.find((v) => String(v.id || '') === refId) || (refLabel ? list.find((v) => String(v.label || '') === refLabel) : null);
    if (exact) return exact;
    const fullKey = voltageFullKey(ref);
    if (!fullKey) return null;
    return list.find((v) => voltageFullKey(v) === fullKey || normalizeVoltageKey(v.id) === fullKey) || null;
  }
  function currentVoltage(id) { return findVoltageOptionStrict(voltageLevelOptions(), id); }
  function yuanKwhToMwh(v) { return Number(v || 0) * 1000; }
  function yuanMwhToKwh(v) { return Number(v || 0) / 1000; }
  function average(arr) { return arr.reduce((a, b) => a + Number(b || 0), 0) / (arr.length || 1); }
  function todayISO() { return new Date().toISOString().slice(0, 10); }
  function selectedMonths() {
    const months = state.data.longTermWeightedClearing?.months || [];
    const chosen = state.ui.longMonths || [];
    return months.filter((m) => chosen.includes(m.month));
  }
  function selectedAgentMonths() {
    const months = state.data.agentPurchase?.months || [];
    const chosen = (state.ui.agentMonths || []).filter(Boolean);
    return months.filter((m) => chosen.includes(m.month));
  }
  function latestLongMonth() {
    const months = state.data.longTermWeightedClearing?.months || [];
    return months[months.length - 1]?.month || '';
  }
  function latestAgentMonth() {
    const months = state.data.agentPurchase?.months || [];
    return months[months.length - 1]?.month || state.data.agentPurchase?.month || latestLongMonth();
  }
  function nextMonthValue(month) {
    const m = String(month || latestAgentMonth() || todayISO().slice(0, 7));
    const mm = m.match(/^(\d{4})-(\d{2})$/);
    if (!mm) return latestAgentMonth() || todayISO().slice(0, 7);
    const y = Number(mm[1]);
    const mon = Number(mm[2]);
    const next = new Date(Date.UTC(y, mon, 1));
    return `${next.getUTCFullYear()}-${String(next.getUTCMonth() + 1).padStart(2, '0')}`;
  }
  function renderMultiPicker({ id, open, selected, items, labelFn, placeholder }) {
    const itemLabel = (x) => {
      if (typeof x === 'string') return labelFn(items.find((it) => it.value === x) || { value: x });
      return labelFn(x);
    };
    const label = selected.length ? '' : `<span class="month-placeholder">${esc(placeholder)}</span>`;
    return `
      <div class="month-picker picker-${id} ${open ? 'open' : ''}">
        <button type="button" id="${id}Picker" class="month-picker-control" aria-expanded="${open ? 'true' : 'false'}">
          <span class="month-picker-values">
            ${label}
            ${selected.map((v) => `<span class="month-chip">${esc(itemLabel(v))}</span>`).join('')}
          </span>
          <span class="month-picker-arrow">⌄</span>
        </button>
        ${open ? `
          <div class="month-menu" role="listbox" aria-multiselectable="true">
            ${items.map((item) => {
              const checked = selected.includes(item.value);
              return `<button type="button" class="month-option ${checked ? 'selected' : ''}" data-picker-id="${id}" data-picker-value="${esc(item.value)}" aria-selected="${checked ? 'true' : 'false'}"><span>${esc(itemLabel(item))}</span><span class="month-check">${checked ? '✓' : ''}</span></button>`;
            }).join('')}
          </div>
        ` : ''}
      </div>
    `;
  }
  function renderLongMonthPicker(months) {
    return renderMultiPicker({
      id: 'longMonth',
      open: state.ui.longMonthOpen,
      selected: state.ui.longMonths || [],
      items: months.map((m) => ({ value: m.month, month: m.month })),
      labelFn: (m) => monthLabel(m.month),
      placeholder: '请选择月份'
    });
  }
  function renderAgentMonthPicker(months) {
    return renderMultiPicker({
      id: 'agentMonth',
      open: state.ui.agentMonthOpen,
      selected: state.ui.agentMonths || [],
      items: months.map((m) => ({ value: m.month, month: m.month })),
      labelFn: (m) => monthLabel(m.month),
      placeholder: '请选择月份'
    });
  }
  function renderVoltagePicker(options = voltageLevelOptions(), value = selectedVoltageId(), id = 'voltageLevel') {
    return `
      <select id="${id}">
        ${options.map((v) => `<option value="${esc(v.id)}" ${v.id === value ? 'selected' : ''}>${esc(v.label || v.id)}</option>`).join('')}
      </select>
    `;
  }
  function renderUserPicker(users) {
    return renderMultiPicker({
      id: 'loadUser',
      open: state.ui.loadUserOpen,
      selected: (state.ui.loadUsers || []).slice(0, 1),
      items: users.map((u) => ({ value: u, name: u })),
      labelFn: (u) => u.name || u.value,
      placeholder: '请选择用户'
    });
  }
  function renderSingleMonthPicker(months, current, id) {
    const selected = months.find((m) => m.month === current) || months[months.length - 1];
    return `
      <select id="${id}">
        ${months.map((m) => `<option value="${m.month}" ${m.month === (selected?.month || current) ? 'selected' : ''}>${monthLabel(m.month)}</option>`).join('')}
      </select>
    `;
  }
  function getAgentMonth(month) {
    const list = state.data.agentPurchase?.months || [];
    return list.find((m) => m.month === month) || list[list.length - 1] || { month: month || latestAgentMonth(), flatPriceKwh: 0, averagePurchaseKwh: 0, historyDeviationKwh: 0 };
  }
  function getAgentMonthEntry(month) {
    const agent = state.data.agentPurchase;
    const list = agent.months || (agent.months = []);
    let item = list.find((m) => m.month === month);
    if (!item) {
      item = { month, flatPriceKwh: agent.flatPriceKwh || 0, averagePurchaseKwh: agent.averagePurchaseKwh || 0, historyDeviationKwh: agent.historyDeviationKwh || 0 };
      list.push(item);
    }
    return item;
  }
  function splitLevelForMonth(levels, voltageId = selectedVoltageId()) {
    const list = Array.isArray(levels) ? levels : [];
    const ref = findVoltageOptionStrict(splitVoltageLevelOptions(), voltageId) || voltageId;
    const selected = findVoltageOptionStrict(list, ref);
    return selected || null;
  }
  function splitVoltageLevelOptions() {
    const splitLevels = (state.data.marketSplit?.months || []).flatMap((m) => m.voltageLevels || []);
    const discovered = normalizeSplitVoltageLevels(splitLevels.length ? splitLevels : (state.data.marketSplit?.voltageLevels || state.data.agentPurchase?.voltageLevels || []));
    const merged = new Map();
    for (const option of canonicalVoltageOptions()) merged.set(voltageFullKey(option), option);
    for (const option of discovered) {
      const key = voltageFullKey(option);
      if (!key) continue;
      const existing = merged.get(key);
      merged.set(key, existing ? { ...option, id: existing.id, label: existing.label, tier: existing.tier || option.tier } : option);
    }
    return normalizeSplitVoltageLevels([...merged.values()]);
  }
  function selectedSplitVoltageId() {
    const options = splitVoltageLevelOptions();
    if (!options.length) return '';
    const selected = findVoltageOptionStrict(options, state.ui.splitVoltageLevelId) || findVoltageOptionStrict(options, state.ui.voltageLevelId);
    if (selected) {
      state.ui.splitVoltageLevelId = selected.id;
      return state.ui.splitVoltageLevelId;
    }
    state.ui.splitVoltageLevelId = options[0].id;
    return state.ui.splitVoltageLevelId;
  }
  function getSplitMonthEntry(month) {
    const split = state.data.marketSplit;
    const list = split.months || (split.months = []);
    let item = list.find((m) => m.month === month);
    if (!item) {
      item = { month, voltageLevels: clone(split.voltageLevels || []) };
      list.push(item);
    }
    return item;
  }
  function sortSplitMonths() {
    const list = state.data.marketSplit?.months || [];
    list.sort((a, b) => String(a.month || '').localeCompare(String(b.month || '')));
    invalidateDataCache();
  }
  function getSplitMonthRows(voltageId = selectedVoltageId()) {
    const list = state.data.marketSplit?.months || [];
    const cacheKey = `${list === cache.splitMonthsRef ? 'same' : 'diff'}__${voltageId || ''}`;
    if (cache.splitCacheKey === cacheKey) return cache.splitRows;
    cache.splitMonthsRef = list;
    cache.splitCacheKey = cacheKey;
    cache.splitRows = list.map((m) => {
      const base = splitLevelForMonth(m.voltageLevels, voltageId);
      if (!base) return { month: m.month, lineLoss: null, transmission: null, fund: null, system: null, total: null };
      const lineLoss = yuanKwhToMwh(Number(base.lineLossKwh || 0));
      const transmission = yuanKwhToMwh(Number(base.transmissionKwh || 0));
      const fund = yuanKwhToMwh(Number(base.fundKwh || 0));
      const system = yuanKwhToMwh(Number(base.systemKwh || 0));
      return { month: m.month, lineLoss, transmission, fund, system, total: lineLoss + transmission + fund + system };
    });
    return cache.splitRows;
  }
  function splitSummaryByMonth(month, voltageId = selectedVoltageId()) {
    return getSplitMonthRows(voltageId).find((r) => r.month === month) || { lineLoss: 0, transmission: 0, fund: 0, system: 0, total: 0 };
  }
  function formatSplitCell(v) { return v === null || v === undefined || Number.isNaN(v) ? '/' : formatNumber(v); }
  function isFiniteDataValue(v) { return v !== null && v !== undefined && v !== '' && Number.isFinite(Number(v)); }
  function splitTotalForPrice(month, voltageId, includeSplit) {
    if (!includeSplit) return 0;
    const summary = splitSummaryByMonth(month, voltageId);
    return isFiniteDataValue(summary?.total) ? Number(summary.total) : null;
  }
  function splitSummary(month, voltageId = selectedVoltageId()) {
    return splitSummaryByMonth(month || state.ui.splitMonth || latestLongMonth(), voltageId);
  }
  function seasonKind(month) {
    const m = monthIndex(month);
    if (m === 0 || m === 1 || m === 11) return 'winter';
    if (m >= 2 && m <= 4) return 'spring';
    if (m >= 5 && m <= 7) return 'summer';
    return 'autumn';
  }
  function hasSuperPeak(month) {
    return seasonKind(month) === 'winter' || seasonKind(month) === 'summer';
  }
  function hourType(month, hour) {
    const map = {
      winter: { peak: [6, 7, 16, 21, 22], superPeak: [17, 18, 19, 20], valley: [2, 3, 4, 10, 11, 12, 13, 14], flat: [0, 1, 5, 8, 9, 15, 23] },
      spring: { peak: [6, 7, 17, 18, 19, 20, 21, 22, 23], superPeak: [], valley: [2, 3, 4, 9, 10, 11, 12, 13, 14], flat: [0, 1, 5, 8, 15, 16] },
      summer: { peak: [6, 7, 22, 23], superPeak: [18, 19, 20, 21], valley: [2, 3, 4, 10, 11, 12, 13, 14], flat: [0, 1, 5, 8, 9, 15, 16, 17] },
      autumn: { peak: [6, 7, 17, 18, 19, 20, 21, 22, 23], superPeak: [], valley: [2, 3, 4, 11, 12, 13, 14], flat: [0, 1, 5, 8, 9, 10, 15, 16] }
    }[seasonKind(month)];
    if (map.superPeak.includes(hour)) return 'superPeak';
    if (map.peak.includes(hour)) return 'peak';
    if (map.valley.includes(hour)) return 'valley';
    return 'flat';
  }
  function touMultiplier(type) {
    if (type === 'superPeak') return 1.92;
    if (type === 'peak') return 1.6;
    if (type === 'valley') return 0.45;
    return 1;
  }
  function agentHourly(month, averageKwh, deviationKwh) {
    return Array.from({ length: 24 }, (_, h) => {
      const mult = touMultiplier(hourType(month, h));
      return yuanKwhToMwh(Number(averageKwh || 0) * mult + Number(deviationKwh || 0));
    });
  }
  function normalizeAgentMonthData(monthObj) {
    if (!monthObj || typeof monthObj !== 'object') return monthObj;
    const flat = Number(monthObj.flatPriceKwh);
    const average = Number.isFinite(Number(monthObj.averagePurchaseKwh))
      ? Number(monthObj.averagePurchaseKwh)
      : (Number.isFinite(flat) ? flat : 0);
    const deviation = Number.isFinite(Number(monthObj.historyDeviationKwh))
      ? Number(monthObj.historyDeviationKwh)
      : (Number.isFinite(flat) ? flat - average : 0);
    monthObj.averagePurchaseKwh = round(average, 6);
    monthObj.historyDeviationKwh = round(deviation, 6);
    monthObj.flatPriceKwh = round(average + deviation, 6);
    if (monthObj.touPricesManual !== true) {
      monthObj.touPricesKwh = {
        superPeak: hasSuperPeak(monthObj.month) ? round(average * touMultiplier('superPeak') + deviation, 6) : null,
        peak: round(average * touMultiplier('peak') + deviation, 6),
        flat: monthObj.flatPriceKwh,
        valley: round(average * touMultiplier('valley') + deviation, 6)
      };
      monthObj.touPricesManual = false;
    }
    return monthObj;
  }
  function agentTouPricesKwh(monthObj) {
    normalizeAgentMonthData(monthObj);
    if (monthObj?.touPricesManual === true && monthObj?.touPricesKwh && typeof monthObj.touPricesKwh === 'object') {
      return {
        superPeak: hasSuperPeak(monthObj.month) ? Number(monthObj.touPricesKwh.superPeak ?? null) : null,
        peak: Number(monthObj.touPricesKwh.peak ?? 0),
        flat: Number(monthObj.flatPriceKwh ?? 0),
        valley: Number(monthObj.touPricesKwh.valley ?? 0)
      };
    }
    const averageKwh = Number(monthObj?.averagePurchaseKwh || 0);
    const deviationKwh = Number(monthObj?.historyDeviationKwh || 0);
    return {
      superPeak: hasSuperPeak(monthObj?.month) ? averageKwh * touMultiplier('superPeak') + deviationKwh : null,
      peak: averageKwh * touMultiplier('peak') + deviationKwh,
      flat: Number(monthObj?.flatPriceKwh || 0),
      valley: averageKwh * touMultiplier('valley') + deviationKwh
    };
  }
  function renderTouValue(v) {
    return v === null || v === undefined || v === '' ? '' : formatNumber(yuanKwhToMwh(v));
  }
  function longModeLabel(includeSplit, k) {
    const prefix = '旬及以上中长期分时段加权出清价';
    return includeSplit ? `${prefix}+市场分摊+K=${k}` : `${prefix}+K=${k}`;
  }
  function baseChartUnits() {
    return Array.from({ length: 24 }, (_, i) => String(i));
  }
  function longSeriesForMonth(monthObj, voltageId = selectedVoltageId(), k = state.ui.longK, includeSplit = state.ui.longSplit) {
    const split = splitTotalForPrice(monthObj?.month, voltageId, includeSplit);
    return (monthObj?.values || []).map((v) => (split === null || !isFiniteDataValue(v) ? null : round(Number(v) + split + Number(k || 0), 3)));
  }
  function agentSeries(monthObj = null, voltageId = selectedVoltageId(), includeSplit = state.ui.agentSplit) {
    const month = monthObj || getAgentMonth(state.ui.agentMonth || latestAgentMonth());
    const split = splitTotalForPrice(month.month, voltageId, includeSplit);
    const tou = agentTouPricesKwh(month);
    return Array.from({ length: 24 }, (_, h) => {
      const type = hourType(month.month || latestAgentMonth(), h);
      const price = tou[type] ?? tou.flat ?? 0;
      return split === null || !isFiniteDataValue(price) ? null : round(yuanKwhToMwh(price) + split, 3);
    });
  }
  function monthSeries(monthObj) {
    return monthObj.values.map((v) => round(Number(v), 3));
  }
  function maxIndex(values) {
    let idx = 0;
    for (let i = 1; i < values.length; i++) if (values[i] > values[idx]) idx = i;
    return idx;
  }
  function minIndex(values) {
    let idx = 0;
    for (let i = 1; i < values.length; i++) if (values[i] < values[idx]) idx = i;
    return idx;
  }
  function formatNumber(v) { return fmt(v, 3); }
  function voltageLevelOptions() {
    const splitOptions = splitVoltageLevelOptions();
    const list = splitOptions.length ? splitOptions : (state.data.agentPurchase?.voltageLevels || []);
    const seen = new Set();
    return list.filter((item) => {
      const key = voltageFullKey(item) || String(item.id || item.label || '');
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }
  function selectedVoltageId() {
    const options = voltageLevelOptions();
    if (!options.length) return '';
    const selected = findVoltageOptionStrict(options, state.ui.voltageLevelId);
    if (selected) {
      state.ui.voltageLevelId = selected.id;
      return state.ui.voltageLevelId;
    }
    state.ui.voltageLevelId = options[0].id;
    return state.ui.voltageLevelId;
  }
  function currentVoltageLevel(levels = null) {
    const options = levels || voltageLevelOptions();
    const id = selectedVoltageId();
    const match = findVoltageOptionStrict(options, id);
    return match || options[0] || { id: '', label: '' };
  }
  function selectedVoltageLabel() {
    return currentVoltageLevel()?.label || '';
  }
  function voltageLevelMatch(levels, voltageId = selectedVoltageId()) {
    const list = levels || [];
    const current = currentVoltageLevel();
    return findVoltageOptionStrict(list, voltageId) || findVoltageOptionStrict(list, current) || list[0] || null;
  }
  function parseMonthParts(month) {
    const m = String(month || '').match(/^(\d{4})-(\d{2})$/);
    return m ? { year: Number(m[1]), month: Number(m[2]) } : null;
  }
  function pickPriceMonthByDate(targetMonth, months, fallbackMonth = '') {
    const list = Array.isArray(months) ? months : [];
    if (!list.length) return null;
    const exact = list.find((m) => m.month === targetMonth);
    if (exact) return exact;
    const target = parseMonthParts(targetMonth);
    if (!target) return list.find((m) => m.month === fallbackMonth) || list[0] || null;
    const sameMonth = list.filter((m) => parseMonthParts(m.month)?.month === target.month);
    if (sameMonth.length) {
      return sameMonth.sort((a, b) => {
        const ay = Math.abs((parseMonthParts(a.month)?.year || 0) - target.year);
        const by = Math.abs((parseMonthParts(b.month)?.year || 0) - target.year);
        if (ay !== by) return ay - by;
        return String(a.month).localeCompare(String(b.month));
      })[0];
    }
    return list.find((m) => m.month === fallbackMonth) || list[0] || null;
  }

  function render() {
    const app = byId('app');
    const tabs = [
      ['longTerm', '旬及以上中长期分时段加权出清价'],
      ['agent', '国网代购价'],
      ['split', '市场分摊'],
      ['load', '用户负荷数据'],
      ['compare', '价格对比']
    ];
    app.innerHTML = `
      <div class="shell">
        <aside class="sidebar">
          <div class="brand">
            <h1>山西销售服务系统</h1>
          </div>
          <nav class="nav">
            ${tabs.map(([id, label]) => `<button data-tab="${id}" class="${state.tab === id ? 'active' : ''}"><span class="mark"></span><span>${label}</span></button>`).join('')}
          </nav>
          <div class="foot"></div>
        </aside>
        <main class="main">
          <div class="topbar">
            <div>
              <h2>${pageTitle()}</h2>
            </div>
            <div class="actions">
              <span class="pill">${state.tab === 'load' ? '电量单位：MWH' : '价格单位：元/MWh'}</span>
              <button class="icon-btn" id="settingsBtn" title="设置">⚙</button>
            </div>
          </div>
          <div class="content" id="pageHost"></div>
        </main>
      </div>
      ${state.settingsOpen ? settingsModal() : ''}
    `;
    app.querySelectorAll('[data-tab]').forEach((btn) => btn.addEventListener('click', () => { state.tab = btn.dataset.tab; state.ui.activeTab = state.tab; persist(); scheduleRender(); }));
    byId('settingsBtn').addEventListener('click', () => { state.settingsOpen = true; scheduleRender(); });
    byId('pageHost').innerHTML = pageHtml();
    bindPage();
    app.querySelectorAll('[data-legend-group]').forEach((btn) => btn.addEventListener('click', () => {
      toggleLegend(btn.dataset.legendGroup, btn.dataset.legendValue);
    }));
    if (state.settingsOpen) bindSettings();
  }

  function pageTitle() {
    return {
      longTerm: '旬及以上中长期分时段加权出清价',
      agent: '国网代购价',
      split: '市场分摊',
      load: '用户负荷数据',
      compare: '价格对比'
    }[state.tab];
  }

  function pageHtml() {
    if (state.tab === 'longTerm') return renderLongTerm();
    if (state.tab === 'agent') return renderAgent();
    if (state.tab === 'split') return renderSplit();
    if (state.tab === 'compare') return renderCompare();
    return renderLoad();
  }

  function renderLongTerm() {
    const months = state.data.longTermWeightedClearing?.months || [];
    const voltageId = selectedVoltageId();
    const chosen = selectedMonths();
    const lines = chosen.map((m) => ({
      label: m.month,
      values: longSeriesForMonth(m, voltageId),
      color: seriesColor(m.month),
      showExtrema: false
    }));
    const hidden = new Set((state.ui.longLegendHidden || []).filter((d) => chosen.some((m) => m.month === d)));
    const legendButtons = lines.map((line) => {
      const on = !hidden.has(line.label);
      return `<button type="button" class="chart-legend-item ${on ? 'active' : 'inactive'}" data-legend-group="long" data-legend-value="${esc(line.label)}" aria-pressed="${on ? 'true' : 'false'}"><span class="legend-swatch" style="background:${line.color}"></span><span class="legend-label">${esc(line.label)}</span></button>`;
    }).join('');
    return `
      <div class="section controls-section">
        <div class="toolbar">
          <div class="group"><label>月份</label>${renderLongMonthPicker(months)}</div>
          <div class="group"><label>电压等级</label>${renderVoltagePicker()}</div>
          <div class="group"><label>市场分摊</label><div class="toggle"><button data-long-split="off" class="${state.ui.longSplit ? '' : 'active'}">不含市场分摊</button><button data-long-split="on" class="${state.ui.longSplit ? 'active' : ''}">含市场分摊</button></div></div>
          <div class="group"><label>K</label><select id="longK">${Array.from({ length: 15 }, (_, i) => `<option value="${i}" ${Number(state.ui.longK) === i ? 'selected' : ''}>${i}</option>`).join('')}</select></div>
        </div>
      </div>
      <div class="section">
        <div class="chart-wrap">
          <div class="chart-head"><div><b>${longModeLabel(state.ui.longSplit, state.ui.longK)}</b></div></div>
          <canvas id="longChart"></canvas>
          <div class="chart-legend" aria-label="旬及以上中长期分时段加权出清价图例">${legendButtons}</div>
        </div>
      </div>
      <div class="section">
        <div class="table-actions"><button class="ghost-btn" data-export="long">导出表格</button></div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>日期</th>
                ${Array.from({ length: 24 }, (_, i) => `<th>${i}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${chosen.map((m) => {
                const vals = longSeriesForMonth(m, voltageId);
                return `<tr><td>${m.month}</td>${vals.map((v) => `<td>${formatNumber(v)}</td>`).join('')}</tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  function renderAgent() {
    const months = state.data.agentPurchase?.months || [];
    const voltageId = selectedVoltageId();
    const chosen = selectedAgentMonths();
    const lines = chosen.map((m) => ({
      label: m.month,
      values: agentSeries(m, voltageId),
      color: seriesColor(m.month),
      showExtrema: false
    }));
    const hidden = new Set((state.ui.agentLegendHidden || []).filter((d) => chosen.some((m) => m.month === d)));
    const legendButtons = lines.map((line) => {
      const on = !hidden.has(line.label);
      return `<button type="button" class="chart-legend-item ${on ? 'active' : 'inactive'}" data-legend-group="agent" data-legend-value="${esc(line.label)}" aria-pressed="${on ? 'true' : 'false'}"><span class="legend-swatch" style="background:${line.color}"></span><span class="legend-label">${esc(line.label)}</span></button>`;
    }).join('');
    return `
      <div class="section controls-section">
        <div class="toolbar">
          <div class="group"><label>月份</label>${renderAgentMonthPicker(months)}</div>
          <div class="group"><label>电压等级</label>${renderVoltagePicker()}</div>
          <div class="group"><label>市场分摊</label><div class="toggle"><button data-agent-split="off" class="${state.ui.agentSplit ? '' : 'active'}">不含市场分摊</button><button data-agent-split="on" class="${state.ui.agentSplit ? 'active' : ''}">含市场分摊</button></div></div>
        </div>
      </div>
      <div class="section">
        <div class="chart-wrap">
          <div class="chart-head"><div><b>${state.ui.agentSplit ? '国网代购价+市场分摊' : '国网代购价'}</b></div></div>
          <canvas id="agentChart"></canvas>
          <div class="chart-legend" aria-label="国网代购价图例">${legendButtons}</div>
        </div>
      </div>
      <div class="section">
        <div class="table-actions"><button class="ghost-btn" data-export="agent">导出表格</button></div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>日期</th>
                ${Array.from({ length: 24 }, (_, i) => `<th>${i}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${chosen.map((m) => `<tr><td>${m.month}</td>${agentSeries(m, voltageId).map((v) => `<td>${formatNumber(v)}</td>`).join('')}</tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  function renderSplit() {
    const voltageId = selectedSplitVoltageId();
    const rows = getSplitMonthRows(voltageId);
    return `
      <div class="section controls-section">
        <div class="toolbar">
          <div class="group"><label>电压等级</label>${renderVoltagePicker(splitVoltageLevelOptions(), voltageId, 'splitVoltageLevel')}</div>
        </div>
      </div>
      <div class="section">
        <div class="table-actions"><button class="ghost-btn" data-export="split">导出表格</button></div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>月份</th>
                <th>上网环节线损折价(元/MWh)</th>
                <th>电量输配电价(元/MWh)</th>
                <th>政府性基金及附加(元/MWh)</th>
                <th>系统运行费折价(元/MWh)</th>
                <th>合计(元/MWh)</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map((r) => `<tr><td>${esc(r.month || '')}</td><td>${formatSplitCell(r.lineLoss)}</td><td>${formatSplitCell(r.transmission)}</td><td>${formatSplitCell(r.fund)}</td><td>${formatSplitCell(r.system)}</td><td>${formatSplitCell(r.total)}</td></tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  function renderLoad() {
    const agg = aggregateLoad(state.data.userLoad?.records || []);
    const users = [...new Set(agg.map((r) => r.userName))].sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'));
    const filtered = agg.filter((r) => {
      const userHit = state.ui.loadUsers.length ? state.ui.loadUsers.includes(r.userName) : false;
      const d = r.date;
      const after = !state.ui.loadDateFrom || d >= state.ui.loadDateFrom;
      const before = !state.ui.loadDateTo || d <= state.ui.loadDateTo;
      return userHit && after && before;
    });
    const selected = filtered;
    const tableRows = selected.slice(0, TABLE_PREVIEW_LIMIT);
    const chartRows = selected.slice(0, CHART_SERIES_LIMIT);
    const previewNote = selected.length > TABLE_PREVIEW_LIMIT
      ? `<span class="preview-note">当前预览前 ${TABLE_PREVIEW_LIMIT} 行，导出包含全部 ${selected.length} 行</span>`
      : '';
    const chartNote = selected.length > CHART_SERIES_LIMIT
      ? `<span class="preview-note">曲线预览前 ${CHART_SERIES_LIMIT} 条</span>`
      : '';
    const lines = chartRows.map((r, idx) => {
      const vals = state.ui.loadMode === '24' ? to24(r.points) : r.points;
      return { label: r.date, values: vals, color: seriesColor(r.date) };
    });
    const hidden = new Set((state.ui.loadLegendHidden || []).filter((d) => chartRows.some((r) => r.date === d)));
    const visibleLines = lines.filter((line) => !hidden.has(line.label));
    const legendButtons = lines.map((line) => {
      const on = !hidden.has(line.label);
      return `<button type="button" class="chart-legend-item ${on ? 'active' : 'inactive'}" data-legend-group="load" data-legend-value="${esc(line.label)}" aria-pressed="${on ? 'true' : 'false'}"><span class="legend-swatch" style="background:${line.color}"></span><span class="legend-label">${esc(line.label)}</span></button>`;
    }).join('');
    return `
      <div class="section controls-section">
        <div class="toolbar">
          <div class="group"><label>用户</label>${renderUserPicker(users)}</div>
          <div class="group"><label>起始</label><input id="loadDateFrom" type="date" value="${state.ui.loadDateFrom || ''}" /></div>
          <div class="group"><label>截止</label><input id="loadDateTo" type="date" value="${state.ui.loadDateTo || ''}" /></div>
          <div class="group"><label>视图</label><div class="toggle"><button data-load-mode="96" class="${state.ui.loadMode === '96' ? 'active' : ''}">96点</button><button data-load-mode="24" class="${state.ui.loadMode === '24' ? 'active' : ''}">24点</button></div></div>
        </div>
      </div>
      <div class="section">
        <div class="chart-wrap">
          <div class="chart-head"><div><b>用户负荷</b></div><div>${chartNote}</div></div>
          <canvas id="loadChart"></canvas>
          <div class="chart-legend" aria-label="用户负荷图例">${legendButtons}</div>
        </div>
      </div>
      <div class="section">
        <div class="table-actions">${previewNote}<button class="ghost-btn" data-export="load">导出表格</button></div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>用户</th>
                <th>日期</th>
                <th>日总量</th>
                ${(state.ui.loadMode === '24' ? Array.from({ length: 24 }, (_, i) => String(i)) : loadLabels()).map((x) => `<th>${x}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${tableRows.map((r) => {
                const vals = state.ui.loadMode === '24' ? to24(r.points) : r.points;
                return `<tr><td>${esc(r.userName)}</td><td>${r.date}</td><td>${fmt(r.dailyTotal)}</td>${vals.map((v) => `<td>${formatNumber(v)}</td>`).join('')}</tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  function renderCompare() {
    const months = state.data.longTermWeightedClearing?.months || [];
    const voltageId = selectedVoltageId();
    const month = state.ui.compareMonth || latestLongMonth();
    const longMonth = pickPriceMonthByDate(month, months, latestLongMonth()) || months[months.length - 1];
    const agentMonth = pickPriceMonthByDate(month, state.data.agentPurchase?.months || [], latestAgentMonth()) || getAgentMonth(month);
    const longVals = longMonth ? longSeriesForMonth(longMonth, voltageId, state.ui.compareK, state.ui.compareSplit) : [];
    const agentVals = agentSeries(agentMonth, voltageId, state.ui.compareSplit);
    const rows = (state.ui.compareRows || []).length ? state.ui.compareRows : [];
    const loadRows = compareRowsForMonthlyLoad(rows, voltageId);
    const loadTableRows = loadRows.slice(0, TABLE_PREVIEW_LIMIT);
    const loadPreviewNote = loadRows.length > TABLE_PREVIEW_LIMIT
      ? `<span class="preview-note">当前预览前 ${TABLE_PREVIEW_LIMIT} 行，导出包含全部 ${loadRows.length} 行</span>`
      : '';
    return `
      <div class="section">
        <div class="toolbar">
          <div class="group"><label>月份</label>${renderSingleMonthPicker(months, month, 'compareMonth')}</div>
          <div class="group"><label>电压等级</label>${renderVoltagePicker()}</div>
          <div class="group"><label>市场分摊</label><div class="toggle"><button data-compare-split="off" class="${state.ui.compareSplit ? '' : 'active'}">不含市场分摊</button><button data-compare-split="on" class="${state.ui.compareSplit ? 'active' : ''}">含市场分摊</button></div></div>
          <div class="group"><label>K</label><select id="compareK">${Array.from({ length: 15 }, (_, i) => `<option value="${i}" ${Number(state.ui.compareK) === i ? 'selected' : ''}>${i}</option>`).join('')}</select></div>
        </div>
      </div>
      <div class="section">
        <div class="table-actions"><button class="ghost-btn" data-export="comparePrice">导出价格对比表</button></div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>时点</th>
                <th>中长期</th>
                <th>国网代购价</th>
                <th>差额</th>
              </tr>
            </thead>
            <tbody>
              ${longVals.map((v, i) => {
                const a = agentVals[i] ?? null;
                const diff = isFiniteDataValue(v) && isFiniteDataValue(a) ? Number(v) - Number(a) : null;
                return `<tr><td>${i}</td><td>${formatNumber(v)}</td><td>${formatNumber(a)}</td><td>${formatNumber(diff)}</td></tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
      <div class="section">
        <div class="table-actions">
          <label class="primary-btn import-btn">导入Excel数据<input id="compareLoadFile" type="file" accept=".xlsx,.xls" multiple class="hidden" /></label>
          ${loadPreviewNote}
          <button class="ghost-btn" data-clear-compare-load>清除测算表</button>
          <button class="ghost-btn" data-export="compareLoad">导出用户测算表</button>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>用户</th>
                <th>月份</th>
                ${Array.from({ length: 24 }, (_, i) => `<th>${i}</th>`).join('')}
                <th>售电公司加权均价</th>
                <th>国网代购加权均价</th>
              </tr>
            </thead>
            <tbody>
              ${loadTableRows.map((r) => `<tr><td>${esc(r.userName)}</td><td>${r.month}</td>${r.values.map((v) => `<td>${formatNumber(v)}</td>`).join('')}<td>${formatNumber(r.longAvg)}</td><td>${formatNumber(r.agentAvg)}</td></tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  function compareRowsForMonthlyLoad(rows, voltageId = selectedVoltageId()) {
    const grouped = new Map();
    for (const r of rows || []) {
      const month = normalizeMonthValue(String(r.date || '').slice(0, 7));
      if (!month) continue;
      const key = `${r.userName || ''}__${month}`;
      if (!grouped.has(key)) grouped.set(key, { userName: r.userName || '', month, values: Array(24).fill(0) });
      const item = grouped.get(key);
      const values = resampleSeriesExact(r.points || [], 24);
      for (let i = 0; i < 24; i++) item.values[i] = Number(item.values[i] || 0) + Number(values[i] || 0);
    }
    return [...grouped.values()].map((row) => {
      const longMonth = pickPriceMonthByDate(row.month, state.data.longTermWeightedClearing?.months || [], latestLongMonth()) || {};
      const agentMonth = pickPriceMonthByDate(row.month, state.data.agentPurchase?.months || [], latestAgentMonth()) || getAgentMonth(row.month || state.ui.compareMonth || latestAgentMonth());
      const longVals = longSeriesForMonth(longMonth, voltageId, state.ui.compareK, state.ui.compareSplit);
      const agentVals = agentSeries(agentMonth, voltageId, state.ui.compareSplit);
      const total = row.values.reduce((a, b) => a + Number(b || 0), 0) || 0;
      const longReady = longVals.length >= 24 && longVals.every(isFiniteDataValue);
      const agentReady = agentVals.length >= 24 && agentVals.every(isFiniteDataValue);
      const longTotal = longReady ? row.values.reduce((a, b, i) => a + Number(b || 0) * Number(longVals[i]), 0) : null;
      const agentTotal = agentReady ? row.values.reduce((a, b, i) => a + Number(b || 0) * Number(agentVals[i]), 0) : null;
      return {
        userName: row.userName,
        month: row.month,
        values: row.values.map((v) => round(v, 6)),
        longAvg: total && longReady ? round(longTotal / total, 3) : null,
        agentAvg: total && agentReady ? round(agentTotal / total, 3) : null
      };
    }).sort((a, b) => {
      const userCompare = normalizeCellText(a.userName || '').localeCompare(normalizeCellText(b.userName || ''), 'zh-Hans-CN');
      if (userCompare) return userCompare;
      return String(a.month || '').localeCompare(String(b.month || ''));
    });
  }

  function to24(points) {
    const out = [];
    for (let i = 0; i < 24; i++) out.push(round((points || []).slice(i * 4, i * 4 + 4).reduce((a, b) => a + Number(b || 0), 0), 3));
    return out;
  }

  function resampleSeriesExact(points, targetCount) {
    const src = (Array.isArray(points) ? points : []).map((v) => Number(v || 0));
    if (!targetCount) return src;
    if (!src.length) return Array(targetCount).fill(0);
    if (src.length === targetCount) return src;
    if (src.length < targetCount && targetCount % src.length === 0) {
      const factor = targetCount / src.length;
      const out = [];
      for (const v of src) {
        const each = Number(v || 0) / factor;
        for (let i = 0; i < factor; i++) out.push(each);
      }
      return out.slice(0, targetCount);
    }
    if (src.length > targetCount && src.length % targetCount === 0) {
      const factor = src.length / targetCount;
      const out = [];
      for (let i = 0; i < targetCount; i++) {
        const slice = src.slice(i * factor, i * factor + factor);
        out.push(slice.reduce((a, b) => a + Number(b || 0), 0));
      }
      return out;
    }
    const ratio = src.length / targetCount;
    const out = [];
    for (let i = 0; i < targetCount; i++) {
      const start = Math.floor(i * ratio);
      const end = Math.max(start + 1, Math.floor((i + 1) * ratio));
      const slice = src.slice(start, end);
      out.push(slice.reduce((a, b) => a + Number(b || 0), 0));
    }
    return out;
  }

  function to96From24(values24) {
    const out = [];
    for (const v of values24 || []) {
      const each = Number(v || 0) / 4;
      for (let i = 0; i < 4; i++) out.push(each);
    }
    while (out.length < 96) out.push(0);
    return out.slice(0, 96);
  }

  function combineLoadRecordsByUserDate(records) {
    const grouped = new Map();
    for (const r of records || []) {
      const userName = normalizeCellText(r.userName || '');
      const date = parseDateMaybe(r.date || '');
      if (!userName || !/^\d{4}-\d{2}-\d{2}$/.test(date)) continue;
      const key = `${userName}__${date}`;
      if (!grouped.has(key)) {
        grouped.set(key, {
          userName,
          accountNo: r.accountNo || '',
          contractNo: r.contractNo || '',
          date,
          signed: r.signed || '',
          meterType: r.meterType || '',
          dailyTotal: 0,
          points: Array(96).fill(0),
          _dailyTotalSum: 0,
          _contractNos: new Set(),
          _meterTypes: new Set()
        });
      }
      const g = grouped.get(key);
      if (!g.accountNo && r.accountNo) g.accountNo = r.accountNo;
      if (r.contractNo) g._contractNos.add(String(r.contractNo));
      if (r.meterType) g._meterTypes.add(String(r.meterType));
      if (!g.signed && r.signed) g.signed = r.signed;
      g._dailyTotalSum += Number(r.dailyTotal || 0);
      const pts = Array.isArray(r.points) ? (r.points.length === 96 ? r.points : resampleSeriesExact(r.points, 96)) : [];
      for (let i = 0; i < 96; i++) g.points[i] = round(Number(g.points[i] || 0) + Number(pts[i] || 0), 6);
    }
    return [...grouped.values()].map((g) => {
      const pointTotal = round(g.points.reduce((a, b) => a + Number(b || 0), 0), 3);
      return {
        userName: g.userName,
        accountNo: g.accountNo || '',
        contractNo: [...g._contractNos].join('、') || g.contractNo || '',
        date: g.date,
        signed: g.signed || '',
        meterType: [...g._meterTypes].join('、') || g.meterType || '',
        dailyTotal: chooseLoadDailyTotal(g._dailyTotalSum, pointTotal),
        points: g.points.map((v) => round(v, 6))
      };
    });
  }

  function aggregateLoad(records) {
    if (cache.loadRecordsRef === records && cache.loadRecordsLength === records.length) return cache.loadAggregated;
    cache.loadRecordsRef = records;
    cache.loadRecordsLength = records.length;
    cache.loadAggregated = sortLoadRecordsByUserDate(combineLoadRecordsByUserDate(records));
    return cache.loadAggregated;
  }

  function sortLoadRecordsByUserDate(records) {
    return [...(records || [])].sort((a, b) => {
      const userCompare = normalizeCellText(a.userName || '').localeCompare(normalizeCellText(b.userName || ''), 'zh-Hans-CN');
      if (userCompare) return userCompare;
      return String(a.date || '').localeCompare(String(b.date || ''));
    });
  }

  function settingsModal() {
    return `
      <div class="modal-backdrop" id="settingsBackdrop">
        <div class="modal">
          <h4>设置</h4>
          ${state.settingsUnlocked ? `
            <div class="cloud-sync-panel">
              <div class="cloud-sync-status">${esc(remoteStatus)}</div>
              <div class="row cloud-sync-controls">
                <input id="remoteEmail" type="email" placeholder="Supabase登录邮箱" autocomplete="username" />
                <input id="remotePassword" type="password" placeholder="Supabase登录密码" autocomplete="current-password" />
                <button type="button" class="primary-btn" id="remoteLoginBtn">${remoteAccessToken ? '重新登录云端' : '登录云端'}</button>
                <button type="button" class="ghost-btn" id="remoteSyncBtn">立即同步</button>
                ${remoteAccessToken ? '<button type="button" class="ghost-btn" id="remoteLogoutBtn">退出云端</button>' : ''}
              </div>
              <div class="note">导入或修改数据后先保存到本机；登录云端后点击“立即同步”，其他访问者即可读取最新数据。</div>
            </div>
            <div class="tabs" style="margin-bottom:12px;">
              <button type="button" data-settings-tab="load" class="${state.settingsTab === 'load' ? 'active' : ''}">用户数据导入</button>
              <button type="button" data-settings-tab="long" class="${state.settingsTab === 'long' ? 'active' : ''}">中长期数据</button>
              <button type="button" data-settings-tab="agent" class="${state.settingsTab === 'agent' ? 'active' : ''}">国网代购价数据</button>
              <button type="button" data-settings-tab="split" class="${state.settingsTab === 'split' ? 'active' : ''}">市场分摊</button>
            </div>
            ${settingsBody()}
          ` : `
            <div class="note">输入密码后可管理数据</div>
            <div class="row">
              <input id="settingsPwd" type="password" placeholder="密码" />
              <button class="primary-btn" id="unlockSettingsBtn">进入</button>
              <button class="ghost-btn" id="closeSettingsBtn">关闭</button>
            </div>
          `}
        </div>
      </div>
    `;
  }

  function settingsBody() {
    if (state.settingsTab === 'load') {
      const allRows = aggregateLoad(state.data.userLoad?.records || []);
      const rows = managedLoadRows();
      const users = [...new Set(allRows.map((r) => r.userName).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'));
      const currentUser = state.ui.loadManageUser && users.includes(state.ui.loadManageUser) ? state.ui.loadManageUser : '';
      const mode = state.ui.loadManageMode === '24' ? '24' : '96';
      const labels = mode === '24' ? Array.from({ length: 24 }, (_, i) => String(i)) : loadPointLabels96();
      return `
        <div class="row load-manage-controls" style="margin-bottom:12px;">
          <label>用户</label>
          <select id="loadManageUser">
            <option value="" ${currentUser ? '' : 'selected'}>全部用户</option>
            ${users.map((u) => `<option value="${esc(u)}" ${u === currentUser ? 'selected' : ''}>${esc(u)}</option>`).join('')}
          </select>
          <label>起始</label>
          <input id="loadManageDateFrom" type="date" value="${state.ui.loadManageDateFrom || ''}" />
          <label>截止</label>
          <input id="loadManageDateTo" type="date" value="${state.ui.loadManageDateTo || ''}" />
          <label>视图</label>
          <div class="toggle">
            <button type="button" data-load-manage-mode="96" class="${mode === '96' ? 'active' : ''}">96点</button>
            <button type="button" data-load-manage-mode="24" class="${mode === '24' ? 'active' : ''}">24点</button>
          </div>
        </div>
        <div class="row">
          <label class="primary-btn" style="display:inline-flex;align-items:center;gap:8px;cursor:pointer;">导入用户Excel<input id="importLoadFile" type="file" accept=".xlsx,.xls" multiple class="file-input-hidden" /></label>
          <button type="button" class="primary-btn" id="saveLoadSettingsBtn">保存</button>
          <button class="ghost-btn" id="exportDataBtn">导出Excel</button>
          <button class="danger-btn" id="resetDataBtn">恢复默认</button>
        </div>
        <div class="table-wrap" style="margin-top:12px; max-height:48vh;">
          <table>
            <thead><tr><th>用户</th><th>日期</th><th>日总量</th>${labels.map((x) => `<th>${esc(x)}</th>`).join('')}<th>操作</th></tr></thead>
            <tbody>
              ${rows.slice(0, 200).map((r) => {
                const values = mode === '24' ? to24(r.points) : (r.points || []).slice(0, 96);
                return `
                  <tr>
                    <td>${esc(r.userName)}</td>
                    <td>${r.date}</td>
                    <td><input data-load-total-edit="${esc(`${r.userName}__${r.date}`)}" type="number" step="0.001" value="${fmt(r.dailyTotal)}" /></td>
                    ${values.map((v, idx) => `<td><input class="load-point-input" data-load-point-edit="${esc(`${r.userName}__${r.date}`)}" data-point-index="${idx}" type="number" step="0.001" value="${formatNumber(v)}" /></td>`).join('')}
                    <td><button type="button" class="ghost-btn" data-load-save="${esc(`${r.userName}__${r.date}`)}">保存</button><button type="button" class="danger-btn" data-load-del="${esc(`${r.userName}__${r.date}`)}">删除</button></td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      `;
    }
    if (state.settingsTab === 'long') {
      const months = state.data.longTermWeightedClearing?.months || [];
      return `
        <div class="note" style="margin-bottom:12px;">可新增、删除或修改月份的 24 时段价格。</div>
        <div class="row" style="margin-bottom:12px;">
          <button type="button" class="primary-btn import-btn" data-open-file="importLongFile">导入中长期Excel</button>
          <input id="importLongFile" type="file" accept=".xlsx,.xls" class="file-input-hidden" />
        </div>
        <div class="table-wrap long-term-table-wrap" style="max-height:60vh;">
          <table>
            <thead>
              <tr>
                <th>月份</th>
                ${Array.from({ length: 24 }, (_, i) => `<th>${i}</th>`).join('')}
                <th>操作</th>
              </tr>
            </thead>
            <tbody>
              ${months.map((m, idx) => `
                <tr>
                  <td><input data-long-month-input="${idx}" value="${m.month}" /></td>
                  ${Array.from({ length: 24 }, (_, i) => {
                    const v = (m.values || [])[i];
                    return `<td><input data-long-price-input="${idx}" data-long-price-index="${i}" type="number" step="0.001" value="${formatNumber(v)}" /></td>`;
                  }).join('')}
                  <td><button class="danger-btn" data-long-del="${idx}">删除</button></td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
        <div class="row"><button class="primary-btn" id="addLongBtn">新增月份</button><button class="primary-btn" id="saveLongBtn">保存</button><button class="ghost-btn" id="closeSettingsBtn">关闭</button></div>
      `;
    }
    if (state.settingsTab === 'agent') {
      const months = state.data.agentPurchase?.months || [];
      const current = getAgentMonthEntry(state.ui.agentMonth || latestAgentMonth());
      normalizeAgentMonthData(current);
      const rows = months.filter((m) => /^\d{4}-\d{2}$/.test(String(m.month || '')));
      return `
        <div class="row" style="margin-bottom:12px;">
          <button type="button" class="primary-btn import-btn" data-open-file="importAgentImageFile">导入图片数据</button>
          <button type="button" class="primary-btn import-btn" data-open-file="importAgentExcelFile">导入Excel数据</button>
          <input id="importAgentImageFile" type="file" accept="image/*" class="file-input-hidden" />
          <input id="importAgentExcelFile" type="file" accept=".xlsx,.xls" class="file-input-hidden" />
        </div>
        <div class="table-wrap" style="max-height:60vh;">
          <table>
            <thead><tr><th>字段</th><th>值</th></tr></thead>
            <tbody>
              <tr><td>月份</td><td><input id="agentEditMonth" type="month" value="${current.month || ''}" /></td></tr>
              <tr><td>平段价(元/MWh)</td><td><input id="agentEditFlat" type="number" step="0.001" value="${formatNumber(yuanKwhToMwh(current.flatPriceKwh || 0))}" readonly /></td></tr>
              <tr><td>当月平均购电价格(元/MWh)</td><td><input id="agentEditAvg" type="number" step="0.001" value="${formatNumber(yuanKwhToMwh(current.averagePurchaseKwh || 0))}" /></td></tr>
              <tr><td>历史偏差电费折价(元/MWh)</td><td><input id="agentEditDev" type="number" step="0.001" value="${formatNumber(yuanKwhToMwh(current.historyDeviationKwh || 0))}" /></td></tr>
            </tbody>
          </table>
        </div>
        <div class="table-actions" style="margin-top:12px;">
          <button type="button" class="primary-btn" id="addAgentMonthBtn">新增月份</button>
        </div>
        <div class="table-wrap" style="max-height:60vh;">
          <table>
            <thead><tr><th>月份</th><th>尖峰(元/MWh)</th><th>峰段(元/MWh)</th><th>平段(元/MWh)</th><th>谷段(元/MWh)</th><th>操作</th></tr></thead>
          <tbody>
              ${rows.map((m, idx) => {
                const tou = agentTouPricesKwh(m);
                return `<tr data-agent-month-row="${esc(m.month || '')}">
                  <td><input data-agent-month-field="month" value="${esc(m.month || '')}" /></td>
                  <td><input data-agent-month-field="superPeak" type="number" step="0.001" value="${renderTouValue(tou.superPeak)}" ${hasSuperPeak(m.month) ? '' : 'placeholder=""'} /></td>
                  <td><input data-agent-month-field="peak" type="number" step="0.001" value="${renderTouValue(tou.peak)}" /></td>
                  <td><input data-agent-month-field="flatTou" type="number" step="0.001" value="${renderTouValue(tou.flat)}" /></td>
                  <td><input data-agent-month-field="valley" type="number" step="0.001" value="${renderTouValue(tou.valley)}" /></td>
                  <td>
                    <button type="button" class="icon-btn" data-agent-month-move="up" data-agent-month-index="${idx}" title="上移">↑</button>
                    <button type="button" class="icon-btn" data-agent-month-move="down" data-agent-month-index="${idx}" title="下移">↓</button>
                    <button type="button" class="danger-btn" data-agent-month-del="${esc(m.month || '')}">删除</button>
                  </td>
                </tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
        <div class="row"><button class="primary-btn" id="saveAgentBtn">保存</button><button class="ghost-btn" id="closeSettingsBtn">关闭</button></div>
      `;
    }
    const months = state.data.marketSplit?.months || [];
    const current = getSplitMonthEntry(state.ui.splitMonth || months[months.length - 1]?.month || latestLongMonth());
    return `
      <div class="note" style="margin-bottom:12px;">市场分摊按月份管理，下面是当前月份的数据。</div>
      <div class="row" style="margin-bottom:12px;">
        <label>月份</label>
        <select id="splitManageMonth">${months.map((m) => `<option value="${m.month}" ${m.month === current.month ? 'selected' : ''}>${m.month}</option>`).join('')}</select>
        <button type="button" class="primary-btn import-btn" data-open-file="importSplitImageFile">导入图片数据</button>
        <button type="button" class="primary-btn import-btn" data-open-file="importSplitExcelFile">导入Excel数据</button>
        <input id="importSplitImageFile" type="file" accept="image/*" class="file-input-hidden" />
        <input id="importSplitExcelFile" type="file" accept=".xlsx,.xls" class="file-input-hidden" />
      </div>
      <div class="table-wrap" style="max-height:28vh; margin-bottom:12px;">
        <table>
          <thead><tr><th>月份</th><th>操作</th></tr></thead>
          <tbody>${months.map((m, idx) => `<tr draggable="true" data-split-month-row="${idx}"><td>${esc(m.month || '')}</td><td><button type="button" class="icon-btn" data-split-month-move="up" data-split-month-index="${idx}" title="上移">↑</button><button type="button" class="icon-btn" data-split-month-move="down" data-split-month-index="${idx}" title="下移">↓</button></td></tr>`).join('')}</tbody>
        </table>
      </div>
      <div class="table-wrap" style="max-height:60vh;">
        <table>
          <thead><tr><th>电压等级</th><th>线损</th><th>输配</th><th>基金</th><th>系统</th></tr></thead>
          <tbody>${(current.voltageLevels || []).map((v, idx) => `<tr><td><input data-split-v="${idx}" data-key="label" value="${v.label}" /></td><td><input data-split-v="${idx}" data-key="lineLossKwh" type="number" step="0.000001" value="${v.lineLossKwh}" /></td><td><input data-split-v="${idx}" data-key="transmissionKwh" type="number" step="0.000001" value="${v.transmissionKwh}" /></td><td><input data-split-v="${idx}" data-key="fundKwh" type="number" step="0.000001" value="${v.fundKwh}" /></td><td><input data-split-v="${idx}" data-key="systemKwh" type="number" step="0.000001" value="${v.systemKwh}" /></td></tr>`).join('')}</tbody>
        </table>
      </div>
      <div class="row"><button class="primary-btn" id="saveSplitBtn">保存</button><button class="ghost-btn" id="closeSettingsBtn">关闭</button></div>
    `;
  }

  function bindPage() {
    if (state.tab === 'longTerm') {
      byId('longMonthPicker').addEventListener('click', (e) => {
        e.stopPropagation();
        state.ui.longMonthOpen = !state.ui.longMonthOpen;
        scheduleRender();
      });
      document.querySelectorAll('[data-picker-id="longMonth"]').forEach((btn) => btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const month = btn.dataset.pickerValue;
        const selected = state.ui.longMonths.includes(month);
        state.ui.longMonths = selected
          ? state.ui.longMonths.filter((m) => m !== month)
          : [...state.ui.longMonths, month];
        state.ui.longMonthOpen = true;
        persist(); scheduleRender();
      }));
      document.addEventListener('click', closeLongMonthPicker, { once: true });
      byId('longK').addEventListener('change', (e) => { state.ui.longK = Number(e.target.value || 0); persist(); scheduleRender(); });
      document.querySelectorAll('[data-long-split]').forEach((btn) => btn.addEventListener('click', () => { state.ui.longSplit = btn.dataset.longSplit === 'on'; persist(); scheduleRender(); }));
      const voltageId = selectedVoltageId();
      const longSeries = selectedMonths().map((m) => ({
        label: m.month,
        values: longSeriesForMonth(m, voltageId, state.ui.longK),
        color: seriesColor(m.month),
        showExtrema: false
      }));
      const allLong = (state.data.longTermWeightedClearing?.months || []).flatMap((m) => longSeriesForMonth(m, voltageId, state.ui.longK));
      longSeries.fixedMax = allLong.length ? Math.ceil(Math.max(...allLong) / 10) * 10 : 10;
      const hidden = new Set((state.ui.longLegendHidden || []).filter((d) => selectedMonths().some((m) => m.month === d)));
      drawLineChart('longChart', longSeries.filter((line) => !hidden.has(line.label)), '元/MWh', null, { legend: false });
    }
    if (state.tab === 'agent') {
      const voltageId = selectedVoltageId();
      byId('agentMonthPicker').addEventListener('click', (e) => {
        e.stopPropagation();
        state.ui.agentMonthOpen = !state.ui.agentMonthOpen;
        scheduleRender();
      });
      document.querySelectorAll('[data-picker-id="agentMonth"]').forEach((btn) => btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const month = btn.dataset.pickerValue;
        const selected = state.ui.agentMonths || [];
        state.ui.agentMonths = selected.includes(month)
          ? selected.filter((m) => m !== month)
          : [...selected, month];
        state.ui.agentMonthOpen = true;
        state.ui.agentMonth = month;
        persist(); scheduleRender();
      }));
      document.addEventListener('click', closeAgentMonthPicker, { once: true });
      document.querySelectorAll('[data-agent-split]').forEach((btn) => btn.addEventListener('click', () => { state.ui.agentSplit = btn.dataset.agentSplit === 'on'; persist(); scheduleRender(); }));
      const chosen = selectedAgentMonths();
      const series = chosen.map((month) => ({
        label: month.month,
        values: agentSeries(month, voltageId, state.ui.agentSplit),
        color: seriesColor(month.month),
        showExtrema: false
      }));
      const hidden = new Set((state.ui.agentLegendHidden || []).filter((d) => chosen.some((m) => m.month === d)));
      drawLineChart('agentChart', series.filter((line) => !hidden.has(line.label)), '元/MWh', null, { legend: false });
    }
    if (state.tab === 'split') {
      const splitVoltageLevel = byId('splitVoltageLevel');
      if (splitVoltageLevel) splitVoltageLevel.addEventListener('change', (e) => { state.ui.splitVoltageLevelId = e.target.value; persist(); scheduleRender(); });
    }
    if (state.tab === 'load') {
      byId('loadUserPicker').addEventListener('click', (e) => {
        e.stopPropagation();
        state.ui.loadUserOpen = !state.ui.loadUserOpen;
        scheduleRender();
      });
      document.querySelectorAll('[data-picker-id="loadUser"]').forEach((btn) => btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const user = btn.dataset.pickerValue;
        const selected = state.ui.loadUsers.includes(user);
        state.ui.loadUsers = selected ? [] : [user];
        state.ui.loadUserOpen = true;
        persist(); scheduleRender();
      }));
      byId('loadDateFrom').addEventListener('change', (e) => { state.ui.loadDateFrom = e.target.value; persist(); scheduleRender(); });
      byId('loadDateTo').addEventListener('change', (e) => { state.ui.loadDateTo = e.target.value; persist(); scheduleRender(); });
      document.querySelectorAll('[data-load-mode]').forEach((btn) => btn.addEventListener('click', () => { state.ui.loadMode = btn.dataset.loadMode; persist(); scheduleRender(); }));
      document.addEventListener('click', closeLoadUserPicker, { once: true });
      const agg = aggregateLoad(state.data.userLoad?.records || []);
      const currentUser = (state.ui.loadUsers || [])[0] || '';
      const filtered = agg.filter((r) => {
        const userHit = currentUser ? r.userName === currentUser : false;
        const after = !state.ui.loadDateFrom || r.date >= state.ui.loadDateFrom;
        const before = !state.ui.loadDateTo || r.date <= state.ui.loadDateTo;
        return userHit && after && before;
      });
      const loadSeries = filtered.slice(0, CHART_SERIES_LIMIT).map((r, idx) => ({
        label: r.date,
        values: state.ui.loadMode === '24' ? to24(r.points) : r.points,
        color: seriesColor(r.date),
        showExtrema: false
      }));
      loadSeries.autoYMin = true;
      const hidden = new Set((state.ui.loadLegendHidden || []).filter((d) => filtered.slice(0, CHART_SERIES_LIMIT).some((r) => r.date === d)));
      const visibleLines = loadSeries.filter((line) => !hidden.has(line.label));
      drawLineChart('loadChart', visibleLines, '', state.ui.loadMode === '24' ? Array.from({ length: 24 }, (_, i) => String(i)) : loadLabels(), { legend: false });
    }
    if (state.tab === 'compare') {
      byId('compareMonth').addEventListener('change', (e) => { state.ui.compareMonth = e.target.value; persist(); scheduleRender(); });
      byId('compareK').addEventListener('change', (e) => { state.ui.compareK = Number(e.target.value || 0); persist(); scheduleRender(); });
      document.querySelectorAll('[data-compare-split]').forEach((btn) => btn.addEventListener('click', () => { state.ui.compareSplit = btn.dataset.compareSplit === 'on'; persist(); scheduleRender(); }));
      byId('compareLoadFile').addEventListener('change', handleCompareImport);
      const clearCompareLoad = document.querySelector('[data-clear-compare-load]');
      if (clearCompareLoad) clearCompareLoad.addEventListener('click', () => {
        state.ui.compareRows = [];
        persist();
        scheduleRender();
      });
    }
    const voltageSelectGlobal = byId('voltageLevel');
    if (voltageSelectGlobal) voltageSelectGlobal.addEventListener('change', (e) => { state.ui.voltageLevelId = e.target.value; persist(); scheduleRender(); });
  }

  function closeLongMonthPicker(e) {
    if (state.tab !== 'longTerm' || !state.ui.longMonthOpen) return;
    if (e.target.closest?.('.month-picker')) return;
    state.ui.longMonthOpen = false;
    scheduleRender();
  }

  function closeAgentMonthPicker(e) {
    if (state.tab !== 'agent' || !state.ui.agentMonthOpen) return;
    if (e.target.closest?.('.month-picker')) return;
    state.ui.agentMonthOpen = false;
    scheduleRender();
  }

  function closeLoadUserPicker(e) {
    if (state.tab !== 'load' || !state.ui.loadUserOpen) return;
    if (e.target.closest?.('.month-picker')) return;
    state.ui.loadUserOpen = false;
    scheduleRender();
  }

  function bindSettings() {
    if (!state.settingsUnlocked) {
      const unlockBtn = byId('unlockSettingsBtn');
      const pwd = byId('settingsPwd');
      if (unlockBtn) unlockBtn.addEventListener('click', unlockSettings);
      if (pwd) pwd.addEventListener('keydown', (e) => { if (e.key === 'Enter') unlockSettings(); });
      return;
    }
    const remoteLoginBtn = byId('remoteLoginBtn');
    const remoteSyncBtn = byId('remoteSyncBtn');
    const remoteLogoutBtn = byId('remoteLogoutBtn');
    if (remoteLoginBtn) remoteLoginBtn.addEventListener('click', async () => {
      remoteLoginBtn.disabled = true;
      try {
        await loginRemote();
        await syncRemoteData();
        alert('云端登录成功');
      } catch (err) {
        alert(err?.message || err);
      } finally {
        remoteLoginBtn.disabled = false;
      }
    });
    if (remoteSyncBtn) remoteSyncBtn.addEventListener('click', async () => {
      remoteSyncBtn.disabled = true;
      try {
        await saveRemoteData();
        alert(remoteAccessToken ? '数据已同步到云端' : '请先登录云端');
      } catch (err) {
        alert(err?.message || err);
      } finally {
        remoteSyncBtn.disabled = false;
      }
    });
    if (remoteLogoutBtn) remoteLogoutBtn.addEventListener('click', logoutRemote);
    document.querySelectorAll('[data-settings-tab]').forEach((btn) => {
      btn.type = 'button';
      btn.addEventListener('click', () => { state.settingsTab = btn.dataset.settingsTab; scheduleRender(); });
    });
    if (state.settingsTab === 'load') {
      byId('importLoadFile').addEventListener('change', handleImport);
      const saveLoadBtn = byId('saveLoadSettingsBtn');
      if (saveLoadBtn) saveLoadBtn.addEventListener('click', saveAllManagedLoadRows);
      byId('resetDataBtn').addEventListener('click', () => { if (confirm('恢复默认数据？')) { clearTimeout(persistTimer); localStorage.removeItem(STORAGE_KEY); location.reload(); } });
      const manageUser = byId('loadManageUser');
      const manageFrom = byId('loadManageDateFrom');
      const manageTo = byId('loadManageDateTo');
      if (manageUser) manageUser.addEventListener('change', (e) => { state.ui.loadManageUser = e.target.value; persist(); scheduleRender(); });
      if (manageFrom) manageFrom.addEventListener('change', (e) => { state.ui.loadManageDateFrom = e.target.value; persist(); scheduleRender(); });
      if (manageTo) manageTo.addEventListener('change', (e) => { state.ui.loadManageDateTo = e.target.value; persist(); scheduleRender(); });
      document.querySelectorAll('[data-load-save]').forEach((btn) => btn.addEventListener('click', () => saveManagedLoadRow(btn.dataset.loadSave)));
      document.querySelectorAll('[data-load-del]').forEach((btn) => btn.addEventListener('click', () => {
        if (confirm('删除该用户该日期的数据？')) deleteManagedLoadRow(btn.dataset.loadDel);
      }));
    }
    if (state.settingsTab === 'long') {
      const importLong = byId('importLongFile');
      if (importLong) importLong.addEventListener('change', handleImportLong);
      document.querySelectorAll('[data-long-month-input]').forEach((el) => el.addEventListener('change', (e) => {
        state.data.longTermWeightedClearing.months[Number(el.dataset.longMonthInput)].month = e.target.value;
        invalidateDataCache();
        persistData();
      }));
      document.querySelectorAll('[data-long-price-input]').forEach((el) => el.addEventListener('change', (e) => {
        const idx = Number(el.dataset.longPriceInput);
        const pointIdx = Number(el.dataset.longPriceIndex);
        const item = state.data.longTermWeightedClearing.months[idx];
        if (!item) return;
        item.values = Array.isArray(item.values) ? item.values : Array(24).fill(0);
        item.values[pointIdx] = Number(e.target.value || 0);
        invalidateDataCache();
        persistData();
      }));
      document.querySelectorAll('[data-long-del]').forEach((btn) => btn.addEventListener('click', () => {
        state.data.longTermWeightedClearing.months.splice(Number(btn.dataset.longDel), 1);
        invalidateDataCache(); persistData(); scheduleRender();
      }));
      const addBtn = byId('addLongBtn');
      if (addBtn) addBtn.addEventListener('click', () => {
        state.data.longTermWeightedClearing.months.push({ month: '2026-08', values: Array(24).fill(0) });
        invalidateDataCache(); persistData(); scheduleRender();
      });
      const saveLongBtn = byId('saveLongBtn');
      if (saveLongBtn) saveLongBtn.addEventListener('click', saveLongSettings);
    }
    if (state.settingsTab === 'agent') {
      const agentImage = byId('importAgentImageFile');
      const agentExcel = byId('importAgentExcelFile');
      if (agentImage) agentImage.addEventListener('change', handleImportAgentImage);
      if (agentExcel) agentExcel.addEventListener('change', handleImportAgentExcel);
      const agentMonthInput = byId('agentEditMonth');
      if (agentMonthInput) {
        agentMonthInput.addEventListener('change', () => {
          const month = normalizeMonthValue(agentMonthInput.value || '') || latestAgentMonth();
          state.ui.agentMonth = month;
          persist();
          render();
        });
      }
      ['agentEditFlat', 'agentEditAvg', 'agentEditDev'].forEach((id) => {
        const el = byId(id);
        if (el) el.addEventListener('change', () => saveAgentSettings(id));
      });
      const addAgentBtn = byId('addAgentMonthBtn');
      if (addAgentBtn) addAgentBtn.addEventListener('click', addAgentMonthRow);
      document.querySelectorAll('[data-agent-month-field]').forEach((el) => el.addEventListener('change', () => {
        const row = el.closest('[data-agent-month-row]');
        if (!row) return;
        saveAgentMonthRow(row, el.dataset.agentMonthField);
      }));
      document.querySelectorAll('[data-agent-month-move]').forEach((btn) => btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        moveAgentMonthRow(Number(btn.dataset.agentMonthIndex || 0), btn.dataset.agentMonthMove === 'up' ? -1 : 1);
      }));
      document.querySelectorAll('[data-agent-month-del]').forEach((btn) => btn.addEventListener('click', () => deleteAgentMonthRow(btn.dataset.agentMonthDel)));
      byId('saveAgentBtn').addEventListener('click', () => { saveAgentSettings('manual', { silent: true }); saveAgentMonthsTable(); showSaved('国网代购价数据已保存'); });
    }
    if (state.settingsTab === 'split') {
      const splitManageMonth = byId('splitManageMonth');
      if (splitManageMonth) splitManageMonth.addEventListener('change', (e) => { state.ui.splitMonth = e.target.value; scheduleRender(); });
      const splitImage = byId('importSplitImageFile');
      const splitExcel = byId('importSplitExcelFile');
      if (splitImage) splitImage.addEventListener('change', handleImportSplitImage);
      if (splitExcel) splitExcel.addEventListener('change', handleImportSplitExcel);
      document.querySelectorAll('[data-split-month-row]').forEach((row) => {
        row.addEventListener('dragstart', () => {
          splitMonthDragIndex = Number(row.dataset.splitMonthRow || 0);
          row.classList.add('dragging');
        });
        row.addEventListener('dragover', (e) => e.preventDefault());
        row.addEventListener('drop', (e) => {
          e.preventDefault();
          const from = splitMonthDragIndex;
          const to = Number(row.dataset.splitMonthRow || 0);
          splitMonthDragIndex = null;
          if (Number.isFinite(from) && Number.isFinite(to) && from !== to) {
            moveSplitMonthRow(from, to - from);
          }
        });
        row.addEventListener('dragend', () => { splitMonthDragIndex = null; row.classList.remove('dragging'); });
      });
      document.querySelectorAll('[data-split-v]').forEach((el) => el.addEventListener('change', () => {
        const idx = Number(el.dataset.splitV);
        const key = el.dataset.key;
        const val = el.type === 'number' ? Number(el.value || 0) : el.value;
        const item = getSplitMonthEntry(state.ui.splitMonth || latestLongMonth());
        item.voltageLevels = item.voltageLevels || clone(state.data.marketSplit.voltageLevels || []);
        item.voltageLevels[idx][key] = val;
        invalidateDataCache();
        persistData();
        render();
      }));
      document.querySelectorAll('[data-split-month-move]').forEach((btn) => btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        moveSplitMonthRow(Number(btn.dataset.splitMonthIndex || 0), btn.dataset.splitMonthMove === 'up' ? -1 : 1);
      }));
      byId('saveSplitBtn').addEventListener('click', saveSplitSettings);
    }
  }

  function closeSettings() {
    state.settingsOpen = false;
    state.settingsUnlocked = false;
    render();
  }

  function showSaved(message = '保存成功') {
    alert(message);
  }

  function saveAgentSettings(changedId = '', opts = {}) {
    const month = normalizeMonthValue(byId('agentEditMonth').value || state.ui.agentMonth || latestAgentMonth()) || (state.ui.agentMonth || latestAgentMonth());
    const item = getAgentMonthEntry(month);
    item.month = month;
    const avg = yuanMwhToKwh(byId('agentEditAvg').value || 0);
    const dev = yuanMwhToKwh(byId('agentEditDev').value || 0);
    item.averagePurchaseKwh = avg;
    item.historyDeviationKwh = dev;
    item.touPricesManual = false;
    normalizeAgentMonthData(item);
    state.ui.agentMonth = month;
    invalidateDataCache();
    persistData();
    scheduleRender();
  }

  function saveLongSettings() {
    const months = state.data.longTermWeightedClearing?.months || [];
    document.querySelectorAll('[data-long-month-input]').forEach((el) => {
      const item = months[Number(el.dataset.longMonthInput)];
      if (item) item.month = normalizeMonthValue(el.value || item.month) || el.value || item.month;
    });
    document.querySelectorAll('[data-long-price-input]').forEach((el) => {
      const item = months[Number(el.dataset.longPriceInput)];
      const pointIdx = Number(el.dataset.longPriceIndex);
      if (!item || !Number.isFinite(pointIdx)) return;
      item.values = Array.isArray(item.values) ? item.values : Array(24).fill(0);
      item.values[pointIdx] = Number(el.value || 0);
    });
    invalidateDataCache();
    persistData();
    showSaved('中长期数据已保存');
    scheduleRender();
  }

  function agentMonthRowData(rowEl) {
    const month = normalizeMonthValue(rowEl.querySelector('[data-agent-month-field="month"]')?.value || '');
    const toKwh = (value) => {
      const text = String(value ?? '').trim();
      return text === '' ? null : yuanMwhToKwh(text);
    };
    const superPeak = toKwh(rowEl.querySelector('[data-agent-month-field="superPeak"]')?.value);
    const peak = toKwh(rowEl.querySelector('[data-agent-month-field="peak"]')?.value);
    const flatTou = toKwh(rowEl.querySelector('[data-agent-month-field="flatTou"]')?.value);
    const valley = toKwh(rowEl.querySelector('[data-agent-month-field="valley"]')?.value);
    return { month, touPricesKwh: { superPeak, peak, flat: flatTou, valley } };
  }

  function saveAgentMonthRow(rowEl, changedField = '') {
    const oldMonth = normalizeMonthValue(rowEl.dataset.agentMonthRow || '');
    const data = agentMonthRowData(rowEl);
    if (!data.month) return;
    const list = state.data.agentPurchase.months || (state.data.agentPurchase.months = []);
    let item = list.find((m) => m.month === oldMonth) || list.find((m) => m.month === data.month);
    if (!item) {
      item = {};
      list.push(item);
    }
    item.month = data.month;
    item.touPricesKwh = data.touPricesKwh;
    item.touPricesManual = true;
    normalizeAgentMonthData(item);
    if (state.ui.agentMonth === oldMonth) state.ui.agentMonth = data.month;
    invalidateDataCache();
    persistData();
    scheduleRender();
  }

  function addAgentMonthRow() {
    const list = state.data.agentPurchase.months || (state.data.agentPurchase.months = []);
    const baseMonth = state.ui.agentMonth || latestAgentMonth();
    const month = nextMonthValue(baseMonth);
    list.push({
      month,
      flatPriceKwh: 0,
      averagePurchaseKwh: 0,
      historyDeviationKwh: 0,
      touPricesManual: false
    });
    state.ui.agentMonth = month;
    invalidateDataCache();
    persistData();
    scheduleRender();
  }

  function deleteAgentMonthRow(month) {
    const list = state.data.agentPurchase.months || (state.data.agentPurchase.months = []);
    state.data.agentPurchase.months = list.filter((m) => m.month !== month);
    if (state.ui.agentMonth === month) state.ui.agentMonth = latestAgentMonth();
    invalidateDataCache();
    persistData();
    scheduleRender();
  }

  function moveAgentMonthRow(index, delta) {
    const list = state.data.agentPurchase.months || (state.data.agentPurchase.months = []);
    const from = Number(index);
    const to = from + Number(delta || 0);
    if (from < 0 || from >= list.length || to < 0 || to >= list.length) return;
    const [item] = list.splice(from, 1);
    list.splice(to, 0, item);
    invalidateDataCache();
    persistData();
    render();
  }

  function moveSplitMonthRow(index, delta) {
    const list = state.data.marketSplit.months || (state.data.marketSplit.months = []);
    const from = Number(index);
    const to = from + Number(delta || 0);
    if (from < 0 || from >= list.length || to < 0 || to >= list.length) return;
    const [item] = list.splice(from, 1);
    list.splice(to, 0, item);
    invalidateDataCache();
    persistData();
    render();
  }

  function saveAgentMonthsTable() {
    const rows = [...document.querySelectorAll('[data-agent-month-row]')];
    const list = [];
    for (const row of rows) {
      const data = agentMonthRowData(row);
      if (!data.month) continue;
      const oldMonth = normalizeMonthValue(row.dataset.agentMonthRow || '');
      const existing = (state.data.agentPurchase.months || []).find((m) => m.month === oldMonth) || (state.data.agentPurchase.months || []).find((m) => m.month === data.month) || {};
      list.push({
        month: data.month,
        flatPriceKwh: Number(existing.flatPriceKwh || 0),
        averagePurchaseKwh: Number(existing.averagePurchaseKwh || 0),
        historyDeviationKwh: Number(existing.historyDeviationKwh || 0),
        touPricesKwh: data.touPricesKwh,
        touPricesManual: true
      });
    }
    if (list.length) state.data.agentPurchase.months = list;
    invalidateDataCache();
    persistData();
  }

  function saveSplitSettings() {
    const item = getSplitMonthEntry(state.ui.splitMonth || latestLongMonth());
    item.voltageLevels = item.voltageLevels || clone(state.data.marketSplit.voltageLevels || []);
    document.querySelectorAll('[data-split-v]').forEach((el) => {
      const idx = Number(el.dataset.splitV);
      const key = el.dataset.key;
      if (!item.voltageLevels[idx]) item.voltageLevels[idx] = {};
      item.voltageLevels[idx][key] = el.type === 'number' ? Number(el.value || 0) : el.value;
    });
    item.voltageLevels = normalizeSplitVoltageLevels(item.voltageLevels);
    invalidateDataCache();
    persistData();
    showSaved('市场分摊数据已保存');
    scheduleRender();
  }

  function managedLoadRecord(key) {
    const rows = aggregateLoad(state.data.userLoad?.records || []);
    return rows.find((r) => `${r.userName || ''}__${r.date || ''}` === key);
  }

  function managedLoadRows() {
    const rows = aggregateLoad(state.data.userLoad?.records || []);
    const currentUser = state.ui.loadManageUser || '';
    const from = state.ui.loadManageDateFrom || '';
    const to = state.ui.loadManageDateTo || '';
    return rows.filter((r) => {
      const userHit = currentUser ? r.userName === currentUser : true;
      const after = !from || r.date >= from;
      const before = !to || r.date <= to;
      return userHit && after && before;
    }).sort((a, b) => {
      const dateCompare = String(b.date || '').localeCompare(String(a.date || ''));
      if (dateCompare) return dateCompare;
      return String(a.userName || '').localeCompare(String(b.userName || ''), 'zh-Hans-CN');
    });
  }

  function saveManagedLoadRow(key) {
    const old = managedLoadRecord(key);
    if (!old) return;
    const selector = (prefix) => `[${prefix}="${CSS.escape(key)}"]`;
    const total = Number(document.querySelector(selector('data-load-total-edit'))?.value || 0);
    const pointInputs = [...document.querySelectorAll(selector('data-load-point-edit'))].sort((a, b) => Number(a.dataset.pointIndex || 0) - Number(b.dataset.pointIndex || 0));
    const rawValues = pointInputs.map((el) => Number(el.value || 0));
    const points = state.ui.loadManageMode === '24' ? to96From24(rawValues) : rawValues;
    const records = state.data.userLoad?.records || [];
    state.data.userLoad.records = records.filter((r) => `${r.userName || ''}__${r.date || ''}` !== key);
    state.data.userLoad.records.push({
      userName: old.userName,
      accountNo: old.accountNo || '',
      contractNo: old.contractNo || '',
      date: old.date,
      signed: old.signed || '',
      meterType: old.meterType || '',
      dailyTotal: total || round(points.reduce((a, b) => a + Number(b || 0), 0), 3),
      points
    });
    invalidateDataCache();
    persistData();
    scheduleRender();
  }

  function saveAllManagedLoadRows() {
    const keys = [...new Set([...document.querySelectorAll('[data-load-total-edit]')].map((el) => el.dataset.loadTotalEdit).filter(Boolean))];
    keys.forEach((key) => saveManagedLoadRow(key, { silent: true }));
    invalidateDataCache();
    persistData();
    showSaved('用户数据已保存');
    scheduleRender();
  }

  function deleteManagedLoadRow(key) {
    const records = state.data.userLoad?.records || [];
    state.data.userLoad.records = records.filter((r) => `${r.userName || ''}__${r.date || ''}` !== key);
    invalidateDataCache();
    persistData();
    scheduleRender();
  }

  function handleImport(e) {
    const input = e?.target?.files ? e.target : e;
    const files = [...(input?.files || [])];
    if (!files.length) return;
    Promise.all(files.map((file) => file.arrayBuffer().then((buf) => parseLoadWorkbook(buf, file.name)).catch((err) => ({ error: err, file: file.name })))).then((results) => {
      const merged = [];
      const failures = [];
      for (const result of results) {
        if (result?.error) failures.push(`${result.file}：${result.error?.message || result.error}`);
        else merged.push(...(result.records || []));
      }
      if (!merged.length) throw new Error(failures[0] || '没有识别到可用负荷数据');
      const combined = combineLoadRecordsByUserDate(merged);
      state.data.userLoad = {
        headers: results.find((r) => r?.headers)?.headers || defaultLoadHeaders(),
        records: combined
      };
      const importedUsers = [...new Set(combined.map((r) => r.userName).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'));
      state.ui.loadUsers = [];
      state.ui.loadManageUser = '';
      invalidateDataCache();
      persistData();
      if (input) input.value = '';
      if (failures.length) alert(`部分文件导入失败：\n${failures.join('\n')}`);
      scheduleRender();
    }).catch((err) => alert('导入失败：' + err.message));
  }

  function handleImportLong(e) {
    const input = e?.target?.files ? e.target : e;
    const file = input?.files && input.files[0];
    if (!file) return;
    file.arrayBuffer().then((buf) => parseWorkbookRows(buf)).then((rows) => {
      const months = parseLongTermRows(rows);
      if (!months.length) throw new Error('没有识别到中长期月份数据');
      state.data.longTermWeightedClearing = state.data.longTermWeightedClearing || {};
      state.data.longTermWeightedClearing.months = months;
      invalidateDataCache();
      persistData();
      if (input) input.value = '';
      scheduleRender();
    }).catch((err) => alert('导入失败：' + err.message));
  }

  async function handleImportAgentImage(e) {
    const input = e?.target?.files ? e.target : e;
    const file = input?.files && input.files[0];
    if (!file) return;
    const currentMonth = inferMonthFromName(file.name) || state.ui.agentMonth || latestAgentMonth();
    let text = '';
    try {
      setImportFeedback('已选择图片文件', `${file.name}，正在准备识别...`, '');
      scheduleRender();
      text = await withTimeout(recognizeImageText(file, 'agent'), 45000, '代购价图片识别超时');
      const parsed = parseAgentImportText(text, currentMonth);
      state.ui.agentImagePreview = [parsed];
      state.ui.splitImagePreview = [];
      applyAgentImport(parsed);
      setImportFeedback('成功', `已识别并导入 ${parsed.month || currentMonth} 的代购价数据`, buildImportPreview({
        fileName: file.name,
        kind: '代购价图片',
        rawText: text,
        parsed
      }));
    } catch (err) {
      resetOcrWorker().catch(() => {});
      setImportFeedback('失败', `代购价图片识别失败：${err?.message || err}`, buildImportPreview({
        fileName: file.name,
        kind: '代购价图片',
        rawText: text,
        error: err?.message || String(err)
      }));
      alert('图片识别失败：' + (err?.message || err));
    } finally {
      if (input) input.value = '';
      scheduleRender();
    }
  }

  async function handleImportAgentExcel(e) {
    const input = e?.target?.files ? e.target : e;
    const file = input?.files && input.files[0];
    if (!file) return;
    const currentMonth = inferMonthFromName(file.name) || state.ui.agentMonth || latestAgentMonth();
    let text = '';
    try {
      setImportFeedback('已选择Excel文件', `${file.name}，正在准备识别...`, '');
      scheduleRender();
      const rows = await withTimeout(parseWorkbookRows(await file.arrayBuffer()), 15000, '代购价Excel解析超时');
      text = rows.map((row) => Object.keys(row).sort(columnSort).map((k) => row[k]).join(' ')).join('\n');
      const parsed = parseAgentStructuredRows(rows, currentMonth) || parseAgentImportText(text, currentMonth);
      state.ui.agentImagePreview = [parsed];
      state.ui.splitImagePreview = [];
      applyAgentImport(parsed);
      setImportFeedback('成功', `已识别并导入 ${parsed.month || currentMonth} 的代购价Excel数据`, buildImportPreview({
        fileName: file.name,
        kind: '代购价Excel',
        rawText: text,
        parsed
      }));
    } catch (err) {
      setImportFeedback('失败', `代购价Excel识别失败：${err?.message || err}`, buildImportPreview({
        fileName: file.name,
        kind: '代购价Excel',
        rawText: text,
        error: err?.message || String(err)
      }));
      alert('Excel识别失败：' + (err?.message || err));
    } finally {
      if (input) input.value = '';
      scheduleRender();
    }
  }

  async function handleImportSplitImage(e) {
    const input = e?.target?.files ? e.target : e;
    const file = input?.files && input.files[0];
    if (!file) return;
    const currentMonth = inferMonthFromName(file.name) || state.ui.splitMonth || latestLongMonth();
    let text = '';
    try {
      setImportFeedback('已选择图片文件', `${file.name}，正在准备识别...`, '');
      scheduleRender();
      text = await withTimeout(recognizeImageText(file, 'split'), 45000, '市场分摊图片识别超时');
      const parsed = parseSplitImportText(text, currentMonth);
      applySplitImport(parsed);
      sortSplitMonths();
      state.ui.agentImagePreview = [];
      state.ui.splitImagePreview = [parsed];
      setImportFeedback('成功', `已识别并导入 ${parsed.month || currentMonth} 的市场分摊数据`, buildImportPreview({
        fileName: file.name,
        kind: '市场分摊图片',
        rawText: text,
        parsed
      }));
    } catch (err) {
      resetOcrWorker().catch(() => {});
      setImportFeedback('失败', `市场分摊图片识别失败：${err?.message || err}`, buildImportPreview({
        fileName: file.name,
        kind: '市场分摊图片',
        rawText: text,
        error: err?.message || String(err)
      }));
      alert('图片识别失败：' + (err?.message || err));
    } finally {
      if (input) input.value = '';
      scheduleRender();
    }
  }

  async function handleImportSplitExcel(e) {
    const input = e?.target?.files ? e.target : e;
    const file = input?.files && input.files[0];
    if (!file) return;
    const currentMonth = inferMonthFromName(file.name) || state.ui.splitMonth || latestLongMonth();
    let text = '';
    try {
      setImportFeedback('已选择Excel文件', `${file.name}，正在准备识别...`, '');
      scheduleRender();
      const rows = await withTimeout(parseWorkbookRows(await file.arrayBuffer()), 15000, '市场分摊Excel解析超时');
      text = rows.map((row) => Object.keys(row).sort(columnSort).map((k) => row[k]).join(' ')).join('\n');
      const parsed = parseSplitStructuredRows(rows, currentMonth) || parseSplitImportText(text, currentMonth);
      applySplitImport(parsed);
      sortSplitMonths();
      state.ui.agentImagePreview = [];
      state.ui.splitImagePreview = [parsed];
      setImportFeedback('成功', `已识别并导入 ${parsed.month || currentMonth} 的市场分摊Excel数据`, buildImportPreview({
        fileName: file.name,
        kind: '市场分摊Excel',
        rawText: text,
        parsed
      }));
    } catch (err) {
      setImportFeedback('失败', `市场分摊Excel识别失败：${err?.message || err}`, buildImportPreview({
        fileName: file.name,
        kind: '市场分摊Excel',
        rawText: text,
        error: err?.message || String(err)
      }));
      alert('Excel识别失败：' + (err?.message || err));
    } finally {
      if (input) input.value = '';
      scheduleRender();
    }
  }

  function applyAgentImport(parsed) {
    const month = normalizeMonthValue(parsed.month || state.ui.agentMonth || latestAgentMonth()) || (state.ui.agentMonth || latestAgentMonth());
    const item = getAgentMonthEntry(month);
    item.month = month;
    item.flatPriceKwh = Number(parsed.flatPriceKwh || 0);
    item.averagePurchaseKwh = Number(parsed.averagePurchaseKwh || 0);
    item.historyDeviationKwh = Number(parsed.historyDeviationKwh || 0);
    item.touPricesManual = false;
    normalizeAgentMonthData(item);
    item.voltageLevels = clone(parsed.voltageLevels || item.voltageLevels || state.data.agentPurchase?.voltageLevels || []);
    state.ui.agentMonth = month;
    invalidateDataCache();
    persistData();
  }

  function applySplitImport(parsed) {
    const month = normalizeMonthValue(parsed.month || state.ui.splitMonth || latestLongMonth()) || (state.ui.splitMonth || latestLongMonth());
    const item = getSplitMonthEntry(month);
    item.month = month;
    item.voltageLevels = normalizeSplitVoltageLevels(parsed.voltageLevels || item.voltageLevels || state.data.marketSplit?.voltageLevels || []);
    state.ui.splitMonth = month;
    invalidateDataCache();
    persistData();
  }

  function rowsToEntries(parsedRows) {
    return (parsedRows || []).map((row, index) => {
      const orderedKeys = Object.keys(row).sort(columnSort);
      const values = orderedKeys.map((k) => row[k]);
      const text = values.map(normalizeCellText).filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
      return { index, row, values, text };
    }).filter((entry) => entry.text || entry.values.some((v) => v !== '' && v !== null && v !== undefined));
  }

  function inferMonthFromName(name) {
    const m = String(name || '').match(/(\d{4})\D?(\d{1,2})/);
    return m ? `${m[1]}-${String(Number(m[2])).padStart(2, '0')}` : '';
  }

  function normalizeCellText(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).replace(/\u3000/g, ' ').trim();
  }

  function extractMonthFromText(text, fallbackMonth) {
    const raw = String(text || '');
    const execMatch = raw.match(/执行时间[:：]?\s*(\d{4})\s*[./年-]\s*(\d{1,2})/);
    if (execMatch) return `${execMatch[1]}-${String(Number(execMatch[2])).padStart(2, '0')}`;
    const monthMatch = raw.match(/(\d{4})\s*[./年-]\s*(\d{1,2})/);
    return monthMatch ? `${monthMatch[1]}-${String(Number(monthMatch[2])).padStart(2, '0')}` : fallbackMonth;
  }

  function entryNumbers(entry) {
    const out = [];
    for (const value of entry?.values || []) {
      if (typeof value === 'number' && Number.isFinite(value)) out.push(value);
      else if (value instanceof Date) continue;
      else if (value !== null && value !== undefined && value !== '') {
        const matches = String(value).match(/-?\d+(?:\.\d+)?/g) || [];
        for (const m of matches) {
          const n = Number(m);
          if (Number.isFinite(n)) out.push(n);
        }
      }
    }
    return out;
  }

  function entryNumberAt(entry, colIdx) {
    const key = colName(colIdx);
    const value = entry?.row?.[key];
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    const n = Number(String(value ?? '').replace(/[^\d.-]/g, ''));
    return Number.isFinite(n) ? n : null;
  }

  function entryTextAt(entry, colIdx) {
    const key = colName(colIdx);
    return normalizeCellText(entry?.row?.[key] ?? '');
  }

  function findColumnByExactText(entry, target) {
    const t = normalizeCellText(target);
    for (const [key, value] of Object.entries(entry?.row || {})) {
      if (normalizeCellText(value) === t) return colIndex(key);
    }
    return null;
  }

  function entryRightmostNumber(entry) {
    const nums = entryNumbers(entry);
    return nums.length ? nums[nums.length - 1] : null;
  }

  function findMonthInEntries(entries, fallbackMonth) {
    for (const entry of entries || []) {
      const month = extractMonthFromText(entry.text, '');
      if (month) return month;
    }
    return fallbackMonth;
  }

  function parseAgentStructuredRows(parsedRows, fallbackMonth) {
    const entries = rowsToEntries(parsedRows);
    if (!entries.length) return null;
    const month = findMonthInEntries(entries, fallbackMonth);
    const infoStart = entries.findIndex((entry) => /4\s*=\s*5\s*\+\s*6/.test(entry.text));
    if (infoStart < 0) return null;
    const averagePurchaseKwh = entryRightmostNumber(entries[infoStart + 1]);
    const historyDeviationKwh = entryRightmostNumber(entries[infoStart + 2]);
    if (!Number.isFinite(averagePurchaseKwh) || !Number.isFinite(historyDeviationKwh)) return null;
    const headerIndex = entries.slice(0, infoStart).findIndex((entry) => /1\s*=\s*2\+3\+4\+5\+6/.test(entry.text));
    if (headerIndex < 0) return null;
    const parsedBlock = parseAgentPriceBlock(entries, headerIndex, infoStart);
    if (!parsedBlock || !parsedBlock.voltageLevels.length) return null;
    const flatPriceKwh = parsedBlock.baseFlatPriceKwh ?? round(Number(averagePurchaseKwh || 0) + Number(historyDeviationKwh || 0), 6);
    const voltageLevels = parsedBlock.voltageLevels;
    return { month, flatPriceKwh, averagePurchaseKwh, historyDeviationKwh, voltageLevels };
  }

  function parseSplitStructuredRows(parsedRows, fallbackMonth) {
    const agentParsed = parseAgentStructuredRows(parsedRows, fallbackMonth);
    if (!agentParsed) return null;
    return {
      month: agentParsed.month,
      voltageLevels: normalizeSplitVoltageLevels(agentParsed.voltageLevels || [])
    };
  }

  function parseAgentPriceBlock(entries, headerIndex, stopIndex) {
    const header = entries[headerIndex];
    const col1 = findColumnByPattern(header, /^1\s*=/);
    const col2 = findColumnByExactText(header, '2');
    const col3 = findColumnByExactText(header, '3');
    const col4 = findColumnByExactText(header, '4');
    const col5 = findColumnByExactText(header, '5');
    const col6 = findColumnByExactText(header, '6');
    if (![col1, col2, col3, col4, col5, col6].every((v) => Number.isFinite(v))) return null;
    let currentType = '';
    let sharedLineLoss = null;
    let sharedFund = null;
    let sharedSystem = null;
    let sharedAgentPrice = null;
    const voltageLevels = [];
    for (const entry of entries.slice(headerIndex + 1, stopIndex)) {
      if (/1\s*=\s*2\+3\+4\+5\+6/.test(entry.text)) break;
      const voltageText = findRowVoltageText(entry);
      const typeText = findRowPowerType(entry) || cleanPowerType(entryTextAt(entry, 1));
      if (typeText) currentType = typeText;
      const flatPrice = entryNumberAt(entry, col1);
      const agentPrice = entryNumberAt(entry, col2);
      const lineLoss = entryNumberAt(entry, col3);
      const transmission = entryNumberAt(entry, col4);
      const fund = entryNumberAt(entry, col5);
      const system = entryNumberAt(entry, col6);
      const hasPrice = [flatPrice, agentPrice, lineLoss, transmission, fund, system].some(Number.isFinite);
      if (!hasPrice) continue;
      if (Number.isFinite(agentPrice)) sharedAgentPrice = agentPrice;
      if (Number.isFinite(lineLoss)) sharedLineLoss = lineLoss;
      if (Number.isFinite(fund)) sharedFund = fund;
      if (Number.isFinite(system)) sharedSystem = system;
      if (!isVoltageText(voltageText) || !Number.isFinite(transmission)) continue;
      voltageLevels.push({
        id: makeVoltageId(currentType, voltageText, voltageLevels.length),
        label: normalizeVoltageLevelLabel(currentType, voltageText),
        tier: /两部制/.test(currentType) ? 'dual' : 'single',
        flatPriceKwh: Number.isFinite(flatPrice) ? flatPrice : null,
        agentPriceKwh: Number.isFinite(agentPrice) ? agentPrice : sharedAgentPrice,
        lineLossKwh: Number.isFinite(lineLoss) ? lineLoss : (sharedLineLoss ?? 0),
        transmissionKwh: transmission,
        fundKwh: Number.isFinite(fund) ? fund : (sharedFund ?? 0),
        systemKwh: Number.isFinite(system) ? system : (sharedSystem ?? 0)
      });
    }
    return {
      baseFlatPriceKwh: voltageLevels.find((v) => Number.isFinite(v.agentPriceKwh))?.agentPriceKwh ?? null,
      voltageLevels
    };
  }

  function rowTextCandidates(entry) {
    return Object.values(entry?.row || {}).map(normalizeCellText).filter(Boolean);
  }

  function findRowVoltageText(entry) {
    const candidates = rowTextCandidates(entry).filter((text) => isVoltageText(text));
    if (!candidates.length) return '';
    return candidates.sort((a, b) => {
      const score = (s) => {
        const t = normalizeCellText(s);
        let n = 0;
        if (/千伏|kV|KV/i.test(t)) n += 5;
        if (/不满1/.test(t) || /1-10\(20\)/.test(t) || /35/.test(t) || /110/.test(t) || /220/.test(t)) n += 3;
        if (/及以上|以下/.test(t)) n += 2;
        if (t.length < 12) n += 1;
        return n;
      };
      return score(b) - score(a);
    })[0];
  }

  function findRowPowerType(entry) {
    const texts = rowTextCandidates(entry);
    const short = texts.find((text) => {
      const s = normalizeCellText(text).replace(/\s+/g, '');
      return /^(单一制|两部制|single|dual|一|二|单|两)$/i.test(s);
    });
    if (short) {
      const s = normalizeCellText(short).replace(/\s+/g, '').toLowerCase();
      if (/^(两部制|dual|二|两)$/.test(s) || /两部/.test(s)) return '两部制';
      if (/^(单一制|single|一|单)$/.test(s) || /单一/.test(s)) return '单一制';
    }
    return cleanPowerType(texts.find((text) => cleanPowerType(text)) || '');
  }

  function findColumnByPattern(entry, pattern) {
    for (const [key, value] of Object.entries(entry?.row || {})) {
      if (pattern.test(normalizeCellText(value))) return colIndex(key);
    }
    return null;
  }

  function findVoltageColumn(entries, headerIndex, stopIndex) {
    const header = entries[headerIndex];
    for (const [key, value] of Object.entries(header?.row || {})) {
      if (/电压|等级/.test(normalizeCellText(value))) return colIndex(key);
    }
    const scores = new Map();
    for (const entry of entries.slice(headerIndex + 1, stopIndex)) {
      for (const key of Object.keys(entry.row || {})) {
        if (isVoltageText(entryTextAt(entry, colIndex(key)))) scores.set(colIndex(key), (scores.get(colIndex(key)) || 0) + 1);
      }
    }
    let best = { col: 3, score: -1 };
    for (const [col, score] of scores) if (score > best.score) best = { col, score };
    return best.col;
  }

  function isVoltageText(text) {
    const s = normalizeCellText(text);
    if (/容量电价|千伏安|kVA|kva/i.test(s)) return false;
    return /千伏|kV|KV|\d+\s*-\s*\d+|\d+\s*\(/i.test(s) && !/^\/$|^一$/.test(s);
  }

  function cleanPowerType(text) {
    const s = normalizeCellText(text).replace(/[－—–-]+$/g, '').trim();
    if (/两部制|dual|two/i.test(s)) return '两部制';
    if (/单一制|single/i.test(s)) return '单一制';
    return '';
  }

  function inferPowerType(level, rawLabel = '') {
    const direct = cleanPowerType(level?.tier || level?.type || '') || cleanPowerType(rawLabel) || cleanPowerType(level?.id || '');
    if (direct) return direct;
    const id = normalizeVoltageKey(level?.id || '');
    const label = normalizeVoltageKey(rawLabel);
    if (/^(under|under1|1under)|under1|不满1/.test(id) || /不满1/.test(label)) return '单一制';
    return '';
  }

  function normalizeVoltageBandLabel(voltage) {
    const s = normalizeCellText(voltage).replace(/\s+/g, '');
    const key = s.toLowerCase();
    if (!s) return '';
    if (/220/.test(s)) return '220千伏及以上';
    if (/110/.test(s)) return '110千伏';
    if (/35/.test(s)) return '35千伏';
    if (/1[_-]?10|10\(?20\)?/.test(key)) return '1-10(20)千伏';
    if (/under[_-]?1|under1|不满1/.test(key)) return '不满1千伏';
    if (/1\s*[-~—–至到]\s*10/.test(s) || /1-10\(20\)/.test(s) || /1\s*[-~—–至到]\s*10\(20\)/.test(s) || /10\(20\)/.test(s) || /10\/20/.test(s)) return '1-10(20)千伏';
    if (/不满\s*1/.test(s) || /1千伏以下/.test(s) || /1千伏及以下/.test(s) || /^1$/i.test(s) || /<\s*1/.test(s)) return '不满1千伏';
    return s
      .replace(/千伏|kV|kv|KV/g, '千伏')
      .replace(/(1[-~—–至到]10)(千伏)?/g, '1-10(20)千伏')
      .replace(/110(千伏)?(及以上)?/g, '110千伏')
      .replace(/220(千伏)?(及以上)?/g, '220千伏及以上');
  }

  function normalizeVoltageLevelLabel(type, voltage) {
    const v = normalizeVoltageBandLabel(voltage);
    const t = cleanPowerType(type);
    return t ? `${t} ${v}` : v;
  }

  function makeVoltageId(type, voltage, idx) {
    const raw = `${normalizeVoltageLevelLabel(type, voltage)}_${idx}`;
    return raw
      .replace(/不满|以下/g, 'under_')
      .replace(/及以上|以上/g, 'up_')
      .replace(/两部制/g, 'dual')
      .replace(/单一制/g, 'single')
      .replace(/[^\w]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .toLowerCase() || `voltage_${idx + 1}`;
  }

  function canonicalVoltageLabels() {
    return [
      '单一制 不满1千伏',
      '单一制 1-10(20)千伏',
      '单一制 35千伏',
      '单一制 110千伏',
      '单一制 220千伏及以上',
      '两部制 1-10(20)千伏',
      '两部制 35千伏',
      '两部制 110千伏',
      '两部制 220千伏及以上'
    ];
  }

  function canonicalVoltageOptions() {
    return canonicalVoltageLabels().map((label, idx) => {
      const [type, ...bandParts] = label.split(' ');
      const band = bandParts.join(' ');
      return {
        id: makeVoltageId(type, band, idx),
        label,
        tier: type
      };
    });
  }

  function normalizeSplitVoltageLevels(levels) {
    const list = Array.isArray(levels) ? levels : [];
    const canonicalOrder = canonicalVoltageLabels();
    const canonicalByLabel = new Map(canonicalVoltageOptions().map((option) => [option.label, option]));
    const normalized = new Map();
    const extras = [];
    list.forEach((level, idx) => {
      const rawLabel = normalizeCellText(level?.label || level?.name || level?.voltage || level?.id || '');
      if (!rawLabel || /容量电价|千伏安|kVA|kva/i.test(rawLabel)) return;
      const type = inferPowerType(level, rawLabel);
      const band = normalizeVoltageBandLabel(rawLabel.replace(/^(单一制|两部制)\s*/, '')) || normalizeVoltageBandLabel(rawLabel);
      const label = type ? `${type} ${band}`.trim() : band;
      const canonical = canonicalByLabel.get(label);
      const item = {
        ...clone(level),
        id: canonical?.id || level?.id || makeVoltageId(type || level?.tier || '', band || rawLabel, idx),
        label: canonical?.label || label || rawLabel,
        tier: canonical?.tier || type || level?.tier || (String(level?.id || '').includes('dual') ? 'dual' : 'single')
      };
      if (canonicalOrder.includes(item.label)) {
        normalized.set(item.label, item);
      } else {
        extras.push(item);
      }
    });
    return [
      ...canonicalOrder.map((label) => normalized.get(label)).filter(Boolean),
      ...extras
    ];
  }

  async function parseLoadWorkbook(buf, sourceName = '') {
    const tables = await readSpreadsheetTables(buf, sourceName);
    const candidate = selectBestLoadCandidate(tables, sourceName);
    if (!candidate) throw new Error('没有识别到可用负荷数据');
    const workbookUserName = inferUserNameFromWorkbook(tables, sourceName);
    const workbookAccountNo = inferAccountNoFromWorkbook(tables);
    return {
      headers: candidate.headers || defaultLoadHeaders(),
      records: (candidate.records || []).map((r) => ({
        ...r,
        userName: isGenericUserName(r.userName) ? workbookUserName : r.userName,
        accountNo: r.accountNo || workbookAccountNo || ''
      })),
      sourceName: isGenericUserName(candidate.sourceName) ? workbookUserName : (candidate.sourceName || sourceName)
    };
  }

  async function parseWorkbookRows(buf, sourceName = '') {
    const tables = await readSpreadsheetTables(buf, sourceName);
    const best = selectBestStructuredSheet(tables);
    return best?.rows || [];
  }

  async function parseXlsx(buf, sourceName = '') {
    return parseLoadWorkbook(buf, sourceName);
  }

  async function readSpreadsheetTables(buf, sourceName = '') {
    if (isZipBuffer(buf)) return readXlsxTables(buf, sourceName);
    return readXlsTables(buf, sourceName);
  }

  function isZipBuffer(buf) {
    const bytes = new Uint8Array(buf || []);
    return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
  }

  async function readXlsxTables(buf, sourceName = '') {
    const zip = await JSZip.loadAsync(buf);
    const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
    if (!workbookXml) throw new Error('不是有效的Excel文件');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string') || '';
    const appXml = await zip.file('docProps/app.xml')?.async('string') || '';
    const sharedStrings = await readSharedStrings(zip);
    const styles = await readStyles(zip);
    const workbookDoc = new DOMParser().parseFromString(workbookXml, 'application/xml');
    const relsDoc = relsXml ? new DOMParser().parseFromString(relsXml, 'application/xml') : null;
    const relMap = new Map();
    if (relsDoc) {
      [...relsDoc.querySelectorAll('Relationship')].forEach((rel) => {
        relMap.set(rel.getAttribute('Id'), rel.getAttribute('Target'));
      });
    }
    const sheets = [...workbookDoc.querySelectorAll('sheet')].map((sheet) => ({
      name: sheet.getAttribute('name') || 'Sheet',
      target: relMap.get(sheet.getAttribute('r:id')) || ''
    })).filter((sheet) => sheet.target);
    const tables = [];
    for (const sheet of sheets) {
      const path = sheet.target.replace(/^\//, '').replace(/^\.{2}\//, '').replace(/^xl\//, '');
      const xmlPath = path.startsWith('worksheets/') ? `xl/${path}` : `xl/worksheets/${path.replace(/^worksheets\//, '')}`;
      const file = zip.file(xmlPath);
      if (!file) continue;
      const sheetXml = await file.async('string');
      tables.push(parseXlsxSheet(sheetXml, sharedStrings, styles, sheet.name, `${workbookXml}\n${appXml}`));
    }
    return tables.filter((t) => t.rows.length);
  }

  function parseXlsxSheet(sheetXml, sharedStrings, styles, sheetName, workbookText = '') {
    const doc = new DOMParser().parseFromString(sheetXml, 'application/xml');
    const rowsEls = [...doc.querySelectorAll('sheetData > row')];
    const mergeRefs = [...doc.querySelectorAll('mergeCells mergeCell')].map((n) => n.getAttribute('ref')).filter(Boolean);
    const cellMap = new Map();
    let maxRow = 0;
    let maxCol = 0;
    rowsEls.forEach((row) => {
      const rowNum = Number(row.getAttribute('r') || 0);
      maxRow = Math.max(maxRow, rowNum);
      [...row.querySelectorAll('c')].forEach((cell) => {
        const ref = cell.getAttribute('r') || '';
        const col = colIndex(ref.replace(/\d+/g, ''));
        const rowIndex = Number(ref.match(/\d+/)?.[0] || rowNum);
        const value = decodeXlsxCell(cell, sharedStrings, styles);
        if (rowIndex && col) {
          cellMap.set(`${rowIndex},${col}`, value);
          maxCol = Math.max(maxCol, col);
        }
      });
    });
    const matrix = Array.from({ length: maxRow }, () => Array(maxCol).fill(null));
    for (const [key, value] of cellMap.entries()) {
      const [r, c] = key.split(',').map(Number);
      if (r > 0 && c > 0) matrix[r - 1][c - 1] = value;
    }
    for (const ref of mergeRefs) {
      const match = ref?.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
      if (!match) continue;
      const c1 = colIndex(match[1]);
      const r1 = Number(match[2]);
      const c2 = colIndex(match[3]);
      const r2 = Number(match[4]);
      const value = matrix[r1 - 1]?.[c1 - 1];
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          if (!matrix[r - 1]) matrix[r - 1] = Array(maxCol).fill(null);
          if (matrix[r - 1][c - 1] === null || matrix[r - 1][c - 1] === undefined || matrix[r - 1][c - 1] === '') {
            matrix[r - 1][c - 1] = value;
          }
        }
      }
    }
    const rows = matrix.map((row) => matrixRowToObject(row));
    const text = [sheetName, workbookText, ...matrix.flat().map(normalizeCellText)].filter(Boolean).join('\n');
    return { name: sheetName, rows, matrix, text };
  }

  function decodeXlsxCell(cell, sharedStrings, styles) {
    const raw = cell.querySelector('v')?.textContent ?? '';
    const type = cell.getAttribute('t');
    const style = Number(cell.getAttribute('s') || 0);
    if (type === 's') return sharedStrings[Number(raw)] ?? '';
    if (type === 'inlineStr') return cell.querySelector('is t')?.textContent ?? '';
    if (type === 'b') return raw === '1';
    if (type === 'str') return raw;
    if (type === 'd') return raw ? String(raw).slice(0, 10) : '';
    if (raw === '') return '';
    const n = Number(raw);
    if (Number.isFinite(n) && styles.dateStyles?.has(style)) return excelSerialToDate(n);
    if (Number.isFinite(n)) return n;
    return raw;
  }

  function matrixRowToObject(row) {
    const out = {};
    row.forEach((value, idx) => {
      if (value !== null && value !== undefined && value !== '') out[colName(idx + 1)] = value;
    });
    return out;
  }

  function defaultLoadHeaders() {
    const out = ['市场成员名称', '户号', '计量点', '日期', '当前日期是否签约', '所属交易单元名称', '日合计值'];
    for (let i = 0; i < 96; i++) {
      const mins = (i + 1) * 15;
      const hh = String(Math.floor(mins / 60)).padStart(2, '0');
      const mm = String(mins % 60).padStart(2, '0');
      out.push(`${hh}:${mm}`);
    }
    return out;
  }

  function tableToMatrix(table) {
    if (Array.isArray(table?.matrix) && table.matrix.length) return table.matrix;
    const rows = Array.isArray(table?.rows) ? table.rows : [];
    if (!rows.length) return [];
    if (Array.isArray(rows[0])) return rows;
    const maxCol = rows.reduce((max, row) => Math.max(max, ...Object.keys(row || {}).map(colIndex)), 0);
    return rows.map((row) => {
      const arr = Array(maxCol).fill(null);
      Object.entries(row || {}).forEach(([key, value]) => {
        arr[colIndex(key) - 1] = value;
      });
      return arr;
    });
  }

  function cellTextAtRow(row, idx) {
    return normalizeCellText(row?.[idx] ?? '');
  }

  function isDateLikeValue(v) {
    if (v instanceof Date) return true;
    if (typeof v === 'number') return v > 20000 && v < 60000;
    const s = normalizeCellText(v);
    return /^\d{4}[-/.]\d{1,2}([-/.\d]+)?$/.test(s) || /^\d{4}年\d{1,2}月/.test(s);
  }

  function parseDateMaybe(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === 'number' && v > 20000 && v < 60000) return excelSerialToDate(Math.floor(v));
    if (typeof v === 'number') return '';
    const s = normalizeCellText(v);
    const m = s.match(/(\d{4})[./年-](\d{1,2})[./月-](\d{1,2})/);
    if (m) {
      const month = Number(m[2]);
      const day = Number(m[3]);
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) return `${m[1]}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
    const ym = s.match(/(\d{4})[./年-](\d{1,2})/);
    if (ym) {
      const month = Number(ym[2]);
      if (month >= 1 && month <= 12) return `${ym[1]}-${String(month).padStart(2, '0')}-01`;
    }
    return s.slice(0, 10);
  }

  function isTimeLabel(v) {
    if (typeof v === 'number') return parseTimeMinutes(v) !== null;
    const s = normalizeCellText(v);
    if (!s || /日期|时间|时点|户号|计量点|合计|求和项/i.test(s)) return false;
    return /^\d{1,2}[:：]\d{2}$/.test(s) || /^\d{1,2}[:：]\d{2}\s*(?:~|-|至)\s*\d{1,2}[:：]\d{2}$/.test(s);
  }

  function parseTimeMinutes(v) {
    if (typeof v === 'number') {
      if (v > 1) {
        const fraction = v - Math.floor(v);
        if (fraction <= 0) return null;
        const mins = Math.round(fraction * 1440);
        return Math.abs(fraction * 1440 - mins) < 0.0001 && mins % 15 === 0 ? mins % 1440 : null;
      }
      if (v >= 0 && v <= 1) {
        const mins = v === 1 ? 1440 : Math.round(v * 1440);
        return Math.abs(v * 1440 - mins) < 0.0001 && mins % 15 === 0 ? mins : null;
      }
    }
    const s = normalizeCellText(v);
    if (!s) return null;
    const inner = s.match(/(\d{1,2})[:：](\d{2})\s*(?:~|-|至)?/);
    if (inner) return Number(inner[1]) * 60 + Number(inner[2]);
    if (s === '24:00') return 24 * 60;
    const m = s.match(/^(\d{1,2})(?::(\d{2}))?$/);
    if (m) return Number(m[1]) * 60 + Number(m[2] || 0);
    const hm = s.match(/^(\d{1,2})[:：](\d{2})$/);
    if (hm) return Number(hm[1]) * 60 + Number(hm[2]);
    return null;
  }

  function parseTimeRangeMaybe(v) {
    const s = normalizeCellText(v);
    if (!s) return null;
    const m = s.match(/(\d{1,2})[:：](\d{2})\s*(?:~|-|至)\s*(\d{1,2})[:：](\d{2})/);
    if (!m) return null;
    const start = Number(m[1]) * 60 + Number(m[2]);
    const end = Number(m[3]) * 60 + Number(m[4]);
    if (!Number.isFinite(start) || !Number.isFinite(end)) return null;
    return { start, end };
  }

  function isLoadPointHeader(v) {
    if (isTimeLabel(v)) return true;
    const s = normalizeCellText(v);
    if (!s || /日期|时间|时点|户号|计量点|合计|求和项/i.test(s)) return false;
    return /^\d{1,2}$/.test(s) && Number(s) >= 0 && Number(s) <= 23;
  }

  function parseLoadNumber(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    const text = normalizeCellText(v).replace(/,/g, '');
    if (!text) return null;
    const n = Number(text.replace(/[^\d.-]/g, ''));
    return Number.isFinite(n) ? n : null;
  }

  function timeIndexFromMinutes(mins, hasZero) {
    if (!Number.isFinite(mins)) return null;
    if (mins === 1440) return 95;
    const idx = hasZero ? Math.round(mins / 15) : Math.round(mins / 15) - 1;
    return idx >= 0 && idx < 96 ? idx : null;
  }

  function build96FromTimedPairs(pairs, options = {}) {
    const valid = (pairs || []).filter((p) => Number.isFinite(p?.mins) && Number.isFinite(p?.value));
    if (!valid.length) return { points: [], validCount: 0 };
    const uniqueMins = [...new Set(valid.map((p) => p.mins))].sort((a, b) => a - b);
    const diffs = uniqueMins.slice(1).map((mins, idx) => mins - uniqueMins[idx]).filter((d) => d > 0 && d % 15 === 0);
    const interval = diffs.length ? Math.min(...diffs) : null;
    const hasZero = uniqueMins.includes(0);
    const startAtTime = Boolean(options.startAtTime || hasZero);
    const preferExactTime = Boolean(options.preferExactTime);
    const points = Array(96).fill(null);
    const counts = Array(96).fill(0);
    for (const pair of valid) {
      if (!preferExactTime && Number.isFinite(interval) && interval > 15 && interval <= 60) {
        const quarters = Math.round(interval / 15);
        const startMins = startAtTime ? pair.mins : pair.mins - interval;
        const startIdx = Math.round(startMins / 15);
        if (startIdx < 0 || startIdx >= 96) continue;
        const each = Number(pair.value || 0) / quarters;
        for (let offset = 0; offset < quarters; offset++) {
          const idx = startIdx + offset;
          if (idx < 0 || idx >= 96) continue;
          points[idx] = Number(points[idx] || 0) + each;
          counts[idx] += 1;
        }
      } else {
        const idx = timeIndexFromMinutes(pair.mins, hasZero);
        if (idx === null) continue;
        points[idx] = Number(points[idx] || 0) + Number(pair.value || 0);
        counts[idx] += 1;
      }
    }
    return { points: points.map((v) => v ?? 0), validCount: counts.filter(Boolean).length };
  }

  function extractWideLoadPoints(row, pointCols, headerRow) {
    const timed = pointCols.map((c) => ({ mins: parseTimeMinutes(headerRow?.[c]), value: parseLoadNumber(row?.[c]) }));
    const mapped = build96FromTimedPairs(timed);
    if (mapped.validCount >= Math.max(8, Math.floor(pointCols.length * 0.5))) return mapped;
    const raw = pointCols.map((c) => parseLoadNumber(row?.[c])).filter((v) => Number.isFinite(v));
    return { points: resampleSeries(raw, 96), validCount: raw.length };
  }

  function chooseLoadDailyTotal(rawTotal, pointTotal) {
    const total = Number(rawTotal);
    const computed = Number(pointTotal || 0);
    if (!Number.isFinite(total)) return round(computed, 3);
    if (!computed) return round(total, 3);
    if (Math.abs(total - computed) <= Math.max(0.01, Math.abs(computed) * 0.05)) return round(total, 3);
    if (Math.abs(total / 1000 - computed) <= Math.max(0.01, Math.abs(computed) * 0.05)) return round(total / 1000, 3);
    if (Math.abs(total * 1000 - computed) <= Math.max(0.01, Math.abs(computed) * 0.05)) return round(total * 1000, 3);
    return round(computed, 3);
  }

  function sheetNameBonus(name) {
    const s = String(name || '').toLowerCase();
    if (/mwh/.test(s)) return 160;
    if (/kwh/.test(s)) return 140;
    if (/15min/.test(s)) return 120;
    if (/30min/.test(s)) return 100;
    if (/1h/.test(s)) return 80;
    if (/1day/.test(s)) return 60;
    if (/sheet2/.test(s)) return 50;
    return 0;
  }

  function normalizeLoadPoints(points, tableName = '', sourceName = '') {
    const source = Array.isArray(points) ? points : [];
    const nums = source.map((v) => Number(v)).filter((v) => Number.isFinite(v));
    if (!nums.length) return [];
    const max = Math.max(...nums);
    const median = nums.slice().sort((a, b) => a - b)[Math.floor(nums.length / 2)] ?? 0;
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    const hint = `${tableName} ${sourceName}`.toLowerCase();
    const explicitMwh = /mwh|兆瓦时/.test(hint);
    const explicitKwh = /kwh|千瓦时/.test(hint);
    const explicitKw = /(^|[^a-z])kw([^a-z]|$)|千瓦/.test(hint);
    const magnitudeLooksKwh = max > 100 || median > 20 || mean > 20;
    const shouldScale = !explicitMwh && (explicitKwh || explicitKw || magnitudeLooksKwh);
    const out = source.map((v) => {
      const n = Number(v);
      if (!Number.isFinite(n)) return 0;
      return shouldScale ? round(n / 1000, 6) : round(n, 6);
    });
    return out;
  }

  function isGenericUserName(name = '') {
    const s = normalizeCellText(name).replace(/\s+/g, '');
    return !s || /^未识别用户$/.test(s) || /^曲线导出/.test(s) || /^导入/.test(s) || /^用户日用电量统计$/.test(s) || /^用户负荷数据$/.test(s) || /^负荷数据$/.test(s) || /^sheet\d*$/i.test(s) || /^kwh|mwh$/i.test(s) || /^(市场成员名称|用户名称|客户名称|用电单元名称|计量点|户号)$/.test(s);
  }

  function inferUserNameFromSource(sourceName = '') {
    const parts = String(sourceName || '').replace(/\\/g, '/').split('/').filter(Boolean);
    const base = (parts[parts.length - 1] || '').replace(/\.[^.]+$/, '');
    const baseCandidate = base
      .replace(/\d{4}[-_./年]\d{1,2}[-_./月]\d{1,2}.*$/g, '')
      .replace(/用电明细|曲线数据|负荷数据/g, '')
      .trim();
    if (baseCandidate && !/^曲线导出|^导入|^sheet/i.test(baseCandidate)) return baseCandidate;
    for (let i = parts.length - 2; i >= 0; i--) {
      const p = parts[i].replace(/\.[^.]+$/, '');
      if (p && !/^曲线导出|^用电明细|^导入|^sheet/i.test(p)) return p;
    }
    const cleaned = base.replace(/\.[^.]+$/, '').replace(/曲线导出.*$/, '').trim();
    return cleaned || '未识别用户';
  }

  function inferUserNameFromWorkbook(tables, sourceName = '') {
    const sourceNameGuess = inferUserNameFromSource(sourceName);
    if (!isGenericUserName(sourceNameGuess)) return sourceNameGuess;
    for (const table of tables || []) {
      const fromTable = inferUserNameFromTable(table);
      if (!isGenericUserName(fromTable)) return fromTable;
    }
    const account = inferAccountNoFromWorkbook(tables);
    return account ? `户号${account}` : sourceNameGuess;
  }

  function inferUserNameFromTable(table) {
    const matrix = tableToMatrix(table);
    const labelPattern = /用户名称|客户名称|用电客户|市场成员名称|企业名称|户名|名称/;
    for (let r = 0; r < Math.min(matrix.length, 30); r++) {
      const row = matrix[r] || [];
      for (let c = 0; c < row.length; c++) {
        const text = normalizeCellText(row[c]);
        if (!labelPattern.test(text)) continue;
        const candidates = [row[c + 1], row[c + 2], matrix[r + 1]?.[c], matrix[r + 1]?.[c + 1]];
        for (const candidate of candidates) {
          const cleaned = cleanUserNameCandidate(candidate);
          if (cleaned) return cleaned;
        }
      }
    }
    for (let r = 0; r < Math.min(matrix.length, 20); r++) {
      for (const value of matrix[r] || []) {
        const cleaned = cleanUserNameCandidate(value);
        if (cleaned && /公司|厂|集团|中心|商贸|工贸|陶瓷|煤|矿|焦化|电力/.test(cleaned)) return cleaned;
      }
    }
    return '';
  }

  function cleanUserNameCandidate(value) {
    const s = normalizeCellText(value).replace(/^[：:\s]+|[：:\s]+$/g, '').trim();
    if (!s || s.length < 2 || s.length > 80) return '';
    if (/^\d+$/.test(s) || /^\d{4}[-/.年]/.test(s) || /^\d{1,2}[:：]\d{2}/.test(s)) return '';
    if (/日期|时间|时点|电量|负荷|合计|求和项|户号|编号|序号|计量点|市场成员名称|用户名称|客户名称|用电单元名称|sheet|kwh|mwh/i.test(s)) return '';
    return s;
  }

  function findHeaderColumn(headerRow, pattern) {
    const row = headerRow || [];
    for (let i = 0; i < row.length; i++) {
      const text = normalizeCellText(row[i]).replace(/\s+/g, '');
      if (pattern.test(text)) return i;
    }
    return -1;
  }

  function cleanIdentifier(value) {
    const text = normalizeCellText(value).trim();
    return text && !/^(户号|计量点|合同号|编号|序号)$/i.test(text) ? text : '';
  }

  function inferAccountNoFromWorkbook(tables) {
    for (const table of tables || []) {
      const account = inferAccountNoFromTable(table);
      if (account) return account;
    }
    return '';
  }

  function inferAccountNoFromTable(table) {
    const matrix = tableToMatrix(table);
    for (let r = 0; r < Math.min(matrix.length, 40); r++) {
      for (const value of matrix[r] || []) {
        const text = normalizeCellText(value).replace(/\D/g, '');
        if (/^\d{12,24}$/.test(text)) return text;
      }
    }
    return '';
  }

  function findBestValueColumn(rows, excludeCols = new Set()) {
    const counts = new Map();
    for (const row of rows || []) {
      for (const [key, value] of Object.entries(row || {})) {
        const idx = colIndex(key);
        if (excludeCols.has(idx)) continue;
        const n = parseLoadNumber(value);
        if (!Number.isFinite(n)) continue;
        if (!counts.has(idx)) counts.set(idx, []);
        counts.get(idx).push(n);
      }
    }
    let best = null;
    for (const [idx, values] of counts.entries()) {
      if (values.length < 3) continue;
      const sorted = [...values].sort((a, b) => a - b);
      const median = sorted[Math.floor(sorted.length / 2)] ?? 0;
      const avg = values.reduce((a, b) => a + b, 0) / values.length;
      const idLikeRatio = values.filter((n) => Math.abs(n) >= 1e10 && Number.isInteger(n)).length / values.length;
      const dateLikeRatio = values.filter((n) => n > 20000 && n < 60000).length / values.length;
      const binaryRatio = values.filter((n) => n === 0 || n === 1).length / values.length;
      const uniqueCount = new Set(values.map((n) => round(n, 6))).size;
      if (idLikeRatio > 0.25 || dateLikeRatio > 0.5 || (binaryRatio > 0.85 && uniqueCount <= 2)) continue;
      const magnitude = Math.min(Math.abs(median), 1000) + Math.min(Math.abs(avg), 1000);
      const score = values.length * 12 + uniqueCount * 8 + magnitude;
      if (!best || score > best.score) best = { idx, score };
    }
    return best?.idx ?? null;
  }

  function selectBestStructuredSheet(tables) {
    let best = null;
    for (const table of tables || []) {
      const matrix = tableToMatrix(table);
      const rows = Array.isArray(table?.rows) ? table.rows : matrix.map((row) => matrixRowToObject(row));
      const text = normalizeCellText(table?.text || '');
      const numericCount = matrix.flat().filter((v) => typeof v === 'number' || /^-?\d+(?:\.\d+)?$/.test(String(v ?? ''))).length;
      const score = rows.length * 8 + (matrix[0]?.length || 0) * 2 + numericCount + sheetNameBonus(table?.name) + (/电压|月份|日期|时段|时点|负荷|购电/.test(text) ? 200 : 0);
      if (!best || score > best.score) best = { ...table, rows, matrix, score };
    }
    return best;
  }

  function selectBestLoadCandidate(tables, sourceName = '') {
    const preferred = (tables || []).filter((table) => /15min/i.test(String(table?.name || '')));
    for (const table of preferred) {
      const interval = parseIntervalLoadTable(table, sourceName);
      if (interval) return interval;
    }
    const scanTables = preferred.length ? preferred : (tables || []);
    let best = null;
    for (const table of scanTables) {
      const canonical = parseCanonicalLoadTable(table, sourceName);
      if (canonical && (!best || canonical.score > best.score)) best = canonical;
      const wide = parseWideLoadTable(table, sourceName);
      if (wide && (!best || wide.score > best.score)) best = wide;
      const interval = parseIntervalLoadTable(table, sourceName);
      if (interval && (!best || interval.score > best.score)) best = interval;
      const long = parseLongLoadTable(table, sourceName);
      if (long && (!best || long.score > best.score)) best = long;
    }
    return best;
  }

  function parseCanonicalLoadTable(table, sourceName = '') {
    const matrix = tableToMatrix(table);
    if (!matrix.length) return null;
    let headerIndex = -1;
    let meta = null;
    for (let r = 0; r < Math.min(matrix.length, 20); r++) {
      const headerRow = matrix[r] || [];
      const userCol = findHeaderColumn(headerRow, /^(市场成员名称|用户名称|客户名称|用电客户|企业名称|户名|名称)$/);
      const dateCol = findHeaderColumn(headerRow, /^(日期|用电日期|数据日期)$/);
      const pointCols = [];
      headerRow.forEach((v, idx) => { if (isLoadPointHeader(v)) pointCols.push(idx); });
      if (userCol >= 0 && dateCol >= 0 && pointCols.length >= 24) {
        headerIndex = r;
        meta = {
          userCol,
          dateCol,
          accountCol: findHeaderColumn(headerRow, /^户号$/),
          meterCol: findHeaderColumn(headerRow, /^(计量点|计量点编号|计量点号)$/),
          signedCol: findHeaderColumn(headerRow, /是否签约|当前日期是否签约/),
          unitCol: findHeaderColumn(headerRow, /所属交易单元名称|交易单元|用电类别|计量方式/),
          totalCol: findHeaderColumn(headerRow, /日合计值|日合计|合计值|总计/),
          pointCols
        };
        break;
      }
    }
    if (!meta) return null;
    const records = [];
    let carryDate = '';
    for (let r = headerIndex + 1; r < matrix.length; r++) {
      const row = matrix[r] || [];
      const userName = cleanUserNameCandidate(row[meta.userCol]);
      if (!userName || /^总计|合计$/i.test(userName)) continue;
      const parsedDate = parseDateMaybe(row[meta.dateCol]);
      if (/^\d{4}-\d{2}-\d{2}$/.test(parsedDate)) carryDate = parsedDate;
      const date = /^\d{4}-\d{2}-\d{2}$/.test(parsedDate) ? parsedDate : carryDate;
      if (!date) continue;
      const extracted = extractWideLoadPoints(row, meta.pointCols, matrix[headerIndex]);
      if (extracted.validCount < Math.max(24, Math.floor(meta.pointCols.length * 0.5))) continue;
      const points = normalizeLoadPoints(extracted.points, table?.name || '', sourceName);
      const pointTotal = round(points.reduce((a, b) => a + Number(b || 0), 0), 6);
      const rawTotal = meta.totalCol >= 0 ? parseLoadNumber(row[meta.totalCol]) : null;
      records.push({
        userName,
        accountNo: meta.accountCol >= 0 ? cleanIdentifier(row[meta.accountCol]) : '',
        contractNo: meta.meterCol >= 0 ? cleanIdentifier(row[meta.meterCol]) : '',
        date,
        signed: meta.signedCol >= 0 ? normalizeCellText(row[meta.signedCol]) : '',
        meterType: meta.unitCol >= 0 ? normalizeCellText(row[meta.unitCol]) : '',
        dailyTotal: chooseLoadDailyTotal(rawTotal, pointTotal),
        points
      });
    }
    if (!records.length) return null;
    const userCount = new Set(records.map((r) => r.userName)).size;
    return {
      headers: defaultLoadHeaders(),
      records,
      sourceName: userCount > 1 ? '多用户负荷数据' : records[0].userName,
      score: records.length * 150 + userCount * 500 + meta.pointCols.length * 10 + sheetNameBonus(table?.name) + 10000
    };
  }

  function parseWideLoadTable(table, sourceName = '') {
    const matrix = tableToMatrix(table);
    if (!matrix.length) return null;
    const tableUserName = inferUserNameFromTable(table) || inferUserNameFromSource(sourceName);
    const accountNo = inferAccountNoFromTable(table);
    let headerIndex = -1;
    let pointCols = [];
    for (let r = 0; r < Math.min(matrix.length, 12); r++) {
      const row = matrix[r] || [];
      const cols = [];
      row.forEach((v, idx) => { if (isLoadPointHeader(v)) cols.push(idx); });
      if (cols.length > pointCols.length) {
        headerIndex = r;
        pointCols = cols;
      }
    }
    if (pointCols.length < 8) return null;
    const headerRow = matrix[headerIndex] || [];
    const userCol = findHeaderColumn(headerRow, /^(市场成员名称|用户名称|客户名称|用电客户|企业名称|户名|名称)$/);
    const accountCol = findHeaderColumn(headerRow, /^户号$/);
    const meterCol = findHeaderColumn(headerRow, /^(计量点|计量点编号|计量点号)$/);
    const dateHeaderCol = findHeaderColumn(headerRow, /^(日期|用电日期|数据日期)$/);
    const signedCol = findHeaderColumn(headerRow, /是否签约|当前日期是否签约/);
    const unitCol = findHeaderColumn(headerRow, /所属交易单元名称|交易单元|用电类别|计量方式/);
    const totalCol = findHeaderColumn(headerRow, /日合计值|日合计|合计值|总计/);
    const dataRows = [];
    let carryDate = '';
    for (let r = headerIndex + 1; r < matrix.length; r++) {
      const row = matrix[r] || [];
      if (!row.some((v) => v !== null && v !== undefined && v !== '')) continue;
      const dateCol = dateHeaderCol >= 0 ? dateHeaderCol : row.findIndex((v) => isDateLikeValue(v));
      const parsedDate = parseDateMaybe(dateCol >= 0 ? row[dateCol] : row[0]);
      if (/^\d{4}-\d{2}-\d{2}$/.test(parsedDate)) carryDate = parsedDate;
      const date = /^\d{4}-\d{2}-\d{2}$/.test(parsedDate) ? parsedDate : carryDate;
      if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) continue;
      const rowUserName = userCol >= 0 ? cleanUserNameCandidate(row[userCol]) : '';
      if (userCol >= 0 && !rowUserName) continue;
      if (/^总计|合计$/i.test(rowUserName)) continue;
      const extracted = extractWideLoadPoints(row, pointCols, matrix[headerIndex]);
      if (extracted.validCount < Math.max(8, Math.floor(pointCols.length * 0.5))) continue;
      const points = normalizeLoadPoints(extracted.points, table?.name || '', sourceName);
      const total = totalCol >= 0 ? parseLoadNumber(row[totalCol]) : null;
      const pointTotal = round(points.reduce((a, b) => a + Number(b || 0), 0), 6);
      dataRows.push({
        userName: rowUserName || tableUserName,
        accountNo: accountCol >= 0 ? cleanIdentifier(row[accountCol]) : accountNo,
        contractNo: meterCol >= 0 ? cleanIdentifier(row[meterCol]) : '',
        date,
        signed: signedCol >= 0 ? normalizeCellText(row[signedCol]) : '',
        meterType: unitCol >= 0 ? normalizeCellText(row[unitCol]) : '',
        dailyTotal: chooseLoadDailyTotal(total, pointTotal),
        points
      });
    }
    if (!dataRows.length) return null;
    return {
      headers: defaultLoadHeaders().slice(0, 7).concat(pointCols.map((idx) => normalizeCellText(matrix[headerIndex]?.[idx] ?? '')).filter(Boolean)),
      records: dataRows,
      sourceName: dataRows.length > 1 ? '多用户负荷数据' : tableUserName,
      score: dataRows.length * 100 + pointCols.length * 5 + sheetNameBonus(table?.name)
    };
  }

  function parseIntervalLoadTable(table, sourceName = '') {
    const rows = Array.isArray(table?.rows) ? table.rows : [];
    if (!rows.length) return null;
    const matrix = tableToMatrix(table);
    if (!matrix.length) return null;
    const timeCol = 0;
    const valueCol = Math.min(2, (matrix[0] || []).length - 1);
    if (valueCol < 1) return null;
    const dataRows = [];
    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r] || [];
      const timeText = normalizeCellText(row[timeCol] ?? row[0]);
      const range = parseTimeRangeMaybe(timeText);
      const date = parseDateMaybe(timeText) || parseDateMaybe(row.find((v) => isDateLikeValue(v)));
      const n = parseLoadNumber(row[valueCol]);
      if (!range || !date || !Number.isFinite(n)) continue;
      dataRows.push({ date, mins: range.start, value: n });
    }
    if (dataRows.length < 24) return null;
    const grouped = new Map();
    for (const item of dataRows) {
      if (!grouped.has(item.date)) grouped.set(item.date, []);
      grouped.get(item.date).push({ mins: item.mins, value: item.value });
    }
    const records = [];
    const tableUserName = inferUserNameFromTable(table) || inferUserNameFromSource(sourceName);
    for (const [date, items] of grouped.entries()) {
      items.sort((a, b) => a.mins - b.mins);
      const points = build96FromTimedPairs(items, { preferExactTime: true }).points;
      if (points.length < 96) continue;
      const mwhPoints = points.map((v) => round(Number(v || 0) / 1000, 6));
      records.push({
        userName: tableUserName,
        accountNo: inferAccountNoFromTable(table),
        contractNo: '',
        date,
        signed: '',
        meterType: '',
        dailyTotal: round(mwhPoints.reduce((a, b) => a + Number(b || 0), 0), 6),
        points: mwhPoints
      });
    }
    if (!records.length) return null;
    return {
      headers: defaultLoadHeaders(),
      records,
      sourceName: tableUserName,
      score: records.length * 180 + sheetNameBonus(table?.name) + 12000
    };
  }

  function findLongLoadHeaderMeta(matrix) {
    for (let r = 0; r < Math.min(matrix.length, 30); r++) {
      const row = matrix[r] || [];
      const userCol = findCombinedHeaderColumn(matrix, r, /^(市场成员名称|用户名称|客户名称|用电客户|企业名称|户名|名称)$/);
      const identityCol = findCombinedHeaderColumn(matrix, r, /电能表资产号|电表资产号|资产号|表号|电能表编号/);
      const dateCol = findCombinedHeaderColumn(matrix, r, /^(日期|用电日期|数据日期)$/);
      const timeCol = findCombinedHeaderColumn(matrix, r, /^(时点|时间|时段|数据时间|采集时间)$/);
      const accountCol = findCombinedHeaderColumn(matrix, r, /^户号$/);
      const meterCol = findCombinedHeaderColumn(matrix, r, /^(计量点|计量点编号|计量点号|电能表资产号|电表资产号|资产号|表号|电能表编号)$/);
      let valueCol = -1;
      for (let c = 0; c < row.length; c++) {
        const text = combinedHeaderText(matrix, r, c).replace(/\s+/g, '');
        const isEnergy = /电量|有功电量|正向有功/.test(text) && !/电费|电价|功率|无功/.test(text);
        const isTotalActivePower = /总有功功率/.test(text);
        if (isEnergy || isTotalActivePower) {
          valueCol = c;
          break;
        }
      }
      if (dateCol >= 0 && timeCol >= 0 && valueCol >= 0) {
        return {
          headerIndex: r,
          userCol,
          identityCol,
          dateCol,
          timeCol,
          accountCol,
          meterCol,
          valueCol,
          valueHeader: combinedHeaderText(matrix, r, valueCol)
        };
      }
    }
    return null;
  }

  function combinedHeaderText(matrix, rowIndex, colIndexValue) {
    return [...new Set([matrix[rowIndex - 1]?.[colIndexValue], matrix[rowIndex]?.[colIndexValue]]
      .map(normalizeCellText)
      .filter(Boolean))]
      .join(' ');
  }

  function findCombinedHeaderColumn(matrix, rowIndex, pattern) {
    const row = matrix[rowIndex] || [];
    for (let i = 0; i < row.length; i++) {
      const text = combinedHeaderText(matrix, rowIndex, i).replace(/\s+/g, '');
      if (pattern.test(text)) return i;
    }
    return -1;
  }

  function parseLongLoadTable(table, sourceName = '') {
    if (/15min/i.test(String(table?.name || ''))) return null;
    const rows = Array.isArray(table?.rows) ? table.rows : [];
    const matrix = tableToMatrix(table);
    if (!rows.length && !matrix.length) return null;
    const tableUserName = inferUserNameFromTable(table) || inferUserNameFromSource(sourceName);
    const accountNo = inferAccountNoFromTable(table);
    const meta = findLongLoadHeaderMeta(matrix);
    if (meta) {
      const grouped = new Map();
      const tableTimeStartsAtZero = matrix.slice(meta.headerIndex + 1).some((row) => parseTimeMinutes(row?.[meta.timeCol]) === 0);
      for (let r = meta.headerIndex + 1; r < matrix.length; r++) {
        const row = matrix[r] || [];
        const date = parseDateMaybe(row[meta.dateCol]);
        const mins = parseTimeMinutes(row[meta.timeCol]);
        const n = parseLoadNumber(row[meta.valueCol]);
        const rowUserName = meta.userCol >= 0 ? cleanUserNameCandidate(row[meta.userCol]) : '';
        const identityName = meta.identityCol >= 0 ? cleanIdentifier(row[meta.identityCol]) : '';
        const userName = rowUserName || identityName || tableUserName;
        if (!userName || isGenericUserName(userName) || !/^\d{4}-\d{2}-\d{2}$/.test(date) || !Number.isFinite(mins) || !Number.isFinite(n)) continue;
        const key = `${userName}__${date}`;
        if (!grouped.has(key)) {
          grouped.set(key, {
            userName,
            accountNo: meta.accountCol >= 0 ? cleanIdentifier(row[meta.accountCol]) : accountNo,
            contractNo: meta.meterCol >= 0 ? cleanIdentifier(row[meta.meterCol]) : identityName,
            date,
            items: [],
            meterNos: new Set()
          });
        }
        const item = grouped.get(key);
        if (!item.accountNo && meta.accountCol >= 0) item.accountNo = cleanIdentifier(row[meta.accountCol]);
        const meterNo = meta.meterCol >= 0 ? cleanIdentifier(row[meta.meterCol]) : '';
        if (meterNo) item.meterNos.add(meterNo);
        item.items.push({ mins, value: n });
      }
      const records = [];
      for (const item of grouped.values()) {
        item.items.sort((a, b) => a.mins - b.mins);
      const timedPoints = build96FromTimedPairs(item.items, { startAtTime: tableTimeStartsAtZero, preferExactTime: true });
        if (timedPoints.validCount < 24) continue;
        const points = normalizeLoadPoints(timedPoints.points, `${table?.name || ''} ${meta.valueHeader || ''}`, sourceName);
        records.push({
          userName: item.userName,
          accountNo: item.accountNo || '',
          contractNo: [...item.meterNos].join('、') || item.contractNo || '',
          date: item.date,
          signed: '',
          meterType: '',
          dailyTotal: round(points.reduce((a, b) => a + Number(b || 0), 0), 6),
          points
        });
      }
      if (records.length) {
        const userCount = new Set(records.map((r) => r.userName)).size;
        return {
          headers: defaultLoadHeaders(),
          records,
          sourceName: userCount > 1 ? '多用户负荷数据' : records[0].userName,
          score: records.length * 120 + userCount * 500 + sheetNameBonus(table?.name) + 8000
        };
      }
    }
    const dateCol = findBestDateColumn(rows);
    let timeCol = findBestTimeColumn(rows, dateCol);
    if (!Number.isFinite(timeCol) && rows.some((row) => parseTimeMinutes(row[colName(dateCol)]) !== null)) timeCol = dateCol;
    const valueCol = findPreferredLoadValueColumn(rows) || findBestValueColumn(rows, new Set([dateCol, timeCol].filter((v) => Number.isFinite(v))));
    if (!Number.isFinite(dateCol) || !Number.isFinite(timeCol) || !Number.isFinite(valueCol)) return null;
    const grouped = new Map();
    const interval = inferLoadIntervalMinutes(table?.name || '');
    let lastDate = '';
    let lastMins = null;
    for (const row of rows) {
      let date = parseDateMaybe(row[colName(dateCol)]);
      let mins = parseTimeMinutes(row[colName(timeCol)]);
      if ((!date || !/^\d{4}-\d{2}-\d{2}$/.test(date) || !Number.isFinite(mins)) && lastDate && Number.isFinite(lastMins) && Number.isFinite(interval)) {
        const nextMinsRaw = lastMins + interval;
        date = nextMinsRaw >= 1440 ? addDays(lastDate, 1) : lastDate;
        mins = nextMinsRaw % 1440;
      }
      const rawValue = row[colName(valueCol)];
      const n = parseLoadNumber(rawValue);
      if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date) || !Number.isFinite(mins) || !Number.isFinite(n)) continue;
      lastDate = date;
      lastMins = mins;
      if (!grouped.has(date)) grouped.set(date, []);
      grouped.get(date).push({ mins, value: n });
    }
    if (!grouped.size) return null;
    const records = [];
    for (const [date, items] of grouped.entries()) {
      items.sort((a, b) => a.mins - b.mins);
      const timedPoints = build96FromTimedPairs(items, { preferExactTime: /15min/i.test(String(table?.name || '')) });
      const rawPoints = timedPoints.validCount >= 8 ? timedPoints.points : items.map((it) => it.value);
      const expectedPoints = Number.isFinite(interval) ? 1440 / interval : rawPoints.length;
      const availablePoints = timedPoints.validCount || rawPoints.length;
      if (expectedPoints > 1 && availablePoints < expectedPoints * 0.75) continue;
      const points = normalizeLoadPoints(rawPoints.length === 96 ? rawPoints : resampleSeries(rawPoints, rawPoints.length > 24 ? 96 : (rawPoints.length === 24 ? 96 : rawPoints.length)), table?.name || '', sourceName);
      records.push({
        userName: tableUserName,
        accountNo,
        contractNo: '',
        date,
        signed: '',
        meterType: '',
          dailyTotal: round(points.reduce((a, b) => a + Number(b || 0), 0), 6),
        points
      });
    }
    return {
      headers: defaultLoadHeaders(),
      records,
      sourceName: tableUserName,
      score: records.length * 80 + sheetNameBonus(table?.name) + (timeCol >= 0 ? 20 : 0)
    };
  }

  function findBestDateColumn(rows) {
    const counts = new Map();
    for (const row of rows || []) {
      for (const [key, value] of Object.entries(row || {})) {
        const idx = colIndex(key);
        const d = parseDateMaybe(value);
        if (/^\d{4}-\d{2}-\d{2}$/.test(d)) counts.set(idx, (counts.get(idx) || 0) + 1);
      }
    }
    let best = null;
    for (const [idx, count] of counts.entries()) {
      if (!best || count > best.count) best = { idx, count };
    }
    return best?.idx ?? 1;
  }

  function findBestTimeColumn(rows, dateCol) {
    const counts = new Map();
    for (const row of rows || []) {
      for (const [key, value] of Object.entries(row || {})) {
        const idx = colIndex(key);
        if (idx === dateCol) continue;
        const text = normalizeCellText(value);
        const isNumericTime = typeof value === 'number' && value >= 0 && value <= 1 && parseTimeMinutes(value) !== null;
        const isTextTime = typeof value !== 'number' && /:|~|至/.test(text) && parseTimeMinutes(value) !== null;
        if (isNumericTime || isTextTime) counts.set(idx, (counts.get(idx) || 0) + 1);
      }
    }
    let best = null;
    for (const [idx, count] of counts.entries()) {
      if (!best || count > best.count) best = { idx, count };
    }
    return best && best.count >= 3 ? best.idx : null;
  }

  function findPreferredLoadValueColumn(rows) {
    const sample = rows.slice(0, 8);
    for (const row of sample) {
      for (const [key, value] of Object.entries(row || {})) {
        const text = normalizeCellText(value).replace(/\s+/g, '');
        if ((/电量|有功电量|正向有功/.test(text) && !/电费|电价|功率|无功/.test(text)) || /总有功功率/.test(text)) return colIndex(key);
      }
    }
    return null;
  }

  function inferLoadIntervalMinutes(name = '') {
    const s = String(name || '').toLowerCase();
    if (/15/.test(s)) return 15;
    if (/30/.test(s)) return 30;
    if (/1h|60/.test(s)) return 60;
    return null;
  }

  function addDays(date, days) {
    const d = new Date(`${date}T00:00:00Z`);
    d.setUTCDate(d.getUTCDate() + Number(days || 0));
    return d.toISOString().slice(0, 10);
  }

  async function readXlsTables(buf, sourceName = '') {
    const bytes = new Uint8Array(buf || []);
    if (bytes.length < 8 || bytes[0] !== 0xd0 || bytes[1] !== 0xcf || bytes[2] !== 0x11 || bytes[3] !== 0xe0) throw new Error('不是有效的Excel文件');
    const doc = parseCompoundFile(bytes);
    const workbookEntry = doc.entries.find((e) => /^(workbook|book)$/i.test(e.name));
    if (!workbookEntry) throw new Error('没有找到工作簿数据');
    const workbookBytes = readCompoundStream(doc, workbookEntry);
    return parseBiffWorkbook(workbookBytes, sourceName);
  }

  function parseCompoundFile(bytes) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const sectorShift = dv.getUint16(30, true);
    const miniSectorShift = dv.getUint16(32, true);
    const sectorSize = 1 << sectorShift;
    const miniSectorSize = 1 << miniSectorShift;
    const fatSectorCount = dv.getUint32(44, true);
    const firstDirSector = dv.getUint32(48, true);
    const miniCutoff = dv.getUint32(56, true);
    const firstMiniFatSector = dv.getUint32(60, true);
    const miniFatSectorCount = dv.getUint32(64, true);
    const firstDifatSector = dv.getUint32(68, true);
    const difatSectorCount = dv.getUint32(72, true);
    const difat = [];
    for (let i = 0; i < 109; i++) {
      const v = dv.getUint32(76 + i * 4, true);
      if (v !== 0xffffffff) difat.push(v);
    }
    const fat = [];
    const readSector = (sector) => bytes.slice((sector + 1) * sectorSize, (sector + 2) * sectorSize);
    let nextDifat = firstDifatSector;
    for (let i = 0; i < fatSectorCount; i++) {
      const sector = difat[i];
      if (sector === undefined || sector === 0xffffffff) break;
      const sec = readSector(sector);
      const view = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
      for (let j = 0; j < sectorSize / 4; j++) fat.push(view.getUint32(j * 4, true));
    }
    for (let i = 0; i < difatSectorCount && nextDifat !== 0xffffffff; i++) {
      const sec = readSector(nextDifat);
      const view = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
      const entriesPerSector = sectorSize / 4 - 1;
      for (let j = 0; j < entriesPerSector; j++) {
        const v = view.getUint32(j * 4, true);
        if (v !== 0xffffffff) {
          const fatSector = readSector(v);
          const fatView = new DataView(fatSector.buffer, fatSector.byteOffset, fatSector.byteLength);
          for (let k = 0; k < sectorSize / 4; k++) fat.push(fatView.getUint32(k * 4, true));
        }
      }
      nextDifat = view.getUint32(sectorSize - 4, true);
    }
    const dirBytes = readChain(bytes, fat, firstDirSector, sectorSize);
    const entries = [];
    for (let off = 0; off + 128 <= dirBytes.length; off += 128) {
      const nameLen = new DataView(dirBytes.buffer, dirBytes.byteOffset + off, 128).getUint16(64, true);
      const type = dirBytes[off + 66];
      const startSector = new DataView(dirBytes.buffer, dirBytes.byteOffset + off, 128).getUint32(116, true);
      const sizeLow = new DataView(dirBytes.buffer, dirBytes.byteOffset + off, 128).getUint32(120, true);
      const name = decodeUtf16le(dirBytes.slice(off, off + Math.max(0, nameLen - 2)));
      if (name) entries.push({ name: name.replace(/\u0000/g, '').trim(), type, startSector, size: sizeLow });
    }
    const miniFat = firstMiniFatSector !== 0xffffffff ? readChain(bytes, fat, firstMiniFatSector, sectorSize) : new Uint8Array(0);
    return { bytes, fat, entries, sectorSize, miniSectorSize, miniCutoff, miniFat, firstDirSector };
  }

  function readCompoundStream(doc, entry) {
    if (entry.size < doc.miniCutoff && doc.miniFat.length) return readMiniStream(doc, entry);
    return readChain(doc.bytes, doc.fat, entry.startSector, doc.sectorSize).slice(0, entry.size);
  }

  function readMiniStream(doc, entry) {
    const root = doc.entries.find((e) => /root entry/i.test(e.name));
    if (!root) return new Uint8Array(0);
    const bigStream = readChain(doc.bytes, doc.fat, root.startSector, doc.sectorSize);
    const out = [];
    let sector = entry.startSector;
    while (sector !== 0xfffffffe && sector !== 0xffffffff) {
      const start = sector * doc.miniSectorSize;
      out.push(...bigStream.slice(start, start + doc.miniSectorSize));
      const nextIndex = sector * 4;
      sector = new DataView(doc.miniFat.buffer, doc.miniFat.byteOffset, doc.miniFat.byteLength).getUint32(nextIndex, true);
    }
    return new Uint8Array(out).slice(0, entry.size);
  }

  function readChain(bytes, fat, startSector, sectorSize) {
    const chunks = [];
    let sector = startSector;
    const seen = new Set();
    while (sector !== 0xfffffffe && sector !== 0xffffffff && sector >= 0 && !seen.has(sector)) {
      seen.add(sector);
      chunks.push(bytes.slice((sector + 1) * sectorSize, (sector + 2) * sectorSize));
      sector = fat[sector];
      if (!Number.isFinite(sector)) break;
    }
    return concatBytes(chunks);
  }

  function concatBytes(chunks) {
    const size = chunks.reduce((n, c) => n + c.length, 0);
    const out = new Uint8Array(size);
    let offset = 0;
    for (const chunk of chunks) {
      out.set(chunk, offset);
      offset += chunk.length;
    }
    return out;
  }

  function decodeUtf16le(bytes) {
    if (!bytes || !bytes.length) return '';
    const view = bytes.length % 2 === 0 ? bytes : bytes.slice(0, bytes.length - 1);
    return new TextDecoder('utf-16le').decode(view);
  }

  function parseBiffWorkbook(bytes, sourceName = '') {
    const sheets = [];
    const globals = {
      sst: [],
      codepage: 1200
    };
    const records = readBiffRecords(bytes);
    let sheetInfo = [];
    for (let i = 0; i < records.length; i++) {
      const rec = records[i];
      if (rec.op === 0x0085) {
        const offset = rec.dataView.getUint32(0, true);
        const name = decodeBiffShortString(rec.data, 6);
        sheetInfo.push({ name: name || 'Sheet', offset });
      } else if (rec.op === 0x00fc) {
        globals.sst = readBiffSst(records, i);
      }
    }
    sheetInfo.sort((a, b) => a.offset - b.offset);
    for (let i = 0; i < sheetInfo.length; i++) {
      const start = sheetInfo[i].offset;
      const end = sheetInfo[i + 1]?.offset ?? bytes.length;
      const sheetRecords = readBiffRecords(bytes.slice(start, end));
      sheets.push(parseBiffSheet(sheetRecords, globals.sst, sheetInfo[i].name, sourceName));
    }
    return sheets.filter((t) => t.rows.length);
  }

  function readBiffRecords(bytes) {
    const out = [];
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let off = 0;
    while (off + 4 <= bytes.length) {
      const op = dv.getUint16(off, true);
      const len = dv.getUint16(off + 2, true);
      const data = bytes.slice(off + 4, off + 4 + len);
      out.push({ op, data, dataView: new DataView(data.buffer, data.byteOffset, data.byteLength) });
      off += 4 + len;
    }
    return out;
  }

  function readBiffSst(records, startIndex) {
    const chunks = [];
    for (let i = startIndex; i < records.length; i++) {
      const rec = records[i];
      if (i !== startIndex && rec.op !== 0x003c && rec.op !== 0x00fc) break;
      if (rec.op === 0x00fc || rec.op === 0x003c) chunks.push(rec.data);
    }
    return parseBiffSstChunks(chunks);
  }

  function parseBiffSstChunks(chunks) {
    if (!chunks?.length || chunks[0].length < 8) return [];
    const firstDv = new DataView(chunks[0].buffer, chunks[0].byteOffset, chunks[0].byteLength);
    const unique = firstDv.getUint32(4, true);
    let chunkIndex = 0;
    let off = 8;
    const strings = [];
    const moveToReadableChunk = () => {
      while (chunkIndex < chunks.length && off >= chunks[chunkIndex].length) {
        chunkIndex += 1;
        off = 0;
      }
      return chunkIndex < chunks.length;
    };
    const readByte = () => {
      if (!moveToReadableChunk()) return null;
      return chunks[chunkIndex][off++];
    };
    const readBytes = (count) => {
      const out = new Uint8Array(count);
      let pos = 0;
      while (pos < count && moveToReadableChunk()) {
        const chunk = chunks[chunkIndex];
        const take = Math.min(count - pos, chunk.length - off);
        out.set(chunk.slice(off, off + take), pos);
        off += take;
        pos += take;
      }
      return pos === count ? out : out.slice(0, pos);
    };
    const readUInt16 = () => {
      const b = readBytes(2);
      return b.length === 2 ? new DataView(b.buffer, b.byteOffset, b.byteLength).getUint16(0, true) : null;
    };
    const readUInt32 = () => {
      const b = readBytes(4);
      return b.length === 4 ? new DataView(b.buffer, b.byteOffset, b.byteLength).getUint32(0, true) : null;
    };
    const skipBytes = (count) => {
      let left = Number(count || 0);
      while (left > 0 && moveToReadableChunk()) {
        const take = Math.min(left, chunks[chunkIndex].length - off);
        off += take;
        left -= take;
      }
    };
    while (strings.length < unique && moveToReadableChunk()) {
      const cch = readUInt16();
      const flags = readByte();
      if (cch === null || flags === null) break;
      let is16 = !!(flags & 0x01);
      const rich = !!(flags & 0x08);
      const asian = !!(flags & 0x04);
      const rt = rich ? (readUInt16() || 0) : 0;
      const ext = asian ? (readUInt32() || 0) : 0;
      let remaining = cch;
      const parts = [];
      while (remaining > 0 && moveToReadableChunk()) {
        const bytesPerChar = is16 ? 2 : 1;
        const available = chunks[chunkIndex].length - off;
        const chars = Math.min(remaining, Math.floor(available / bytesPerChar));
        if (chars > 0) {
          const len = chars * bytesPerChar;
          const textBytes = chunks[chunkIndex].slice(off, off + len);
          parts.push(is16 ? decodeUtf16le(textBytes) : new TextDecoder('latin1').decode(textBytes));
          off += len;
          remaining -= chars;
        }
        if (remaining > 0) {
          chunkIndex += 1;
          off = 0;
          const continueFlags = readByte();
          if (continueFlags === null) break;
          is16 = !!(continueFlags & 0x01);
        }
      }
      skipBytes(rt * 4 + ext);
      strings.push(parts.join(''));
    }
    return strings;
  }

  function decodeBiffUnicodeString(bytes, offset = 0) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let off = offset;
    if (off + 3 > bytes.length) return { text: '', next: bytes.length };
    const cch = dv.getUint16(off, true); off += 2;
    const flags = dv.getUint8(off); off += 1;
    const is16 = !!(flags & 0x01);
    const rich = !!(flags & 0x08);
    const asian = !!(flags & 0x04);
    let rt = 0;
    let ext = 0;
    if (rich && off + 2 <= bytes.length) { rt = dv.getUint16(off, true); off += 2; }
    if (asian && off + 4 <= bytes.length) { ext = dv.getUint32(off, true); off += 4; }
    const byteLen = is16 ? cch * 2 : cch;
    const textBytes = bytes.slice(off, off + byteLen);
    const text = is16 ? decodeUtf16le(textBytes) : new TextDecoder('latin1').decode(textBytes);
    off += byteLen + rt * 4 + ext;
    return { text, next: off };
  }

  function decodeBiffShortString(bytes, offset = 0) {
    if (!bytes || offset + 2 > bytes.length) return '';
    const len = bytes[offset];
    const flags = bytes[offset + 1] || 0;
    const is16 = !!(flags & 0x01);
    const start = offset + 2;
    const textBytes = bytes.slice(start, start + len * (is16 ? 2 : 1));
    return is16 ? decodeUtf16le(textBytes) : new TextDecoder('latin1').decode(textBytes);
  }

  function parseBiffSheet(records, sst, sheetName, sourceName = '') {
    const rows = [];
    const matrix = [];
    const merged = [];
    let currentRow = -1;
    let maxCol = 0;
    for (let i = 0; i < records.length; i++) {
      const rec = records[i];
      if (rec.op === 0x0809 || rec.op === 0x000a) continue;
      if (rec.op === 0x0085) continue;
      if (rec.op === 0x00e5) {
        const dv = rec.dataView;
        const count = dv.getUint16(0, true);
        let off = 2;
        for (let j = 0; j < count; j++) {
          merged.push({
            r1: dv.getUint16(off, true) + 1,
            r2: dv.getUint16(off + 2, true) + 1,
            c1: dv.getUint16(off + 4, true) + 1,
            c2: dv.getUint16(off + 6, true) + 1
          });
          off += 8;
        }
        continue;
      }
      if (rec.op === 0x00bd && rec.data.length >= 8) {
        const rowIndex = rec.dataView.getUint16(0, true) + 1;
        const firstCol = rec.dataView.getUint16(2, true) + 1;
        const lastCol = rec.dataView.getUint16(rec.data.length - 2, true) + 1;
        if (!matrix[rowIndex - 1]) matrix[rowIndex - 1] = [];
        let off = 4;
        for (let col = firstCol; col <= lastCol && off + 6 <= rec.data.length - 2; col++, off += 6) {
          const value = decodeRkValue(rec.dataView.getUint32(off + 2, true));
          matrix[rowIndex - 1][col - 1] = value;
          maxCol = Math.max(maxCol, col);
        }
        continue;
      }
      if (![0x0203, 0x0204, 0x00fd, 0x027e, 0x0006, 0x0002, 0x0001].includes(rec.op) || rec.data.length < 6) continue;
      const rowIndex = rec.dataView.getUint16(0, true) + 1;
      const colIndex0 = rec.dataView.getUint16(2, true) + 1;
      if (!matrix[rowIndex - 1]) matrix[rowIndex - 1] = [];
      currentRow = Math.max(currentRow, rowIndex - 1);
      let value = '';
      if (rec.op === 0x0203) value = rec.dataView.getFloat64(6, true);
      else if (rec.op === 0x0204) value = decodeBiffUnicodeString(rec.data, 6).text;
      else if (rec.op === 0x00fd) value = sst[rec.dataView.getUint32(6, true)] ?? '';
      else if (rec.op === 0x027e) value = decodeRkValue(rec.dataView.getUint32(6, true));
      else if (rec.op === 0x0006) value = decodeFormulaValue(rec);
      else if (rec.op === 0x0002) value = rec.dataView.getUint16(6, true);
      else if (rec.op === 0x0001) value = !!rec.dataView.getUint8(6);
      else continue;
      matrix[rowIndex - 1][colIndex0 - 1] = value;
      maxCol = Math.max(maxCol, colIndex0);
    }
    for (const ref of merged) {
      const value = matrix[ref.r1 - 1]?.[ref.c1 - 1];
      for (let r = ref.r1; r <= ref.r2; r++) {
        if (!matrix[r - 1]) matrix[r - 1] = [];
        for (let c = ref.c1; c <= ref.c2; c++) {
          if (matrix[r - 1][c - 1] === null || matrix[r - 1][c - 1] === undefined || matrix[r - 1][c - 1] === '') matrix[r - 1][c - 1] = value;
        }
      }
    }
    for (let r = 0; r < matrix.length; r++) {
      const row = matrix[r] || [];
      const obj = {};
      for (let c = 0; c < row.length; c++) {
        const value = row[c];
        if (value !== null && value !== undefined && value !== '') obj[colName(c + 1)] = value;
      }
      if (Object.keys(obj).length) rows.push(obj);
    }
    const text = [sheetName, sourceName, ...matrix.flat().map(normalizeCellText)].filter(Boolean).join('\n');
    return { name: sheetName, rows, matrix, text };
  }

  function decodeRkValue(raw) {
    const isInt = raw & 0x02;
    const multiply100 = raw & 0x01;
    let value;
    if (isInt) value = (raw >> 2) >> 0;
    else {
      const buf = new ArrayBuffer(8);
      const dv = new DataView(buf);
      dv.setUint32(4, raw & 0xfffffffc, true);
      value = dv.getFloat64(0, true);
    }
    return multiply100 ? value / 100 : value;
  }

  function decodeFormulaValue(rec) {
    const dv = rec.dataView;
    const resultType = dv.getUint8(6);
    if (resultType === 0x00) return dv.getFloat64(6, true);
    if (resultType === 0x01) return !!dv.getUint8(6);
    return '';
  }

  function parseLongTermRows(parsedRows) {
    const rows = (parsedRows || []).map((row) => {
      const ordered = Object.keys(row).sort(columnSort).map((k) => row[k]);
      while (ordered.length && (ordered[ordered.length - 1] === '' || ordered[ordered.length - 1] === null || ordered[ordered.length - 1] === undefined)) ordered.pop();
      return ordered;
    }).filter((row) => row.some((v) => v !== '' && v !== null && v !== undefined));
    if (!rows.length) return [];
    const layout = detectLongLayout(rows);
    const out = [];
    for (let i = layout.startIndex; i < rows.length; i++) {
      const row = rows[i];
      const month = normalizeMonthValue(row[layout.monthCol]);
      if (!month) continue;
      const values = [];
      for (let j = layout.valueStartCol; j < row.length && values.length < 24; j++) {
        const n = parseLongNumber(row[j], null);
        if (n !== null) values.push(round(n, 3));
      }
      if (values.length >= 24) out.push({ month, values: values.slice(0, 24) });
    }
    return out;
  }

  function detectLongLayout(rows) {
    const first = rows[0] || [];
    let monthCol = first.findIndex((v) => /月份|月度/.test(String(v ?? '').trim()));
    if (monthCol < 0) monthCol = first.findIndex((v) => normalizeMonthValue(v));
    if (monthCol < 0) monthCol = 0;
    let valueStartCol = monthCol + 1;
    for (let i = monthCol + 1; i < first.length; i++) {
      const s = String(first[i] ?? '').trim();
      if (/^时段\s*\d+$/i.test(s) || /^时点\s*\d+$/i.test(s) || /^\d{1,2}$/.test(s)) {
        valueStartCol = i;
        break;
      }
    }
    const firstText = String(first[monthCol] ?? '').trim();
    const hasMonthHeader = /月份|月份价格|月度/.test(firstText);
    const hasPointHeaders = first.slice(valueStartCol, valueStartCol + 24).some((v) => {
      const s = String(v ?? '').trim();
      return /^\d{1,2}$/.test(s) || /^0?\d{1,2}:?/.test(s) || /时段|时点/.test(s);
    });
    const startIndex = (hasMonthHeader || hasPointHeaders) ? 1 : 0;
    return { startIndex, monthCol, valueStartCol };
  }

  function normalizeMonthValue(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}`;
    const s = String(v).trim();
    if (/^\d{4}-\d{2}$/.test(s)) return s;
    const d = s.match(/(\d{4})[./-](\d{1,2})[./-]\d{1,2}/);
    if (d) return `${d[1]}-${String(Number(d[2])).padStart(2, '0')}`;
    const ym = s.match(/(\d{4})[./年-](\d{1,2})/);
    if (ym) return `${ym[1]}-${String(Number(ym[2])).padStart(2, '0')}`;
    if (s.length >= 7 && /^\d{4}-\d{1,2}/.test(s)) return `${s.slice(0, 4)}-${String(Number(s.slice(5, 7))).padStart(2, '0')}`;
    return '';
  }

  function parseLongNumber(v, fallback = 0) {
    if (v === null || v === undefined || v === '') return fallback;
    if (typeof v === 'number') return v;
    const n = Number(String(v).replace(/[^\d.-]/g, ''));
    return Number.isFinite(n) ? n : fallback;
  }

  async function recognizeImageText(file, kind = 'agent') {
    if (!window.Tesseract?.createWorker) throw new Error('OCR组件未加载成功');
    let lastProgressAt = 0;
    const worker = await getOcrWorker((m) => {
      const now = Date.now();
      if (now - lastProgressAt < 350) return;
      lastProgressAt = now;
      const pct = Number.isFinite(m?.progress) ? ` ${Math.round(m.progress * 100)}%` : '';
      setImportFeedback('正在识别图片数据，请稍候...', `${m?.status || ''}${pct}`.trim(), state.ui.importPreview || '正在读取图片内容...');
      scheduleRender();
    });
    try {
      const inputs = await preprocessImageForOcr(file, kind);
      const passes = kind === 'split' ? [6, 11] : [6, 4];
      const texts = [];
      let best = { text: '', score: -Infinity, meta: '' };
      for (let inputIdx = 0; inputIdx < inputs.length; inputIdx++) {
        const input = inputs[inputIdx];
        const inputText = typeof input === 'string' ? input : input?.dataUrl || '';
        const inputTag = typeof input === 'string' ? '' : (input?.tag || '');
        for (let passIdx = 0; passIdx < passes.length; passIdx++) {
          const psm = passes[passIdx];
          await worker.setParameters({
            tessedit_pageseg_mode: psm,
            tessedit_char_whitelist: '0123456789.-:/年月日 万千时元kwhKW MWHmwh ',
            preserve_interword_spaces: '1',
            user_defined_dpi: '300'
          });
          setImportFeedback('正在识别图片数据，请稍候...', `${inputTag || '图像'} · 轮次 ${inputIdx + 1}/${inputs.length} · psm ${psm}`, state.ui.importPreview || '正在读取图片内容...');
          scheduleRender();
          const { data } = await withTimeout(worker.recognize(inputText), 60000, '图片识别超时');
          const text = String(data?.text || '').trim();
          if (text) {
            texts.push(text);
            const score = scoreOcrText(text, kind);
            if (score > best.score) best = { text, score, meta: `crop=${inputTag || 'unknown'} psm=${psm} score=${score}` };
            if ((kind === 'agent' && score >= 70) || (kind === 'split' && score >= 55)) break;
          }
        }
        if ((kind === 'agent' && best.score >= 70) || (kind === 'split' && best.score >= 55)) break;
      }
      const text = texts.join('\n');
      if (!text.trim()) throw new Error('未识别到可用文字');
      if (best.text && best.text !== text) {
        setImportFeedback('正在整理识别结果', best.meta || '已找到更优识别结果', buildImportPreview({
          fileName: file.name,
          kind: kind === 'split' ? '市场分摊图片' : '代购价图片',
          rawText: best.text
        }));
        scheduleRender();
      }
      return text;
    } finally {
    }
  }

  async function getOcrWorker(logger) {
    if (!ocrWorkerPromise) {
      ocrWorkerPromise = withTimeout((async () => {
        if (!window.Tesseract?.createWorker) throw new Error('OCR组件未加载成功');
        return Tesseract.createWorker('eng', 1, {
          cacheMethod: 'none',
          workerBlobURL: true,
          logger
        });
      })(), 20000, 'OCR启动超时').catch((err) => {
        ocrWorkerPromise = null;
        throw err;
      });
    }
    return ocrWorkerPromise;
  }

  async function resetOcrWorker() {
    const worker = ocrWorkerPromise ? await ocrWorkerPromise.catch(() => null) : null;
    ocrWorkerPromise = null;
    try { if (worker?.terminate) await worker.terminate(); } catch (_) {}
  }

  function scoreOcrText(text, kind = 'agent') {
    const raw = String(text || '');
    const nums = raw.match(/-?\d+(?:\.\d+)?/g) || [];
    let score = nums.length * 4;
    score += (raw.match(/0\.\d{5,6}/g) || []).length * 6;
    score += (raw.match(/\d{4}\s*[年/-]\s*\d{1,2}/g) || []).length * 20;
    score += (raw.match(/\b1=\d\+\d\+\d\+\d\+\d\b/g) || []).length * 8;
    score += (raw.match(/\b4=5\+6\b/g) || []).length * 8;
    score += (raw.match(/\b\d{1,3},\d{3}\.\d{2}\b/g) || []).length * 2;
    if (kind === 'split') score += (raw.match(/0\.\d{6}/g) || []).length * 4;
    if (kind === 'agent') score += (raw.match(/0\.\d{6}/g) || []).length * 2;
    return score;
  }

  function assetUrl(path) {
    return new URL(path, window.location.href).href;
  }

  async function preprocessImageForOcr(file, kind = 'agent') {
    const url = URL.createObjectURL(file);
    try {
      const img = await loadImage(url);
      const w = img.naturalWidth || img.width;
      const h = img.naturalHeight || img.height;
      const crops = kind === 'split'
        ? [
            { tag: 'split-main', left: 0.12, width: 0.76, top: 0.43, height: 0.22 },
            { tag: 'split-info', left: 0.12, width: 0.76, top: 0.60, height: 0.16 }
          ]
        : [
            { tag: 'agent-main', left: 0.12, width: 0.76, top: 0.20, height: 0.28 },
            { tag: 'agent-info', left: 0.12, width: 0.76, top: 0.48, height: 0.18 }
          ];
      const outputs = [];
      for (const crop of crops) {
        outputs.push(renderOcrCrop(img, w, h, crop, kind === 'split' ? 195 : 200));
      }
      return outputs;
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  function renderOcrCrop(img, width, height, crop, threshold = 200) {
    const scale = width < 2200 ? 2.1 : 1.5;
    const cropLeft = Math.max(0, Math.min(width - 1, Math.round((crop.left || 0) * width)));
    const cropTop = Math.max(0, Math.min(height - 1, Math.round((crop.top || 0) * height)));
    const cropRight = Math.max(cropLeft + 1, Math.min(width, Math.round(cropLeft + (crop.width || 0.7) * width)));
    const cropBottom = Math.max(cropTop + 1, Math.min(height, Math.round(cropTop + (crop.height || 0.3) * height)));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round((cropRight - cropLeft) * scale));
    canvas.height = Math.max(1, Math.round((cropBottom - cropTop) * scale));
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.drawImage(img, cropLeft, cropTop, cropRight - cropLeft, cropBottom - cropTop, 0, 0, canvas.width, canvas.height);
    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const d = imageData.data;
    for (let i = 0; i < d.length; i += 4) {
      const gray = Math.round(d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114);
      const boosted = gray > threshold ? 255 : gray < 100 ? 0 : gray;
      d[i] = d[i + 1] = d[i + 2] = boosted;
    }
    ctx.putImageData(imageData, 0, 0);
    return { dataUrl: canvas.toDataURL('image/png'), tag: crop.tag || '' };
  }

  function loadImage(src) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = src;
    });
  }

  function parseAgentImportText(text, fallbackMonth) {
    const raw = String(text || '');
    const monthMatch = raw.match(/(\d{4})\s*[./年-]\s*(\d{1,2})/);
    const month = monthMatch ? `${monthMatch[1]}-${String(Number(monthMatch[2])).padStart(2, '0')}` : fallbackMonth;
    const body = raw.replace(/\d{4}[./-]\d{1,2}/g, ' ');
    const lines = body.split(/\n+/).map((s) => s.trim()).filter(Boolean);
    const nums = lines.flatMap((line) => (line.match(/-?\d*\.\d+/g) || []).map(Number));
    const flatPriceKwh = pickAgentFlat(nums);
    const averagePurchaseKwh = pickAveragePurchase(nums, flatPriceKwh);
    const historyDeviationKwh = pickHistoryDeviation(nums, flatPriceKwh, averagePurchaseKwh);
    if (!flatPriceKwh || !averagePurchaseKwh) throw new Error('没有识别到代理购电价格或当月平均购电价格');
    const labels = state.data.agentPurchase?.voltageLevels || [];
    const splitParts = inferSplitParts(nums);
    const voltageLevels = labels.map((base, idx) => {
      const trans = pickTransmissionForLevel(nums, base.transmissionKwh);
      return {
        ...clone(base),
        lineLossKwh: splitParts.lineLoss ?? base.lineLossKwh ?? 0,
        transmissionKwh: trans ?? base.transmissionKwh ?? 0,
        fundKwh: splitParts.fund ?? base.fundKwh ?? 0,
        systemKwh: splitParts.system ?? base.systemKwh ?? 0
      };
    });
    return { month, flatPriceKwh, averagePurchaseKwh, historyDeviationKwh, voltageLevels };
  }

  function parseSplitImportText(text, fallbackMonth) {
    const raw = String(text || '');
    const monthMatch = raw.match(/(\d{4})\s*[./年-]\s*(\d{1,2})/);
    const month = monthMatch ? `${monthMatch[1]}-${String(Number(monthMatch[2])).padStart(2, '0')}` : fallbackMonth;
    const baseLevels = normalizeSplitVoltageLevels(state.data.marketSplit?.voltageLevels || state.data.agentPurchase?.voltageLevels || []);
    const fallbackNums = raw.replace(/\d{4}[./-]\d{1,2}/g, ' ').match(/-?\d*\.\d+/g)?.map(Number) || [];
    const splitParts = inferSplitParts(fallbackNums);
    if (splitParts.lineLoss === null && splitParts.fund === null && splitParts.system === null && !fallbackNums.some((n) => n > 0.02 && n < 0.16)) {
      throw new Error('没有识别到市场分摊价格');
    }
    const voltageLevels = baseLevels.map((base, idx) => {
      const trans = pickTransmissionForLevel(fallbackNums, base.transmissionKwh);
      return {
        ...clone(base),
        lineLossKwh: splitParts.lineLoss ?? base.lineLossKwh ?? 0,
        transmissionKwh: trans ?? base.transmissionKwh ?? 0,
        fundKwh: splitParts.fund ?? base.fundKwh ?? 0,
        systemKwh: splitParts.system ?? base.systemKwh ?? 0
      };
    });
    return { month, voltageLevels: normalizeSplitVoltageLevels(voltageLevels) };
  }

  function pickAgentFlat(nums) {
    const candidates = nums.filter((n) => n >= 0.22 && n <= 0.5);
    if (!candidates.length) return 0;
    return round(Math.min(...candidates), 6);
  }

  function pickAveragePurchase(nums, flat) {
    const candidates = nums.filter((n) => n >= 0.22 && n <= 0.5 && Math.abs(n - flat) <= 0.08 && Math.abs(n - flat) > 0.000001);
    if (!candidates.length) return flat;
    return round(candidates.reduce((best, n) => Math.abs(n - flat) < Math.abs(best - flat) ? n : best, candidates[0]), 6);
  }

  function pickHistoryDeviation(nums, flat, avg) {
    const diff = round(flat - avg, 6);
    const candidates = nums.filter((n) => Math.abs(n) < 0.08 && Math.abs(n - flat) > 0.1 && Math.abs(n - avg) > 0.1);
    const exact = candidates.find((n) => Math.abs(n - diff) < 0.002);
    return round(exact ?? diff, 6);
  }

  function inferSplitParts(nums) {
    const lineLoss = pickFrequentDecimal(nums, (n) => n >= 0.008 && n <= 0.025);
    const fund = pickClosestDecimal(nums, 0.04336875, (n) => n >= 0.035 && n <= 0.055);
    const system = pickClosestDecimal(nums, 0.097, (n) => n >= 0.08 && n <= 0.12);
    return { lineLoss, fund, system };
  }

  function pickTransmissionForLevel(nums, expected) {
    return pickClosestDecimal(nums, Number(expected || 0), (n) => n > 0.02 && n < 0.16);
  }

  function pickFrequentDecimal(nums, predicate) {
    const counts = new Map();
    for (const n of nums.filter(predicate)) {
      const key = round(n, 6);
      counts.set(key, (counts.get(key) || 0) + 1);
    }
    let best = null;
    for (const [value, count] of counts) {
      if (!best || count > best.count) best = { value, count };
    }
    return best ? best.value : null;
  }

  function pickClosestDecimal(nums, target, predicate) {
    const candidates = nums.filter(predicate);
    if (!candidates.length) return null;
    return round(candidates.reduce((best, n) => Math.abs(n - target) < Math.abs(best - target) ? n : best, candidates[0]), 6);
  }

  async function readSharedStrings(zip) {
    const file = zip.file('xl/sharedStrings.xml');
    if (!file) return [];
    const xml = await file.async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    return [...doc.querySelectorAll('si')].map((si) => si.textContent || '');
  }

  async function readStyles(zip) {
    const dateStyles = new Set();
    const file = zip.file('xl/styles.xml');
    if (!file) return { dateStyles };
    const xml = await file.async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    [...doc.querySelectorAll('cellXfs xf')].forEach((xf, idx) => {
      const numFmtId = Number(xf.getAttribute('numFmtId') || 0);
      if ([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58].includes(numFmtId)) dateStyles.add(idx);
    });
    return { dateStyles };
  }

  function excelSerialToDate(serial) {
    return new Date(Math.round((serial - 25569) * 86400 * 1000)).toISOString().slice(0, 10);
  }

  function normalizeDate(v) {
    if (!v) return '';
    if (typeof v === 'string') return v.slice(0, 10);
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).slice(0, 10);
  }

  function columnSort(a, b) { return colIndex(a) - colIndex(b); }
  function colIndex(c) { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
  function colName(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function parseNumList(text, fixed) {
    const nums = String(text || '').split(/[, \n\r\t]+/).map((x) => x.trim()).filter(Boolean).map(Number).filter((n) => !Number.isNaN(n));
    if (fixed) { while (nums.length < fixed) nums.push(0); return nums.slice(0, fixed); }
    return nums;
  }
  function excelDateKey(v) { return String(v || '').slice(0, 10); }

  function loadPointLabels96() {
    return defaultLoadHeaders().slice(7);
  }

  function loadLabels() {
    if (state.ui.loadMode === '24') return baseChartUnits();
    return loadPointLabels96();
  }

  function drawLineChart(id, series, unit = '', xLabels = null, opts = {}) {
    const canvas = byId(id);
    if (!canvas) return;
    const parent = canvas.parentElement;
    const w = parent.clientWidth - 24;
    const legendRows = opts.legend === false ? 0 : estimateLegendRows(series, Math.max(120, w - 80));
    const legendHeight = opts.legend === false ? 0 : Math.max(24, legendRows * 22);
    const h = 336 + legendHeight + 18;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.max(10, w * dpr);
    canvas.height = h * dpr;
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, w, h);
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, w, h);
    if (!series.length || !series.some((s) => s.values.length)) return;
    const pad = { l: 58, r: 22, t: 24, b: 74 + legendHeight + 22 };
    const plotW = w - pad.l - pad.r;
    const plotH = h - pad.t - pad.b;
    const all = series.flatMap((s) => s.values || []).filter(isFiniteDataValue).map(Number);
    if (!all.length) return;
    const min = Math.min(...all);
    const maxRaw = Number.isFinite(series.fixedMax) ? series.fixedMax : Math.max(...all);
    const yMax = Number.isFinite(series.fixedMax) ? (Math.ceil(maxRaw / 10) * 10 || 10) : niceAxisMax(maxRaw);
    const yMin = series.autoYMin ? niceAxisMin(min, yMax) : 0;
    const ySpan = yMax - yMin || 1;
    const ticks = 5;
    ctx.strokeStyle = '#e6efed';
    ctx.fillStyle = '#49615d';
    ctx.font = '12px system-ui';
    for (let i = 0; i <= ticks; i++) {
      const y = yMin + (ySpan * i) / ticks;
      const py = pad.t + plotH - (i / ticks) * plotH;
      ctx.beginPath(); ctx.moveTo(pad.l, py); ctx.lineTo(w - pad.r, py); ctx.stroke();
      ctx.fillText(formatAxisValue(y), 6, py + 4);
    }
    const pointCount = Math.max(...series.map((s) => s.values.length));
    const labels = xLabels || Array.from({ length: pointCount }, (_, i) => String(i));
    labels.forEach((lab, i) => {
      if (pointCount > 24 && i % 4 !== 0) return;
      const shown = pointCount > 24 ? String(Math.floor(i / 4)) : lab;
      const x = pad.l + (plotW * i) / Math.max(1, pointCount - 1);
      ctx.fillText(shown, x - 3, h - legendHeight - 28);
    });
    const labelBoxes = [];
    series.forEach((s, idx) => {
      const color = s.color || seriesColor(idx);
      const validPoints = (s.values || [])
        .map((v, i) => ({ value: Number(v), index: i }))
        .filter((p) => isFiniteDataValue(s.values[p.index]));
      if (!validPoints.length) return;
      ctx.strokeStyle = color;
      ctx.lineWidth = 2;
      ctx.beginPath();
      let hasSegment = false;
      s.values.forEach((v, i) => {
        if (!isFiniteDataValue(v)) {
          hasSegment = false;
          return;
        }
        const x = pad.l + (plotW * i) / Math.max(1, s.values.length - 1);
        const y = pad.t + plotH - ((Number(v) - yMin) / ySpan) * plotH;
        if (!hasSegment) {
          ctx.moveTo(x, y);
          hasSegment = true;
        } else {
          ctx.lineTo(x, y);
        }
      });
      ctx.stroke();
      s.values.forEach((v, i) => {
        if (!isFiniteDataValue(v)) return;
        const x = pad.l + (plotW * i) / Math.max(1, s.values.length - 1);
        const y = pad.t + plotH - ((Number(v) - yMin) / ySpan) * plotH;
        ctx.fillStyle = color;
        ctx.beginPath(); ctx.arc(x, y, 2.4, 0, Math.PI * 2); ctx.fill();
      });
      if (s.showExtrema !== false) {
        const hi = validPoints.reduce((best, p) => (p.value > best.value ? p : best), validPoints[0]).index;
        const lo = validPoints.reduce((best, p) => (p.value < best.value ? p : best), validPoints[0]).index;
        [hi, lo].forEach((i, order) => {
          const x = pad.l + (plotW * i) / Math.max(1, s.values.length - 1);
          const y = pad.t + plotH - ((Number(s.values[i]) - yMin) / ySpan) * plotH;
          const label = formatNumber(s.values[i]);
          const labelW = ctx.measureText(label).width;
          const boxW = labelW + 8;
          const boxH = 18;
          const candidates = [
            { lx: x + 10, ly: y - 18 },
            { lx: x + 10, ly: y + 22 },
            { lx: x - labelW - 12, ly: y - 18 },
            { lx: x - labelW - 12, ly: y + 22 },
            { lx: x - labelW / 2, ly: y - 30 },
            { lx: x - labelW / 2, ly: y + 34 }
          ];
          const picked = pickLabelPosition(candidates, labelBoxes, pad, w, h, boxW, boxH, x, y);
          const lx = picked.lx;
          const ly = picked.ly;
          labelBoxes.push({ x: lx - 4, y: ly - 14, w: boxW, h: boxH });
          ctx.fillStyle = order === 0 ? '#dc2626' : '#16a34a';
          ctx.beginPath(); ctx.arc(x, y, 4.5, 0, Math.PI * 2); ctx.fill();
          ctx.fillStyle = 'rgba(255,255,255,0.86)';
          ctx.fillRect(lx - 4, ly - 14, boxW, boxH);
          ctx.fillStyle = '#1f2f2d';
          ctx.fillText(label, lx, ly);
        });
      }
    });
    if (opts.legend !== false) drawLegend(ctx, series, pad.l, h - legendHeight + 4, w - pad.l - pad.r);
    return { min, max: maxRaw };
  }

  function pickLabelPosition(candidates, existing, pad, width, height, boxW, boxH, pointX, pointY) {
    const clamp = (candidate) => ({
      lx: Math.max(pad.l + 2, Math.min(width - pad.r - boxW, candidate.lx)),
      ly: Math.max(pad.t + boxH, Math.min(height - pad.b - 2, candidate.ly))
    });
    const overlaps = (box, other) => !(box.x + box.w < other.x || other.x + other.w < box.x || box.y + box.h < other.y || other.y + other.h < box.y);
    let best = null;
    for (const c of candidates.map(clamp)) {
      const box = { x: c.lx - 4, y: c.ly - 14, w: boxW, h: boxH };
      const pointPenalty = pointX >= box.x - 4 && pointX <= box.x + box.w + 4 && pointY >= box.y - 6 && pointY <= box.y + box.h + 6 ? 1000 : 0;
      const overlapPenalty = existing.reduce((sum, e) => sum + (overlaps(box, e) ? 500 : 0), 0);
      const distancePenalty = Math.abs(c.lx - pointX) + Math.abs(c.ly - pointY);
      const score = pointPenalty + overlapPenalty + distancePenalty;
      if (!best || score < best.score) best = { ...c, score };
      if (score < 500) break;
    }
    return best || clamp(candidates[0] || { lx: pointX + 10, ly: pointY - 18 });
  }

  function seriesColor(key) {
    const palette = ['#0f766e', '#ca8a04', '#2563eb', '#7c3aed', '#dc2626', '#14b8a6', '#f97316', '#4f46e5', '#db2777', '#16a34a'];
    if (typeof key === 'string') {
      let hash = 0;
      for (let i = 0; i < key.length; i++) hash = ((hash << 5) - hash + key.charCodeAt(i)) | 0;
      return palette[Math.abs(hash) % palette.length];
    }
    return palette[Math.abs(Number(key || 0)) % palette.length];
  }

  function estimateLegendRows(series, maxWidth) {
    let x = 0;
    let rows = 1;
    const approxChar = 7;
    for (const s of series) {
      const itemW = Math.min(260, 26 + String(s.label || '').length * approxChar);
      if (x > 0 && x + itemW > maxWidth) {
        rows += 1;
        x = 0;
      }
      x += itemW + 22;
    }
    return rows;
  }

  function drawLegend(ctx, series, startX, startY, maxWidth) {
    let x = startX;
    let y = startY;
    series.forEach((s, idx) => {
      const color = s.color || seriesColor(idx);
      const label = String(s.label || '');
      const labelW = ctx.measureText(label).width;
      const itemW = Math.min(maxWidth, labelW + 26);
      if (x > startX && x + itemW > startX + maxWidth) {
        x = startX;
        y += 22;
      }
      ctx.fillStyle = color;
      ctx.fillRect(x, y - 9, 10, 10);
      ctx.fillStyle = '#47615d';
      ctx.fillText(label, x + 16, y);
      x += itemW + 22;
    });
  }

  function toggleLegend(group, value) {
    const key = `${group}LegendHidden`;
    const hidden = new Set(state.ui[key] || []);
    if (hidden.has(value)) hidden.delete(value);
    else hidden.add(value);
    state.ui[key] = [...hidden];
    persist();
    scheduleRender();
  }

  function resampleSeries(points, targetCount) {
    const src = (Array.isArray(points) ? points : []).map((v) => Number(v || 0));
    if (!targetCount) return src;
    if (!src.length) return Array(targetCount).fill(0);
    if (src.length === targetCount) return src.map((v) => round(v, 3));
    if (src.length < targetCount && targetCount % src.length === 0) {
      const factor = targetCount / src.length;
      const out = [];
      for (const v of src) {
        const each = round(Number(v || 0) / factor, 3);
        for (let i = 0; i < factor; i++) out.push(each);
      }
      return out.slice(0, targetCount);
    }
    if (src.length > targetCount && src.length % targetCount === 0) {
      const factor = src.length / targetCount;
      const out = [];
      for (let i = 0; i < targetCount; i++) {
        const slice = src.slice(i * factor, i * factor + factor);
        out.push(round(slice.reduce((a, b) => a + Number(b || 0), 0), 3));
      }
      return out;
    }
    if (targetCount === 24) {
      if (src.length === 96) return to24(src);
      if (src.length === 48) return resampleSeries(src, 24);
      if (src.length === 288) return resampleSeries(src, 24);
    }
    if (targetCount === 96) {
      if (src.length === 24) return to96From24(src);
      if (src.length === 48) return resampleSeries(src, 96);
      if (src.length === 288) return resampleSeries(src, 96);
    }
    const ratio = src.length / targetCount;
    const out = [];
    for (let i = 0; i < targetCount; i++) {
      const start = Math.floor(i * ratio);
      const end = Math.max(start + 1, Math.floor((i + 1) * ratio));
      const slice = src.slice(start, end);
      out.push(round(slice.reduce((a, b) => a + Number(b || 0), 0), 3));
    }
    return out;
  }

  function niceAxisMax(max) {
    if (!Number.isFinite(max) || max <= 0) return 1;
    if (max <= 1) return Math.ceil(max * 10) / 10;
    if (max <= 10) return Math.ceil(max);
    return Math.ceil(max / 10) * 10;
  }

  function niceAxisMin(min, max) {
    if (!Number.isFinite(min) || min <= 0) return 0;
    const span = Math.max(max - min, max * 0.08, 1);
    const raw = min - span * 0.18;
    if (max <= 1) return Math.max(0, Math.floor(raw * 100) / 100);
    if (max <= 10) return Math.max(0, Math.floor(raw * 10) / 10);
    return Math.max(0, Math.floor(raw / 10) * 10);
  }

  function formatAxisValue(v) {
    if (Math.abs(v) < 1) return v.toFixed(2);
    if (Math.abs(v) < 10) return v.toFixed(1);
    return String(Math.round(v));
  }

  function maxIndex(values) {
    let idx = 0;
    for (let i = 1; i < values.length; i++) if (Number(values[i]) > Number(values[idx])) idx = i;
    return idx;
  }
  function minIndex(values) {
    let idx = 0;
    for (let i = 1; i < values.length; i++) if (Number(values[i]) < Number(values[idx])) idx = i;
    return idx;
  }

  function aggregateLoad(records) {
    if (cache.loadRecordsRef === records && cache.loadRecordsLength === records.length) return cache.loadAggregated;
    cache.loadRecordsRef = records;
    cache.loadRecordsLength = records.length;
    cache.loadAggregated = sortLoadRecordsByUserDate(combineLoadRecordsByUserDate(records));
    return cache.loadAggregated;
  }

  function formatNumber(v) { return fmt(v, 3); }
  function longSeriesForMonth(m, voltageId = selectedVoltageId(), k = state.ui.longK, includeSplit = state.ui.longSplit) {
    const split = splitTotalForPrice(m?.month, voltageId, includeSplit);
    return (m?.values || []).map((v) => (split === null || !isFiniteDataValue(v) ? null : round(Number(v) + split + Number(k || 0), 3)));
  }

  function longChartSeries() {
    return selectedMonths().map((m) => ({
      label: selectedMonths().length > 1 ? `${longModeLabel(state.ui.longSplit, state.ui.longK)} ${m.month}` : longModeLabel(state.ui.longSplit, state.ui.longK),
      values: longSeriesForMonth(m, selectedVoltageId())
    }));
  }

  function agentSeries(monthObj, voltageId = selectedVoltageId(), includeSplit = state.ui.agentSplit) {
    const agent = monthObj || {};
    const split = splitTotalForPrice(agent.month, voltageId, includeSplit);
    const tou = agentTouPricesKwh(agent);
    return Array.from({ length: 24 }, (_, h) => {
      const type = hourType(agent.month || latestAgentMonth(), h);
      const price = tou[type] ?? tou.flat ?? 0;
      return split === null || !isFiniteDataValue(price) ? null : round(yuanKwhToMwh(price) + split, 3);
    });
  }

  async function exportTable(kind, sourceBtn) {
    let rows = [];
    let name = `${kind}.xlsx`;
    if (kind === 'long') {
      rows = [['日期', ...Array.from({ length: 24 }, (_, i) => String(i))], ...selectedMonths().map((m) => [m.month, ...longSeriesForMonth(m, selectedVoltageId())])];
      name = '中长期分时段价格.xlsx';
    } else if (kind === 'agent') {
      const voltageId = selectedVoltageId();
      const chosen = selectedAgentMonths();
      rows = [['日期', ...Array.from({ length: 24 }, (_, i) => String(i))], ...chosen.map((month) => [month.month, ...agentSeries(month, voltageId, state.ui.agentSplit)])];
      name = '国网代购价.xlsx';
    } else if (kind === 'split') {
      rows = [['月份', '上网环节线损折价(元/MWh)', '电量输配电价(元/MWh)', '政府性基金及附加(元/MWh)', '系统运行费折价(元/MWh)', '合计(元/MWh)'], ...getSplitMonthRows(selectedSplitVoltageId()).map((r) => [r.month, r.lineLoss ?? '/', r.transmission ?? '/', r.fund ?? '/', r.system ?? '/', r.total ?? '/'])];
      name = '市场分摊.xlsx';
    } else if (kind === 'load') {
      const agg = aggregateLoad(state.data.userLoad?.records || []);
      const currentUser = (state.ui.loadUsers || [])[0] || '';
      const filtered = agg.filter((r) => {
        const userHit = currentUser ? r.userName === currentUser : false;
        const after = !state.ui.loadDateFrom || r.date >= state.ui.loadDateFrom;
        const before = !state.ui.loadDateTo || r.date <= state.ui.loadDateTo;
        return userHit && after && before;
      });
      const labels = state.ui.loadMode === '24' ? Array.from({ length: 24 }, (_, i) => String(i)) : loadLabels();
      rows = [['用户', '日期', '日总量', ...labels], ...filtered.map((r) => [r.userName, r.date, r.dailyTotal, ...(state.ui.loadMode === '24' ? to24(r.points) : r.points)])];
      name = '用户负荷数据.xlsx';
    } else if (kind === 'comparePrice') {
      const month = state.ui.compareMonth || latestLongMonth();
      const voltageId = selectedVoltageId();
      const longMonth = pickPriceMonthByDate(month, state.data.longTermWeightedClearing?.months || [], latestLongMonth());
      const agentMonth = pickPriceMonthByDate(month, state.data.agentPurchase?.months || [], latestAgentMonth()) || getAgentMonth(month);
      const longVals = longMonth ? longSeriesForMonth(longMonth, voltageId, state.ui.compareK, state.ui.compareSplit) : [];
      const agentVals = agentSeries(agentMonth, voltageId, state.ui.compareSplit);
      rows = [['时点', '中长期', '国网代购价', '差额'], ...longVals.map((v, i) => {
        const a = agentVals[i] ?? null;
        const diff = isFiniteDataValue(v) && isFiniteDataValue(a) ? Number(v) - Number(a) : '';
        return [i, v ?? '', a ?? '', diff];
      })];
      name = '价格对比.xlsx';
    } else if (kind === 'compareLoad') {
      const temp = compareRowsForMonthlyLoad(state.ui.compareRows || []);
      rows = [['用户', '月份', ...Array.from({ length: 24 }, (_, i) => String(i)), '售电公司加权均价', '国网代购加权均价'], ...temp.map((r) => [r.userName, r.month, ...r.values, r.longAvg, r.agentAvg])];
      name = '用户测算结果.xlsx';
    }
    if (!rows.length) {
      alert('当前没有可导出的表格数据');
      return;
    }
    await exportWorkbook([{ name: '表格数据', rows }], name, sourceBtn);
  }

  function handleCompareImport(e) {
    const files = [...(e.target.files || [])];
    if (!files.length) return;
    Promise.all(files.map((file) => file.arrayBuffer().then((buf) => parseLoadWorkbook(buf, file.name)).catch((err) => ({ error: err, file: file.name })))).then((results) => {
      const merged = [];
      const failures = [];
      for (const result of results) {
        if (result?.error) failures.push(`${result.file}：${result.error?.message || result.error}`);
        else merged.push(...(result.records || []));
      }
      if (!merged.length) throw new Error(failures[0] || '没有识别到可用负荷数据');
      state.ui.compareRows = sortLoadRecordsByUserDate(aggregateLoad(merged));
      persist();
      e.target.value = '';
      if (failures.length) alert(`部分文件导入失败：\n${failures.join('\n')}`);
      scheduleRender();
    }).catch((err) => alert('导入失败：' + err.message));
  }

  async function exportData() {
    const loadRows = managedLoadRows();
    const loadMode = state.ui.loadManageMode === '24' ? '24' : '96';
    const loadLabelsOut = loadMode === '24' ? Array.from({ length: 24 }, (_, i) => String(i)) : loadLabels();
    const sheets = [
      { name: 'long_term', rows: [['月份', ...Array.from({ length: 24 }, (_, i) => String(i))], ...(state.data.longTermWeightedClearing?.months || []).map((m) => [m.month, ...(m.values || [])])] },
      { name: 'agent_price', rows: [['月份', '平段价', '当月平均', '历史偏差'], ...(state.data.agentPurchase?.months || []).map((m) => [m.month, m.flatPriceKwh ?? '', m.averagePurchaseKwh ?? '', m.historyDeviationKwh ?? ''])] },
      { name: 'market_split', rows: [['月份', '电压等级', '线损', '输配', '基金', '系统'], ...(state.data.marketSplit?.months || []).flatMap((m) => (m.voltageLevels || []).map((v) => [m.month, v.label, v.lineLossKwh ?? '', v.transmissionKwh ?? '', v.fundKwh ?? '', v.systemKwh ?? '']))] },
      { name: 'user_load', rows: [['用户', '日期', '日总量', ...loadLabelsOut], ...loadRows.map((r) => [r.userName, r.date, r.dailyTotal, ...(loadMode === '24' ? to24(r.points) : (r.points || []).slice(0, 96))])] },
      { name: 'user_load_raw', rows: [['用户', '日期', '日总量', ...Array.from({ length: 96 }, (_, i) => String(i))], ...(state.data.userLoad?.records || []).map((r) => [r.userName || '', r.date || '', r.dailyTotal ?? '', ...(r.points || []).slice(0, 96)])] }
    ];
    await exportWorkbook(sheets, 'shanxi-electric-dashboard-data.xlsx', byId('exportDataBtn'));
  }

  async function exportWorkbook(sheets, fileName, sourceBtn) {
    if (typeof JSZip === 'undefined') throw new Error('Excel组件没有加载成功，请刷新页面后重试');
    const zip = new JSZip();
    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${xmlEscape(String(s.name).slice(0, 31))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`;
    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}</Relationships>`;
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`);
    zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.folder('xl').file('workbook.xml', workbookXml);
    zip.folder('xl').folder('_rels').file('workbook.xml.rels', relsXml);
    zip.folder('xl').file('styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`);
    sheets.forEach((sheet, i) => {
      zip.folder('xl').folder('worksheets').file(`sheet${i + 1}.xml`, sheetXml(sheet.rows));
    });
    const buf = await zip.generateAsync({ type: 'blob' });
    downloadBlob(buf, fileName, sourceBtn);
  }

  function sheetXml(rows) {
    const body = rows.map((row, r) => `<row r="${r + 1}">${row.map((cell, c) => {
      const ref = `${colName(c + 1)}${r + 1}`;
      if (typeof cell === 'number' && Number.isFinite(cell)) return `<c r="${ref}"><v>${round(cell, 3)}</v></c>`;
      return `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(cell)}</t></is></c>`;
    }).join('')}</row>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
  }

  function xmlEscape(v) {
    return String(v ?? '').replace(/[<>&"]/g, (m) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' }[m]));
  }

  function downloadBlob(blob, name, sourceBtn) {
    if (lastExportUrl) URL.revokeObjectURL(lastExportUrl);
    const url = URL.createObjectURL(blob);
    lastExportUrl = url;
    document.querySelectorAll('.download-ready').forEach((x) => x.remove());
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

})();
