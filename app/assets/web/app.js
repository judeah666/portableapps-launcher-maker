const appRoot = document.getElementById('app');
const modalRoot = document.getElementById('modal-root');

const state = {
  schema: [],
  previewTabs: [],
  activeTab: 'appinfo',
  activePreviewTab: 'folder',
  snapshot: null,
  draft: {},
  draftTimers: new Map(),
  booted: false,
  bootInFlight: false,
  scrollTopByTab: {},
  sidebarScrollTopByTab: {},
  previewScrollTop: 0,
};

function renderLoading(message = 'Loading PortableApps Launcher Maker...') {
  appRoot.innerHTML = `
    <div class="app-shell">
      <section class="panel">
        <div class="panel-header">
          <h2>PortableApps.com Launcher Maker</h2>
          <p class="panel-note">${escapeHtml(message)}</p>
        </div>
      </section>
    </div>
  `;
}

function renderFatal(message) {
  appRoot.innerHTML = `
    <div class="app-shell">
      <section class="panel">
        <div class="panel-header">
          <h2>PortableApps.com Launcher Maker</h2>
          <p class="panel-note">The web UI could not start.</p>
        </div>
        <div class="panel-body preview-mode">
          <div class="preview-only">
            <div class="side-card">
              <div class="section-head"><h3>Startup Error</h3></div>
              <div class="section-body">
                <pre class="preview-text">${escapeHtml(message)}</pre>
              </div>
            </div>
          </div>
        </div>
      </section>
    </div>
  `;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

function debounceSave(key, value, delay = 180) {
  if (state.draftTimers.has(key)) {
    clearTimeout(state.draftTimers.get(key));
  }
  const timer = setTimeout(async () => {
    const snapshot = await window.pywebview.api.set_value(key, value);
    applySnapshot(snapshot);
    render();
  }, delay);
  state.draftTimers.set(key, timer);
}

function applySnapshot(snapshot) {
  state.snapshot = snapshot;
  state.draft = { ...snapshot.state };
}

function captureScrollState() {
  const formScroll = document.querySelector('.form-scroll');
  if (formScroll) {
    state.scrollTopByTab[state.activeTab] = formScroll.scrollTop;
  }
  const sidebarScroll = document.querySelector('.sidebar-scroll');
  if (sidebarScroll) {
    state.sidebarScrollTopByTab[state.activeTab] = sidebarScroll.scrollTop;
  }
  const previewScroll = document.querySelector('.preview-scroll');
  if (previewScroll) {
    state.previewScrollTop = previewScroll.scrollTop;
  }
}

function restoreScrollState() {
  const formScroll = document.querySelector('.form-scroll');
  if (formScroll) {
    formScroll.scrollTop = state.scrollTopByTab[state.activeTab] || 0;
  }
  const sidebarScroll = document.querySelector('.sidebar-scroll');
  if (sidebarScroll) {
    sidebarScroll.scrollTop = state.sidebarScrollTopByTab[state.activeTab] || 0;
  }
  const previewScroll = document.querySelector('.preview-scroll');
  if (previewScroll) {
    previewScroll.scrollTop = state.previewScrollTop || 0;
  }
}

function isRegistryDisabled(key) {
  return ['registry_keys', 'registry_cleanup_if_empty', 'registry_cleanup_force'].includes(key) && !state.draft.registry_enabled;
}

function fieldActionButton(field) {
  if (field.key === 'icon_source') {
    return `<button class="ghost-btn" type="button" data-action="choose-icon">Browse</button>`;
  }
  return '';
}

function renderField(field) {
  const value = state.draft[field.key] ?? '';
  const wideClass = field.wide || field.type === 'textarea' ? 'full' : '';
  if (field.type === 'checkbox') {
    return `
      <div class="field checkbox ${wideClass}">
        <input class="checkbox-input" id="${field.key}" data-key="${field.key}" type="checkbox" ${value ? 'checked' : ''}>
        <label for="${field.key}">${escapeHtml(field.label)}</label>
      </div>
    `;
  }

  if (field.type === 'readonly') {
    return `
      <div class="field ${wideClass}">
        <label>${escapeHtml(field.label)}</label>
        <div class="readonly-field">${escapeHtml(value)}</div>
      </div>
    `;
  }

  if (field.type === 'select') {
    return `
      <div class="field ${wideClass}">
        <label for="${field.key}">${escapeHtml(field.label)}</label>
        <div class="input-row">
          <select class="select-input" id="${field.key}" data-key="${field.key}" ${isRegistryDisabled(field.key) ? 'disabled' : ''}>
            ${field.options.map(option => `<option value="${escapeHtml(option)}" ${option === value ? 'selected' : ''}>${escapeHtml(option)}</option>`).join('')}
          </select>
          ${fieldActionButton(field)}
        </div>
      </div>
    `;
  }

  if (field.type === 'textarea') {
    return `
      <div class="field full">
        <label for="${field.key}">${escapeHtml(field.label)}</label>
        <textarea class="text-area" id="${field.key}" data-key="${field.key}" rows="${field.rows || 5}" ${isRegistryDisabled(field.key) ? 'disabled' : ''}>${escapeHtml(value)}</textarea>
      </div>
    `;
  }

  return `
    <div class="field ${wideClass}">
      <label for="${field.key}">${escapeHtml(field.label)}</label>
      <div class="input-row">
        <input class="text-input" id="${field.key}" data-key="${field.key}" type="text" value="${escapeHtml(value)}" ${isRegistryDisabled(field.key) ? 'disabled' : ''}>
        ${fieldActionButton(field)}
      </div>
    </div>
  `;
}

function renderSections(tab) {
  const current = state.schema.find(item => item.key === tab);
  if (!current) {
    return '';
  }
  return current.sections.map(section => `
    <details class="section-card" open>
      <summary class="section-head">
        <div>
          <h3>${escapeHtml(section.title)}</h3>
        </div>
      </summary>
      <div class="section-body">
        ${
          tab === 'registry' && section.title === 'Registry'
            ? '<div style="margin-bottom:12px; display:flex; gap:10px;"><button class="ghost-btn" type="button" data-action="import-registry" ' + (!state.draft.registry_enabled ? 'disabled' : '') + '>Import Saved Registry (.reg)</button></div>'
            : tab === 'splash' && section.title === 'Splash Asset'
              ? '<div style="margin-bottom:12px; display:flex; gap:10px; flex-wrap:wrap;"><button class="ghost-btn" type="button" data-action="open-assets-folder">Open Assets Folder</button><button class="ghost-btn" type="button" data-action="open-splash-asset">Open Splash</button><button class="ghost-btn" type="button" data-action="replace-splash-asset">Replace Splash</button></div>'
            : ''
        }
        <div class="field-grid">
          ${section.fields.map(renderField).join('')}
        </div>
      </div>
    </details>
  `).join('');
}

function renderPreviewSidebar() {
  const previews = state.snapshot.previews || {};
  const activePreview = state.activePreviewTab;
  const iconPayload = state.snapshot.iconPreviews || { items: [], message: '' };
  const iconPreviews = iconPayload.items || [];
  const iconMessage = iconPayload.message || '';
  const splash = state.snapshot.splashPreview;
  return `
    <div class="sidebar-scroll">
      <div class="side-card">
        <div class="section-head">
          <h3>Preview</h3>
        </div>
        <div class="section-body">
          <div class="preview-tabs">
            ${state.previewTabs.map(tab => `
              <button class="preview-tab-btn ${tab.key === activePreview ? 'active' : ''}" data-preview-tab="${tab.key}">${escapeHtml(tab.label)}</button>
            `).join('')}
          </div>
          <div class="preview-frame" style="margin-top:12px;">
            <pre class="preview-text">${escapeHtml(previews[activePreview] || '')}</pre>
          </div>
        </div>
      </div>

      <div class="side-card">
        <div class="section-head"><h3>Icon Preview</h3></div>
        <div class="section-body">
          <div class="icon-strip">
            ${iconPreviews.map(icon => `
              <div class="icon-tile" style="width:${icon.width}px;">
                <div class="icon-box" style="width:${icon.width}px;height:${icon.height}px;"><img src="${icon.src}" alt="${escapeHtml(icon.label)}"></div>
                <div class="icon-label">${escapeHtml(icon.label)}</div>
              </div>
            `).join('')}
          </div>
          <p class="section-note" style="margin-top:12px;">${escapeHtml(iconMessage)}</p>
        </div>
      </div>

      <div class="side-card">
        <div class="section-head"><h3>Splash Preview</h3></div>
        <div class="section-body">
          <div class="splash-box">
            ${splash ? `<img src="${splash}" alt="Splash Preview">` : '<p class="section-note">Splash preview unavailable.</p>'}
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderPreviewMode() {
  const previews = state.snapshot.previews || {};
  return `
    <div class="preview-only">
      <div class="preview-scroll">
        <div class="side-card">
          <div class="section-head"><h3>Preview</h3></div>
          <div class="section-body">
            <div class="preview-tabs">
              ${state.previewTabs.map(tab => `
                <button class="preview-tab-btn ${tab.key === state.activePreviewTab ? 'active' : ''}" data-preview-tab="${tab.key}">${escapeHtml(tab.label)}</button>
              `).join('')}
            </div>
            <div class="preview-frame" style="margin-top:12px;">
              <pre class="preview-text">${escapeHtml(previews[state.activePreviewTab] || '')}</pre>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderFooter() {
  return `
    <section class="footer">
      <div class="footer-header">
        <div>Project Paths</div>
        <div class="generator-chip">Generator: ${escapeHtml(state.snapshot.generatorStatus || '')}</div>
        <div class="status-copy">${escapeHtml(state.snapshot.status || '')}</div>
      </div>
      <div class="footer-body">
        <div class="paths-grid">
          <div class="path-field">
            <label>Application EXE</label>
            <div class="input-row">
              <input class="text-input" data-key="app_exe" type="text" value="${escapeHtml(state.draft.app_exe || '')}">
              <button class="ghost-btn" type="button" data-action="choose-app-exe">Browse</button>
            </div>
          </div>
          <div class="path-field">
            <label>Output Folder</label>
            <div class="input-row">
              <input class="text-input" data-key="output_dir" type="text" value="${escapeHtml(state.draft.output_dir || '')}">
              <button class="ghost-btn" type="button" data-action="choose-output-dir">Browse</button>
            </div>
          </div>
        </div>
        <div class="footer-actions">
          <button class="ghost-btn" type="button" data-action="validate">Validate</button>
          <button class="primary-btn" type="button" data-action="create-project">Create Project + EXE</button>
        </div>
      </div>
    </section>
  `;
}

function render() {
  captureScrollState();
  if (state.activeTab === 'preview') {
    state.activeTab = 'appinfo';
  }
  const launcherTab = state.schema.find(tab => tab.key === 'launcher');
  if (launcherTab) {
    launcherTab.label = state.snapshot.launcherTabLabel || launcherTab.label;
  }
  appRoot.innerHTML = `
    <div class="app-shell">
      <section class="panel">
        <div class="panel-header">
          <div class="topbar">
            <div>
              <h2>Project Settings</h2>
              <p class="panel-note">Build and preview the PortableApps project from one place.</p>
            </div>
            <button class="danger-btn" type="button" data-action="open-help">Help</button>
          </div>
          <div class="tabs" style="margin-top:14px;">
            ${state.schema.map(tab => `<button class="tab-btn ${tab.key === state.activeTab ? 'active' : ''}" data-tab="${tab.key}">${escapeHtml(tab.label)}</button>`).join('')}
          </div>
        </div>
        <div class="panel-body">
          <div class="form-column"><div class="form-scroll">${renderSections(state.activeTab)}</div></div>
          <aside class="sidebar-column">${renderPreviewSidebar()}</aside>
        </div>
      </section>
      ${renderFooter()}
    </div>
  `;

  bindEvents();
  requestAnimationFrame(() => restoreScrollState());
}

function showValidationModal(payload) {
  const items = payload.items || [];
  modalRoot.innerHTML = `
    <div class="modal-backdrop">
      <div class="modal">
        <div class="modal-header">
          <h3>${escapeHtml(payload.title)}</h3>
        </div>
        <div class="modal-body">
          <div class="validation-list">
            ${items.map(item => `
              <div class="validation-item ${item.level}">
                <strong>${escapeHtml(item.label)}</strong>
                <div style="margin-top:6px;">${escapeHtml(item.message)}</div>
              </div>
            `).join('')}
          </div>
        </div>
        <div class="modal-footer">
          <button class="ghost-btn" type="button" data-close-modal>Close</button>
        </div>
      </div>
    </div>
  `;
  modalRoot.querySelector('[data-close-modal]').addEventListener('click', () => {
    modalRoot.innerHTML = '';
  });
}

async function handleAction(action) {
  if (action === 'choose-app-exe') {
    applySnapshot(await window.pywebview.api.choose_app_exe());
    return render();
  }
  if (action === 'choose-output-dir') {
    applySnapshot(await window.pywebview.api.choose_output_dir());
    return render();
  }
  if (action === 'choose-icon') {
    applySnapshot(await window.pywebview.api.choose_icon());
    return render();
  }
  if (action === 'open-assets-folder') {
    applySnapshot(await window.pywebview.api.open_assets_folder());
    return render();
  }
  if (action === 'open-splash-asset') {
    applySnapshot(await window.pywebview.api.open_splash_asset());
    return render();
  }
  if (action === 'replace-splash-asset') {
    applySnapshot(await window.pywebview.api.replace_splash_asset());
    return render();
  }
  if (action === 'import-registry') {
    applySnapshot(await window.pywebview.api.import_registry_file());
    return render();
  }
  if (action === 'validate') {
    const result = await window.pywebview.api.validate();
    applySnapshot(result.snapshot);
    render();
    return showValidationModal(result);
  }
  if (action === 'create-project') {
    const result = await window.pywebview.api.create_project();
    if (result.snapshot) {
      applySnapshot(result.snapshot);
      render();
    }
    if (result.missingGenerator) {
      const open = confirm(`${result.message}\n\nOpen the PortableApps.com development page?`);
      if (open) {
        await window.pywebview.api.open_generator_download();
      }
      return;
    }
    alert(result.message);
    return;
  }
  if (action === 'open-help') {
    alert('Help is still the old desktop reference for now. We can port it next once the main web shell feels right.');
  }
}

function bindEvents() {
  document.querySelectorAll('[data-tab]').forEach(button => {
    button.addEventListener('click', () => {
      state.activeTab = button.dataset.tab;
      render();
    });
  });

  document.querySelectorAll('[data-preview-tab]').forEach(button => {
    button.addEventListener('click', () => {
      state.activePreviewTab = button.dataset.previewTab;
      render();
    });
  });

  document.querySelectorAll('[data-action]').forEach(button => {
    button.addEventListener('click', () => handleAction(button.dataset.action));
  });

  document.querySelectorAll('[data-key]').forEach(input => {
    const key = input.dataset.key;
    if (input.type === 'checkbox') {
      input.addEventListener('change', async () => {
        state.draft[key] = input.checked;
        applySnapshot(await window.pywebview.api.set_value(key, input.checked));
        render();
      });
      return;
    }
    if (input.tagName === 'SELECT') {
      input.addEventListener('change', async () => {
        state.draft[key] = input.value;
        applySnapshot(await window.pywebview.api.set_value(key, input.value));
        render();
      });
      return;
    }
    input.addEventListener('input', () => {
      state.draft[key] = input.value;
      debounceSave(key, input.value);
    });
  });
}

async function boot() {
  if (state.booted || state.bootInFlight) {
    return;
  }
  if (!window.pywebview || !window.pywebview.api || typeof window.pywebview.api.bootstrap !== 'function') {
    return;
  }
  state.bootInFlight = true;
  const data = await window.pywebview.api.bootstrap();
  state.schema = data.schema;
  state.previewTabs = data.previewTabs;
  applySnapshot(data);
  state.booted = true;
  state.bootInFlight = false;
  render();
}

async function safeBoot() {
  try {
    await boot();
  } catch (error) {
    state.bootInFlight = false;
    const detail = error && error.stack ? error.stack : String(error);
    renderFatal(detail);
  }
}

function waitForPywebviewApi() {
  if (state.booted) {
    return;
  }
  const start = Date.now();
  let waitingMessageShown = false;
  const timer = setInterval(() => {
    if (state.booted) {
      clearInterval(timer);
      return;
    }
    if (window.pywebview && window.pywebview.api && typeof window.pywebview.api.bootstrap === 'function') {
      clearInterval(timer);
      safeBoot();
      return;
    }
    const elapsed = Date.now() - start;
    if (!waitingMessageShown && elapsed > 5000) {
      waitingMessageShown = true;
      renderLoading('Waiting for the webview bridge to start...');
    }
    if (elapsed > 60000) {
      clearInterval(timer);
      const pywebviewState = typeof window.pywebview === 'undefined'
        ? 'window.pywebview missing'
        : (window.pywebview.api ? 'window.pywebview.api present but bootstrap unavailable' : 'window.pywebview present but api missing');
      renderFatal(`Timed out waiting for the pywebview bridge to become available.\n\nBridge state: ${pywebviewState}`);
    }
  }, 150);
}

renderLoading();
window.addEventListener('pywebviewready', safeBoot);
window.addEventListener('DOMContentLoaded', waitForPywebviewApi);
waitForPywebviewApi();
