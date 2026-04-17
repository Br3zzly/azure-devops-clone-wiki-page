'use strict';

// ─────────────────────────────────────────────────────────────────────────────
// ADO Wiki Duplicator — content.js
//
// Adds a "Clone" option to the Azure DevOps wiki page context menu.
//
// Resilience strategy (designed to survive ADO UI updates):
//   • Menu detection  : presence of #__bolt-move-page — a stable semantic ID
//                       controlled by ADO's Bolt framework, unlikely to change.
//   • Page identity   : callout ID → aria-controls → trigger button → tree row.
//                       Falls back to mouseover tracking, then mousedown, then
//                       the current URL — so at least one path always works.
//   • Dialog UI       : reuses ADO's own Bolt panel / tree CSS classes so the
//                       look stays consistent even if ADO reskins.
//   • API             : version-pinned to 7.0; uses the browser session cookie.
// ─────────────────────────────────────────────────────────────────────────────

const DUP_ROW_ID = 'ado-dup-injected-row';

// ── Selectors (centralised for easy update if ADO renames them) ──────────────
const SEL_TREE_ROW  = '[role="row"].bolt-tree-row, [role="treeitem"]';
const SEL_WIKI_LINK = 'a[href*="/_wiki/wikis/"]';

// ── URL helpers ──────────────────────────────────────────────────────────────
function parseWikiUrl(href) {
  try {
    const m = new URL(href).pathname.match(
      /^\/([^/]+)\/([^/]+)\/_wiki\/wikis\/([^/]+)\/(\d+)/
    );
    return m ? { org: m[1], project: m[2], wikiId: m[3], pageId: parseInt(m[4], 10) } : null;
  } catch { return null; }
}

function parentPath(p) { const i = p.lastIndexOf('/'); return i <= 0 ? '/' : p.slice(0, i); }
function leafName(p)   { return p.split('/').filter(Boolean).pop() || '/'; }

// ── Page context capture ─────────────────────────────────────────────────────
//
// The 3-dot button may render outside the tree row's DOM subtree, so mousedown
// alone can miss it. We track hover + mousedown + callout-trace as fallbacks.

let lastClickedContext = null;
let lastHoveredContext = null;

document.addEventListener('mouseover', e => {
  const row = e.target.closest?.(SEL_TREE_ROW);
  if (row) lastHoveredContext = capturePageContext(row);
}, true);

document.addEventListener('mousedown', e => {
  lastClickedContext = capturePageContext(e.target);
}, true);

function firstLine(s) {
  return (s?.split('\n').map(x => x.trim()).filter(Boolean)[0]) || '';
}

function capturePageContext(el) {
  if (!el) return null;

  // Walk up to the enclosing tree row (ADO uses both role="row" and role="treeitem")
  let node = el, treeRow = null, depth = 0;
  while (node && node !== document.body && depth < 20) {
    const role = node.getAttribute?.('role');
    if (role === 'treeitem' || (role === 'row' && node.classList?.contains('bolt-tree-row'))) {
      treeRow = node; break;
    }
    node = node.parentElement; depth++;
  }

  if (treeRow) {
    const link = treeRow.querySelector(SEL_WIKI_LINK);
    if (link) {
      const info = parseWikiUrl(link.href);
      if (info) return { kind: 'href', ...info };
    }
    const title = treeRow.getAttribute('aria-label')?.trim() || firstLine(treeRow.textContent);
    const level = parseInt(treeRow.getAttribute('aria-level') || '0', 10) || null;
    if (title) return { kind: 'tree-row', title, level, el: treeRow };
  }

  // Limited-depth href scan (stays inside one sidebar row)
  node = el; depth = 0;
  while (node && node !== document.body && depth < 6) {
    if (node.tagName === 'A' && node.href?.includes('/_wiki/wikis/')) {
      const info = parseWikiUrl(node.href);
      if (info) return { kind: 'href', ...info };
    }
    const link = node.querySelector?.(SEL_WIKI_LINK);
    if (link) {
      const info = parseWikiUrl(link.href);
      if (info) return { kind: 'href', ...info };
    }
    node = node.parentElement; depth++;
  }

  return null;
}

// ── Callout → trigger → tree row tracing ─────────────────────────────────────
//
// ADO's Bolt dropdown links callout ↔ trigger via IDs:
//   Callout: id="__bolt-dropdown-N-callout"  →  Trigger: aria-controls="__bolt-dropdown-N"

function extractContextFromMenu(menuRoot) {
  const callout = menuRoot.closest('[id$="-callout"]') || menuRoot;
  let dropdownId = null;

  if (callout.id?.endsWith('-callout')) {
    dropdownId = callout.id.slice(0, -'-callout'.length);
  }
  if (!dropdownId) {
    const table = menuRoot.querySelector('table[role="menu"][id]');
    if (table) dropdownId = table.id;
  }

  let trigger = null;
  if (dropdownId) {
    trigger = document.querySelector(`[aria-controls="${dropdownId}"]`)
           || document.querySelector(`[aria-owns="${dropdownId}"]`);
  }
  if (!trigger) {
    trigger = document.querySelector('.wiki-tree [aria-expanded="true"]')
           || document.querySelector('[role="tree"] [aria-expanded="true"]');
  }
  if (!trigger) return null;

  const treeRow = trigger.closest(SEL_TREE_ROW);
  return treeRow ? capturePageContext(treeRow) : null;
}

// ── Page info resolution ─────────────────────────────────────────────────────

async function resolvePageInfo(ctx) {
  if (!ctx) return parseWikiUrl(window.location.href);

  if (ctx.kind === 'href') {
    return { org: ctx.org, project: ctx.project, wikiId: ctx.wikiId, pageId: ctx.pageId };
  }

  if (ctx.kind === 'tree-row') {
    const curr = parseWikiUrl(window.location.href);
    if (!curr) return null;
    const { org, project, wikiId } = curr;

    const pathParts = reconstructTreePath(ctx.el, ctx.level, ctx.title);
    const path = '/' + pathParts.join('/');

    try {
      const page = await getPageByPath(org, project, wikiId, path);
      return { org, project, wikiId, pageId: page.id };
    } catch {
      try {
        const tree = await getWikiTree(org, project, wikiId);
        const matches = findByLeafName(tree, ctx.title);
        if (matches.length === 1) return { org, project, wikiId, pageId: matches[0].id };
      } catch { /* exhausted */ }
    }
  }

  return null;
}

function reconstructTreePath(rowEl, level, title) {
  const parts = [title];
  if (!level || level <= 1) return parts;

  const allRows = Array.from(document.querySelectorAll(SEL_TREE_ROW));
  const idx = allRows.indexOf(rowEl);
  if (idx < 0) return parts;

  let expected = level - 1;
  for (let i = idx - 1; i >= 0 && expected >= 1; i--) {
    const l = parseInt(allRows[i].getAttribute('aria-level') || '0', 10);
    if (l === expected) {
      const t = allRows[i].getAttribute('aria-label')?.trim() || firstLine(allRows[i].textContent);
      if (t) { parts.unshift(t); expected--; }
    }
  }
  return parts;
}

function findByLeafName(node, name) {
  const out = [];
  (function walk(n) {
    if (leafName(n.path) === name) out.push(n);
    (n.subPages || []).forEach(walk);
  })(node);
  return out;
}

// ── ADO REST API (version-pinned) ────────────────────────────────────────────
const API_VER = 'api-version=7.0';

function adoBase(org, project, wikiId) {
  return `https://dev.azure.com/${org}/${project}/_apis/wiki/wikis/${wikiId}`;
}

async function apiJSON(url, opts = {}) {
  const res = await fetch(url, { credentials: 'include', ...opts });
  if (!res.ok) {
    let msg = `HTTP ${res.status}`;
    try { const b = await res.json(); msg = b.message || b.value?.Message || msg; } catch { /* use status */ }
    if (res.status === 412) msg = 'A page already exists at that path. Choose a different name.';
    throw new Error(msg);
  }
  return res.json();
}

const getPageById   = (o, p, w, id)   => apiJSON(`${adoBase(o,p,w)}/pages/${id}?includeContent=true&${API_VER}`);
const getWikiTree   = (o, p, w)       => apiJSON(`${adoBase(o,p,w)}/pages?path=/&recursionLevel=full&includeContent=false&${API_VER}`);
const getPageByPath = (o, p, w, path) => apiJSON(`${adoBase(o,p,w)}/pages?path=${encodeURIComponent(path)}&includeContent=true&${API_VER}`);

function createPage(o, p, w, path, content) {
  return apiJSON(
    `${adoBase(o,p,w)}/pages?path=${encodeURIComponent(path)}&${API_VER}`,
    {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', 'If-None-Match': '*' },
      body: JSON.stringify({ content: content || '' })
    }
  );
}

// ── Tree utilities ───────────────────────────────────────────────────────────

function collectDescendants(root, srcPath) {
  const out = [];
  (function walk(n) {
    if (n.path !== srcPath && n.path.startsWith(srcPath + '/')) out.push(n.path);
    (n.subPages || []).forEach(walk);
  })(root);
  return out;
}

// ── Context menu injection ───────────────────────────────────────────────────

function buildDuplicateRow(pageCtx, focusZone) {
  const tr = document.createElement('tr');
  tr.id = DUP_ROW_ID;
  tr.setAttribute('role', 'menuitem');
  tr.setAttribute('tabindex', '-1');
  tr.className = 'bolt-menuitem-row bolt-list-row bolt-menuitem-row-normal cursor-pointer';
  if (focusZone) tr.setAttribute('data-focuszone', focusZone);

  const emptyCell = (withDiv = false) => {
    const td = document.createElement('td');
    td.className = 'bolt-menuitem-cell bolt-list-cell';
    if (withDiv) {
      const d = document.createElement('div');
      d.className = 'bolt-menuitem-cell-content flex-row';
      td.appendChild(d);
    }
    return td;
  };

  tr.appendChild(emptyCell(true));
  tr.appendChild(emptyCell(false));

  const iconTd = document.createElement('td');
  iconTd.className = 'bolt-menuitem-cell bolt-list-cell';
  iconTd.innerHTML = `
    <div class="bolt-menuitem-cell-content bolt-menuitem-cell-icon flex-row">
      <span class="fluent-icons-enabled">
        <span aria-hidden="true" class="flex-noshrink" style="display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px">
          <svg width="14" height="14" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
            <rect x="6.5" y="1" width="11.5" height="11.5" rx="1.8" stroke="currentColor" stroke-width="1.7"/>
            <rect x="1" y="6.5" width="11.5" height="11.5" rx="1.8" fill="currentColor" fill-opacity="0.15" stroke="currentColor" stroke-width="1.7"/>
          </svg>
        </span>
      </span>
    </div>`;
  tr.appendChild(iconTd);

  const textTd = document.createElement('td');
  textTd.className = 'bolt-menuitem-cell bolt-list-cell';
  textTd.innerHTML = '<div class="bolt-menuitem-cell-content bolt-menuitem-cell-text flex-row">Clone</div>';
  tr.appendChild(textTd);

  tr.appendChild(emptyCell(false));
  tr.appendChild(emptyCell(false));
  tr.appendChild(emptyCell(true));

  const activate = async e => {
    e.stopPropagation();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
    const info = await resolvePageInfo(pageCtx);
    if (!info) { showToast('Could not identify which page to clone.', 'error'); return; }
    setTimeout(() => showDuplicateDialog(info), 60);
  };
  tr.addEventListener('click', activate);
  tr.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); activate(e); }
  });

  return tr;
}

function injectIntoMenu(menuRoot) {
  if (menuRoot.querySelector(`#${DUP_ROW_ID}`)) return;

  const movePageRow = menuRoot.querySelector('#__bolt-move-page');
  if (!movePageRow) return;

  const capturedCtx = extractContextFromMenu(menuRoot) || lastHoveredContext || lastClickedContext;
  const focusZone = movePageRow.getAttribute('data-focuszone') || '';
  const newRow = buildDuplicateRow(capturedCtx, focusZone);
  movePageRow.insertAdjacentElement('afterend', newRow);

  const allRows = Array.from(menuRoot.querySelectorAll('[role="menuitem"]'));
  allRows.forEach((row, i) => {
    row.setAttribute('aria-posinset', String(i + 1));
    row.setAttribute('aria-setsize',  String(allRows.length));
  });
}

// ── MutationObserver — detect context menu portals ──────────────────────────
//
// Scoped to ADO's portal host (where callouts/menus are mounted) so we don't
// fire on every unrelated SPA mutation. The host may not exist at content-
// script load, so a one-shot bootstrap observer waits for it.

function watchForContextMenu(host) {
  new MutationObserver(mutations => {
    for (const mut of mutations) {
      for (const node of mut.addedNodes) {
        if (node.nodeType !== 1) continue;
        const menu = node.id === '__bolt-move-page'
          ? node.closest('[role="menu"]')?.closest('[id$="-callout"]') || node.parentElement
          : node.querySelector?.('#__bolt-move-page')
            ? node
            : null;
        if (menu) injectIntoMenu(menu);
      }
    }
  }).observe(host, { childList: true, subtree: true });
}

const portalHost = document.querySelector('.bolt-portal-host');
if (portalHost) {
  watchForContextMenu(portalHost);
} else {
  const bootstrap = new MutationObserver(() => {
    const host = document.querySelector('.bolt-portal-host');
    if (host) { bootstrap.disconnect(); watchForContextMenu(host); }
  });
  bootstrap.observe(document.body, { childList: true, subtree: true });
}


// ─────────────────────────────────────────────────────────────────────────────
// Clone dialog (Bolt panel matching ADO's Move-page dialog)
// ─────────────────────────────────────────────────────────────────────────────

function flattenTree(node, depth = 0) {
  const items = [{ node, depth }];
  if (node.subPages) {
    for (const child of node.subPages) items.push(...flattenTree(child, depth + 1));
  }
  return items;
}

function buildTreeRow(item, state) {
  const { node, depth } = item;
  const hasKids = node.subPages?.length > 0;
  const isRoot  = depth === 0;

  const tr = document.createElement('tr');
  tr.setAttribute('role', 'treeitem');
  tr.setAttribute('aria-level', String(depth + 1));
  tr.setAttribute('aria-busy', 'false');
  if (hasKids) tr.setAttribute('aria-expanded', String(state.expanded.has(node.path)));
  tr.className = 'bolt-tree-row bolt-table-row bolt-list-row single-click-activation';
  tr.dataset.path = node.path;

  const spacerTd = () => {
    const td = document.createElement('td');
    td.className = 'bolt-table-cell-compact bolt-table-cell bolt-list-cell';
    td.setAttribute('role', 'presentation');
    return td;
  };

  const td2 = document.createElement('td');
  td2.className = 'wiki-node-tree-cell bolt-table-cell bolt-list-cell';
  td2.setAttribute('data-column-index', '0');
  td2.setAttribute('role', 'presentation');

  const contentDiv = document.createElement('div');
  contentDiv.className = 'no-padding bolt-table-cell-content flex-row flex-center';

  const nameCell = document.createElement('span');
  nameCell.className = 'wiki-page-name-cell';

  if (isRoot) {
    const rootDiv = document.createElement('div');
    rootDiv.className = 'wiki-tree-root-node';
    rootDiv.textContent = state.wikiId || leafName(node.path);
    nameCell.appendChild(rootDiv);
  } else {
    nameCell.style.paddingLeft = `${(depth - 1) * 20}px`;

    if (hasKids) {
      const chevWrap = document.createElement('span');
      chevWrap.className = 'fluent-icons-enabled';
      const chevron = document.createElement('span');
      chevron.setAttribute('aria-hidden', 'true');
      chevron.className = 'bolt-tree-expand-button font-size cursor-pointer flex-noshrink fabric-icon small';
      chevron.setAttribute('role', 'presentation');
      const updateChevron = () => {
        const open = state.expanded.has(node.path);
        chevron.classList.toggle('ms-Icon--ChevronRightMed', !open);
        chevron.classList.toggle('ms-Icon--ChevronDownMed', open);
      };
      updateChevron();
      chevron.addEventListener('click', e => {
        e.stopPropagation();
        if (state.expanded.has(node.path)) state.expanded.delete(node.path);
        else state.expanded.add(node.path);
        updateChevron();
        tr.setAttribute('aria-expanded', String(state.expanded.has(node.path)));
        updateTreeVisibility(state);
      });
      chevWrap.appendChild(chevron);
      nameCell.appendChild(chevWrap);
      item.updateChevron = updateChevron;
    }

    const draggable = document.createElement('div');
    draggable.className = 'tree-node-draggable wiki-tree-node text-ellipsis';
    const reparent = document.createElement('div');
    reparent.className = 'reparent-target';
    reparent.setAttribute('draggable', 'false');
    const textDiv = document.createElement('div');
    textDiv.className = 'text-ellipsis';

    const iconWrap = document.createElement('span');
    iconWrap.className = 'fluent-icons-enabled';
    const icon = document.createElement('span');
    icon.setAttribute('aria-hidden', 'true');
    icon.className = `type-icon fontSizeML flex-noshrink fabric-icon ${depth === 1 ? 'ms-Icon--Home' : 'ms-Icon--Page'}`;
    iconWrap.appendChild(icon);

    const nameSpan = document.createElement('span');
    nameSpan.className = 'bolt-list-cell-child flex-row flex-center bolt-list-cell-text tree-name-cell tree-name-text';
    nameSpan.textContent = leafName(node.path);

    textDiv.appendChild(iconWrap);
    textDiv.appendChild(nameSpan);
    reparent.appendChild(textDiv);
    draggable.appendChild(reparent);
    nameCell.appendChild(draggable);
  }

  contentDiv.appendChild(nameCell);
  td2.appendChild(contentDiv);

  tr.appendChild(spacerTd());
  tr.appendChild(td2);
  tr.appendChild(spacerTd());

  if (!isRoot) {
    tr.addEventListener('click', () => {
      state.tbody.querySelectorAll('.bolt-tree-row.selected').forEach(r => r.classList.remove('selected'));
      tr.classList.add('selected');
      state.selectedPath = node.path;
      state.onSelect?.(node.path);
    });
  }

  item.tr = tr;
  return tr;
}

function updateTreeVisibility(state) {
  for (let i = 0; i < state.items.length; i++) {
    const item = state.items[i];
    if (item.depth === 0) { item.tr.style.display = ''; continue; }
    let visible = true, target = item.depth - 1;
    for (let j = i - 1; j >= 0 && target >= 0; j--) {
      if (state.items[j].depth === target) {
        if (!state.expanded.has(state.items[j].node.path)) { visible = false; break; }
        target--;
      }
    }
    item.tr.style.display = visible ? '' : 'none';
  }
}

function renderBoltTree(tbody, rootNode, targetPath, onSelect, wikiId) {
  const items    = flattenTree(rootNode);
  const expanded = new Set([rootNode.path]);

  if (targetPath && targetPath !== '/') {
    let path = '';
    for (const part of targetPath.split('/').filter(Boolean)) {
      path += '/' + part;
      expanded.add(path);
    }
  }

  const state = { items, expanded, tbody, selectedPath: targetPath || '/', onSelect, wikiId };

  tbody.innerHTML = '';
  for (const item of items) tbody.appendChild(buildTreeRow(item, state));
  items.forEach(item => item.updateChevron?.());
  updateTreeVisibility(state);

  const target = items.find(item => item.node.path === targetPath);
  if (target) {
    target.tr.classList.add('selected');
    setTimeout(() => target.tr.scrollIntoView({ block: 'nearest' }), 0);
  }

  return state;
}

// ── Dialog ───────────────────────────────────────────────────────────────────

async function showDuplicateDialog(info) {
  if (!info) { showToast('Could not detect wiki page info for cloning.', 'error'); return; }

  document.getElementById('ado-dup-overlay')?.remove();

  const overlay = document.createElement('div');
  overlay.id = 'ado-dup-overlay';
  overlay.className = 'bolt-portal absolute-fill';
  overlay.style.zIndex = '999999';
  overlay.innerHTML = await fetch(chrome.runtime.getURL('dialog.html')).then(r => r.text());

  const portalHost = document.querySelector('.bolt-portal-host') || document.body;
  portalHost.appendChild(overlay);

  const $         = id => document.getElementById(id);
  const loading   = $('ado-dup-loading');
  const treePane  = $('ado-dup-tree-pane');
  const errBox    = $('ado-dup-error');
  const errMsg    = $('ado-dup-errmsg');
  const progWrap  = $('ado-dup-progress');
  const progBar   = $('ado-dup-progbar');
  const progTxt   = $('ado-dup-progtxt');
  const progCnt   = $('ado-dup-progcnt');
  const succBox   = $('ado-dup-success');
  const footer    = $('ado-dup-footer');
  const nameInput = $('ado-dup-name');
  const subToggle = $('ado-dup-subpages');
  const subPill   = subToggle.querySelector('.bolt-toggle-button-pill');
  const subCount  = $('ado-dup-subcount');
  const submitBtn = $('ado-dup-submit');
  const treeBody  = $('ado-dup-tree');

  let destPath = '/', sourcePage = null, wikiTree = null, subpageCount = 0;

  subToggle.addEventListener('click', () => {
    if (subToggle.classList.contains('disabled')) return;
    const on = !subToggle.classList.contains('checked');
    subToggle.classList.toggle('checked', on);
    subPill.setAttribute('aria-checked', String(on));
  });

  const close = () => overlay.remove();
  $('ado-dup-close').onclick = close;
  $('ado-dup-cancel').onclick = close;
  $('ado-dup-dismiss').addEventListener('click', close);
  const escKey = e => { if (e.key === 'Escape') { close(); document.removeEventListener('keydown', escKey); } };
  document.addEventListener('keydown', escKey);

  const showErr = msg => { errMsg.textContent = msg; errBox.style.display = 'flex'; };
  const hideErr = ()  => { errBox.style.display = 'none'; };

  const setSubmitEnabled = enabled => {
    submitBtn.classList.toggle('disabled', !enabled);
    submitBtn.classList.toggle('enabled', enabled);
    submitBtn.setAttribute('aria-disabled', String(!enabled));
    submitBtn.setAttribute('tabindex', enabled ? '0' : '-1');
  };

  function updateSubCount() {
    if (!wikiTree || !sourcePage) return;
    const desc = collectDescendants(wikiTree, sourcePage.path);
    subpageCount = desc.length;
    subCount.textContent = subpageCount > 0 ? `(${subpageCount} page${subpageCount !== 1 ? 's' : ''})` : '(none)';
    const dis = subpageCount === 0;
    subToggle.classList.toggle('disabled', dis);
    subPill.setAttribute('aria-disabled', String(dis));
    subPill.setAttribute('tabindex', dis ? '-1' : '0');
    if (dis) {
      subToggle.classList.remove('checked');
      subPill.setAttribute('aria-checked', 'false');
    }
  }

  const { org, project, wikiId, pageId } = info;

  Promise.all([
    getPageById(org, project, wikiId, pageId),
    getWikiTree(org, project, wikiId)
  ]).then(([page, tree]) => {
    sourcePage = page; wikiTree = tree;
    const pageName = leafName(page.path);
    nameInput.value = `Clone of ${pageName}`;
    $('ado-dup-dialog-title').textContent = `Clone '${pageName}'`;
    destPath = parentPath(page.path);

    const treeState = renderBoltTree(treeBody, tree, destPath, selectedPath => {
      destPath = selectedPath;
      updateSubCount();
      setSubmitEnabled(!!nameInput.value.trim());
    }, wikiId);

    const filterInput = $('ado-dup-filter');
    const filterClear = $('ado-dup-filter-clear');
    filterInput.addEventListener('input', () => {
      const q = filterInput.value.trim().toLowerCase();
      filterClear.classList.toggle('disabled', !q);
      filterClear.classList.toggle('enabled', !!q);
      filterClear.setAttribute('aria-disabled', String(!q));
      if (!q) { updateTreeVisibility(treeState); return; }
      for (const item of treeState.items) {
        if (item.depth === 0) { item.tr.style.display = ''; continue; }
        item.tr.style.display = leafName(item.node.path).toLowerCase().includes(q) ? '' : 'none';
      }
    });
    filterClear.addEventListener('click', () => {
      filterInput.value = '';
      filterInput.dispatchEvent(new Event('input'));
    });

    updateSubCount();
    loading.style.display = 'none';
    treePane.style.display = '';
    setSubmitEnabled(true);
    nameInput.focus();
    nameInput.select();
  }).catch(err => {
    loading.style.display = 'none';
    treePane.style.display = '';
    showErr(`Failed to load wiki: ${err.message}`);
  });

  nameInput.addEventListener('input', () => { hideErr(); setSubmitEnabled(!!nameInput.value.trim()); });
  nameInput.addEventListener('keydown', e => { if (e.key === 'Enter' && !submitBtn.classList.contains('disabled')) submitBtn.click(); });

  submitBtn.addEventListener('click', async () => {
    if (!sourcePage || submitBtn.classList.contains('disabled')) return;
    hideErr();
    const newName = nameInput.value.trim();
    if (!newName) { showErr('Please enter a page name.'); return; }
    const newRoot = destPath === '/' ? `/${newName}` : `${destPath}/${newName}`;
    if (newRoot === sourcePage.path) { showErr('New path is the same as the source.'); return; }

    const pagesToCopy = [{ src: sourcePage.path, dst: newRoot }];
    if (subToggle.classList.contains('checked') && subpageCount > 0) {
      collectDescendants(wikiTree, sourcePage.path).forEach(sp =>
        pagesToCopy.push({ src: sp, dst: newRoot + sp.slice(sourcePage.path.length) })
      );
    }

    const total = pagesToCopy.length;
    treePane.style.display = 'none';
    footer.style.display = 'none';
    progWrap.style.display = 'block';
    progTxt.textContent = total > 1 ? 'Cloning pages…' : 'Cloning page…';
    progBar.style.width = '0%';

    let firstPage = null, failed = 0;

    for (let i = 0; i < pagesToCopy.length; i++) {
      const { src, dst } = pagesToCopy[i];
      progTxt.textContent = `Copying: ${leafName(dst)}`;
      progCnt.textContent = `${i + 1} / ${total}`;
      try {
        const content = src === sourcePage.path
          ? (sourcePage.content || '')
          : ((await getPageByPath(org, project, wikiId, src)).content || '');
        const created = await createPage(org, project, wikiId, dst, content);
        if (i === 0) firstPage = created;
      } catch { failed++; }
      progBar.style.width = `${Math.round(((i + 1) / total) * 100)}%`;
    }

    progWrap.style.display = 'none';
    succBox.style.display = 'flex';
    const succMsg = $('ado-dup-succmsg');
    if (failed === 0) {
      succMsg.textContent = total > 1 ? `${total} pages cloned!` : 'Page cloned!';
    } else {
      succMsg.textContent = `${total - failed} of ${total} copied. ${failed} failed.`;
      succMsg.style.color = '#a4262c';
    }
    const btnContainer = footer.querySelector('.bolt-panel-footer-buttons');
    btnContainer.innerHTML = '';

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'bolt-button enabled bolt-focus-treatment';
    closeBtn.innerHTML = '<span class="bolt-button-text body-m">Close</span>';
    closeBtn.onclick = close;
    btnContainer.appendChild(closeBtn);

    if (firstPage) {
      const gotoBtn = document.createElement('button');
      gotoBtn.type = 'button';
      gotoBtn.className = 'bolt-button enabled primary bolt-focus-treatment';
      gotoBtn.innerHTML = '<span class="bolt-button-text body-m">Go to cloned page</span>';
      gotoBtn.onclick = () => {
        window.location.href = `https://dev.azure.com/${org}/${project}/_wiki/wikis/${wikiId}/${firstPage.id}`;
        close();
      };
      btnContainer.appendChild(gotoBtn);
    }
    footer.style.display = 'flex';
  });
}

function showToast(msg, type = 'info') {
  const t = document.createElement('div');
  t.className = `ado-dup-toast ado-dup-toast--${type}`;
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 4000);
}
