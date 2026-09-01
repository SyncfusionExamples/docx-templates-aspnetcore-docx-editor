// editor.js — client-side wiring for the Editor Razor Page.
//
// The EJ2 DocumentEditorContainer is server-rendered via the
// <ejs-documenteditorcontainer> tag helper. This script bootstraps
// the runtime behaviour the tag helper can't express declaratively:
//   - Loading the .docx on mount (POST /api/DocumentEditor/ImportFileURL
//     with the template id so the server unions doc + template +
//     common merge fields into the response envelope).
//   - Refreshing the right-rail Merge Fields panel from the envelope.
//   - Save / Download buttons.
//   - Drag-and-drop from the panel into the editor with caret
//     placement at the drop point (ported from the React reference
//     implementation).
//   - Add-Field dialog (POST /api/studio/mergefield).

var MERGE_FIELD_MIME = 'application/x-ts-mergefield';
var MERGE_FIELD_PAYLOAD_MIME = 'application/json';

// Dirty-state tracking for the Save and Publish button.
// - TRUE (button enabled) once any contentChange fires — i.e. the user
//   actually edited the document.
// - Suppressed while open()/openBlank()/import/merge replace the doc
//   (programmatic loads must NOT count as edits — requirement: Save
//   must not enable after the Preview With Data action).
var tsDirty = false;
var tsSuppressContentChange = false;

function setSaveEnabled(enabled, btn) {
    if (arguments.length > 1 && btn) {
        // Busy-path: apply directly to the supplied button host.
        var sHost = btn;
        var sInst = sHost.ej2_instances && sHost.ej2_instances[0];
        var sInner = sHost.querySelector('button');
        if (sInst) { sInst.disabled = !enabled; if (sInst.dataBind) sInst.dataBind(); }
        if (sInner) sInner.disabled = !enabled;
        sHost.classList.toggle('ts-btn-wrapped--disabled', !enabled);
        return;
    }
    var host = document.getElementById('tsSaveBtn');
    if (!host) return;
    // ejs-button renders a host <div> wrapping an inner <button>. Set
    // the EJ2 instance's disabled property AND the inner element's DOM
    // attribute so both the control state and styling stay in sync.
    var inst = host.ej2_instances && host.ej2_instances[0];
    var inner = host.querySelector('button');
    if (inst) {
        inst.disabled = !enabled;
        if (inst.dataBind) inst.dataBind();
    }
    if (inner) inner.disabled = !enabled;
    host.classList.toggle('ts-btn-wrapped--disabled', !enabled);
}

// Programmatic content swaps go through this gate so the contentChange
// events they fire don't mark the document dirty.
function openDocumentInEditor(de, sfdt) {
    tsSuppressContentChange = true;
    try {
        de.open(sfdt);
    } finally {
        // EJ2 fires contentChange synchronously during open(); release
        // after a microtask in case any event lands asynchronously.
        setTimeout(function () { tsSuppressContentChange = false; }, 0);
    }
}

function getEditorContainer() {
    return document.getElementById('DocumentEditor');
}
function getEditor() {
    var c = getEditorContainer();
    return c && c.ej2_instances ? c.ej2_instances[0] : null;
}

// ---------- Bootstrap: load the .docx into the editor ----------
// The actual bootstrap runs from DOMContentLoaded (see below) — it
// waits for EJ2 to finish initialising the DocumentEditorContainer
// instance before grabbing it and loading the .docx.

// ---- Document loading progress overlay (tsDocLoading) ----
// Shown by the Editor page from the first paint (visible-by-default in
// the markup, no flash-of-content behind it). Editor JS updates the
// caption per phase and hides the overlay once the document is open
// (or on failure, with a toast) — see showDocLoading/hideDocLoading.
var tsDocLoadTimer = null;
function showDocLoading(message) {
    var ov = document.getElementById('tsDocLoading');
    if (!ov) return;
    if (message) {
        var sub = document.getElementById('tsDocLoadingSub');
        if (sub) sub.textContent = message;
    }
    ov.style.display = 'flex';
    // Safety net: if something goes catastrophically wrong (JS error
    // between show and hide), don't leave the user stuck behind the
    // overlay forever — auto-dismiss after 30s.
    if (!tsDocLoadTimer) {
        tsDocLoadTimer = setTimeout(function () { hideDocLoading(); }, 30000);
    }
}
function hideDocLoading() {
    if (tsDocLoadTimer) { clearTimeout(tsDocLoadTimer); tsDocLoadTimer = null; }
    var ov = document.getElementById('tsDocLoading');
    if (ov) ov.style.display = 'none';
}

function loadDocxIntoEditor(de, absoluteUrl, templateId) {
    showDocLoading('Importing the document…');
    fetch('/api/DocumentEditor/ImportFileURL', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json;charset=UTF-8' },
        body: JSON.stringify({ fileUrl: absoluteUrl, templateId: templateId }),
    })
    .then(function (res) {
        if (!res.ok) throw new Error('ImportFileURL failed (' + res.status + ')');
        return res.json();
    })
    .then(function (envelope) {
        if (!envelope || !envelope.sfdt) throw new Error('Empty SFDT in response.');
        showDocLoading('Opening the document in the editor…');
        openDocumentInEditor(de, envelope.sfdt);
        // Refresh the right-rail panel from the server-derived union
        // (doc MERGEFIELDs + template.fieldKeys + common-merge-fields).
        if (Array.isArray(envelope.mergeFields)) {
            renderMergeFieldsPanel(envelope.mergeFields);
        }
        hideDocLoading();
        showEditorToast('Document loaded.', 'success');
    })
    .catch(function (err) {
        console.error('Import failed:', err);
        hideDocLoading();
        de.openBlank();
        showEditorToast('Could not open the template document: ' + (err.message || err), 'error');
    });
}

// Re-render the panel's chip list from a fresh list of field keys.
// Replaces the server-rendered <ul> contents so newly-imported
// fields appear without a full page reload.
function renderMergeFieldsPanel(keys) {
    var ul = document.getElementById('tsFieldChipList');
    if (!ul) {
        // The server may have rendered the "no fields" empty state —
        // hoist a fresh <ul> in its place.
        var groups = document.getElementById('tsFieldsGroups');
        if (!groups) return;
        groups.innerHTML = '';
        ul = document.createElement('ul');
        ul.className = 'ts-field-chip-list';
        ul.id = 'tsFieldChipList';
        groups.appendChild(ul);
    }
    ul.innerHTML = '';
    var seen = {};
    keys.forEach(function (key) {
        if (!key || seen[key]) return;
        seen[key] = true;
        var li = document.createElement('li');
        li.className = 'ts-field-chip-li';
        li.setAttribute('draggable', 'true');
        li.setAttribute('data-merge-field', key);
        li.addEventListener('dragstart', function (e) { onChipDragStart(e, key); });
        li.addEventListener('dragend', function (e) { onChipDragEnd(e); });
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'ts-field-chip';
        btn.innerHTML = '<span class="ts-field-chip-label">' + escapeHtml(key) + '</span>';
        btn.addEventListener('click', function () { insertMergeField(key); });
        li.appendChild(btn);
        ul.appendChild(li);
    });
    if (keys.length === 0) {
        var p = document.createElement('p');
        p.className = 'ts-empty';
        p.innerHTML = 'No merge fields yet. Use the <strong>Add Field</strong> button below to add one.';
        ul.parentNode.replaceChild(p, ul);
    }
}

function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
}

// ---------- Insert a merge field at the current caret ----------
function insertMergeField(key) {
    var inst = getEditor();
    if (!inst) return;
    var de = inst.documentEditor;
    if (!de) return;
    var fieldName = String(key).replace(/[\r\n]+/g, '');
    var fieldCode = 'MERGEFIELD  ' + fieldName + '  \\* MERGEFORMAT ';
    de.focusIn();
    de.editor.insertField(fieldCode, '«' + fieldName + '»');
}

// ---------- Save ----------
document.addEventListener('DOMContentLoaded', function () {
    var saveBtn = document.getElementById('tsSaveBtn');
    if (saveBtn) saveBtn.addEventListener('click', function () { handleSave(saveBtn); });
    var pvBtn = document.getElementById('tsPreviewBtn');
    if (pvBtn) pvBtn.addEventListener('click', openPreviewDialog);
    var dlBtn = document.getElementById('tsDownloadBtn');
    if (dlBtn) dlBtn.addEventListener('click', handleDownload);
    // Wire the DocumentEditor (and only the DocumentEditor) as the
    // drop target for merge-field chips.
    wireCanvasDropTarget();
    // has finished initialising the DocumentEditorContainer. The
    // <ejs-documenteditorcontainer> tag helper emits an inline init
    // script that runs `.appendTo("#DocumentEditor")` synchronously
    // BEFORE this file loads, so by DOMContentLoaded the
    // `ej2_instances` array on the host <div> is populated and we can
    // grab the instance directly. (We previously used a `created=`
    // attribute on the tag helper, but that callback fires during
    // `.appendTo()` — before editor.js has loaded — so the function
    // was undefined and the load never happened. Calling it here
    // sidesteps the load-order race entirely.)
    bootstrapEditor();
});

// Kept as a named global so any future `created="onEditorCreated"`
// wiring on the tag helper still resolves. The real entry point is
// bootstrapEditor() called from DOMContentLoaded above.
function onEditorCreated() { bootstrapEditor(); }

function bootstrapEditor() {
    var container = getEditorContainer();
    if (!container || !container.ej2_instances || !container.ej2_instances[0]) {
        // No editor instance — release the loading overlay so the page
        // stays usable (the toast will surface anything odd).
        hideDocLoading();
        return;
    }
    var inst = container.ej2_instances[0];
    var de = inst.documentEditor;
    if (!de) { hideDocLoading(); return; }

    // Unlock edit mode (matches the React app's behaviour).
    de.isReadOnly = false;
    de.restrictEditing = false;

    // Dirty tracking: the first real content change flips the dirty
    // flag and enables Save and Publish. Loads/merges are gated through
    // tsSuppressContentChange so they never enable it.
    if (!de.__tsContentChangeWired) {
        de.__tsContentChangeWired = true;
        de.contentChange = function () {
            if (tsSuppressContentChange) return;
            if (!tsDirty) {
                tsDirty = true;
                setSaveEnabled(true);
            }
        };
    }

    var docxUrlEl = document.getElementById('tsDocxUrl');
    var templateIdEl = document.getElementById('tsTemplateId');
    var docxUrl = docxUrlEl ? docxUrlEl.value : '';
    var templateId = templateIdEl ? templateIdEl.value : '';

    if (docxUrl) {
        // Make the URL absolute (server-side DocxUrl may be stored as
        // /Templates/<slug>.docx). Resolve against the current origin.
        var absolute = docxUrl;
        if (!/^https?:\/\//i.test(absolute)) {
            absolute = window.location.origin + absolute;
        }
        loadDocxIntoEditor(de, absolute, templateId);
    } else {
        // Blank template — no .docx yet. The Merge Fields panel is
        // already server-rendered with the common-merge-fields keys.
        // Gate openBlank so its contentChange doesn't enable Save.
        tsSuppressContentChange = true;
        try { de.openBlank(); }
        finally { setTimeout(function () { tsSuppressContentChange = false; }, 0); }
        hideDocLoading();
    }
}

// Toast helper — transient feedback bar pinned under the editor
// header (success / error / info). Auto-dismisses after ~3.5s.
function showEditorToast(message, kind) {
    var host = document.getElementById('tsToastHost');
    if (!host) return;
    var toast = document.createElement('div');
    toast.className = 'ts-editor-toast ts-editor-toast--' + (kind || 'info');
    toast.textContent = message;
    host.appendChild(toast);
    // Animate in.
    requestAnimationFrame(function () { toast.classList.add('ts-editor-toast--show'); });
    setTimeout(function () {
        toast.classList.remove('ts-editor-toast--show');
        setTimeout(function () { if (toast.parentNode) toast.parentNode.removeChild(toast); }, 250);
    }, 3500);
}

function setSaveBusy(busy, btn) {
    var saveBtn = btn || document.getElementById('tsSaveBtn');
    if (!saveBtn) return;
    // While busy the button is always disabled; when free it reflects
    // the dirty state (a successful save clears tsDirty, which then
    // keeps it disabled until the next real edit).
    setSaveEnabled(busy ? false : tsDirty, saveBtn);
    saveBtn.classList.toggle('ts-btn--busy', busy);
    saveBtn.setAttribute('aria-busy', busy ? 'true' : 'false');
    // EJ2 Button: update the content text via the instance property.
    var inst = saveBtn.ej2_instances && saveBtn.ej2_instances[0];
    if (inst) {
        inst.content = busy ? 'Publishing…' : 'Save and Publish';
        if (inst.dataBind) inst.dataBind();
    }
}

function handleSave(btn) {
    var inst = getEditor();
    if (!inst) return;
    var de = inst.documentEditor;
    if (!de) return;
    var templateId = document.getElementById('tsTemplateId').value;
    var templateName = document.getElementById('tsTemplateName').value;

    var sfdt = '';
    try { sfdt = de.serialize(); }
    catch (err) { showEditorToast('Could not serialize the document.', 'error'); return; }
    if (!sfdt) { showEditorToast('Document is empty.', 'error'); return; }

    // Derive the on-disk slug. For an already-published template we
    // reuse the slug implied by the docxUrl; for a brand-new blank
    // template we derive a fresh unique slug.
    var docxUrlEl = document.getElementById('tsDocxUrl');
    var docxUrl = docxUrlEl ? docxUrlEl.value : '';
    var docxBaseName;
    if (docxUrl) {
        var tail = docxUrl.split('/').pop() || '';
        docxBaseName = tail.replace(/\.docx$/i, '').trim();
    } else {
        var safe = (templateName || 'template').replace(/[^A-Za-z0-9-_]+/g, '_').replace(/^_+|_+$/g, '') || 'template';
        docxBaseName = safe + '-' + Date.now().toString(36);
    }

    setSaveBusy(true, btn);
    fetch('/api/DocumentEditor/Save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            Content: sfdt,
            FileName: docxBaseName,
            Format: 'Docx',
            TemplateId: templateId,
        }),
    })
    .then(function (res) {
        return res.json().then(function (j) {
            if (!res.ok || !j.ok) throw new Error(j.error || ('Save failed (' + res.status + ')'));
            return j;
        });
    })
    .then(function (j) {
        // Keep the hidden tsDocxUrl in sync so a follow-up Save reuses the
        // same slug (matching the server's FileMode.Create overwrite
        // behaviour) instead of minting a new one each time.
        if (j.savedPath && docxUrlEl) docxUrlEl.value = j.savedPath;
        // Saved content == on-disk content → document is clean again.
        tsDirty = false;
        setSaveEnabled(false);
        showEditorToast(j.published
            ? 'Saved and published to "' + j.savedPath + '".'
            : (j.warning || 'Saved to disk (not linked to a template).'), 'success');
    })
    .catch(function (err) {
        showEditorToast('Save failed: ' + (err.message || err), 'error');
    })
    .finally(function () { setSaveBusy(false, btn); });
}

function handleDownload() {
    var inst = getEditor();
    if (!inst) return;
    var de = inst.documentEditor;
    if (!de) return;
    // Pass the TemplateName as-is to the DocumentEditor save API,
    // e.g. documentEditor.save('Donor Impact Letter', 'Docx').
    var templateName = document.getElementById('tsTemplateName').value || 'Document';
    try {
        de.save(templateName, 'Docx');
        showEditorToast('Preparing "' + templateName + '.docx" for download…', 'info');
    } catch (err) {
        showEditorToast('Download failed: ' + (err.message || err), 'error');
    }
}

// ---------- Drag-and-drop from MergeFieldsPanel into the editor ----------
// Ported from the React reference. The drop target is .ts-viewer-canvas
// ONLY (not the panel / header / toolbar). preventDefault on dragover
// marks it as a valid drop target; on drop we place the caret at the
// cursor's position via de.selection.select({x, y, extend: false}) and
// then call insertField.
var isDragOver = false;
var dragDepth = 0;

function buildDragGhost(key) {
    var ghost = document.createElement('div');
    ghost.className = 'ts-drag-ghost';
    ghost.style.position = 'fixed';
    ghost.style.top = '-9999px';
    ghost.style.left = '-9999px';
    ghost.style.zIndex = '2147483647';
    ghost.textContent = '« ' + key + ' »';
    document.body.appendChild(ghost);
    return ghost;
}

function onChipDragStart(e, key) {
    if (!e.dataTransfer) return;
    try {
        e.dataTransfer.setData(MERGE_FIELD_MIME, key);
        e.dataTransfer.setData('text/plain', key);
        e.dataTransfer.setData(MERGE_FIELD_PAYLOAD_MIME,
            JSON.stringify({ source: 'merge-fields-panel', key: key }));
        e.dataTransfer.effectAllowed = 'copy';
        var ghost = buildDragGhost(key);
        try {
            e.dataTransfer.setDragImage(ghost, ghost.offsetWidth / 2, ghost.offsetHeight / 2);
        } catch (_) { /* default ghost */ }
        e.currentTarget.__tsDragGhost = ghost;
    } catch (_) { /* allow native drag */ }
}

function onChipDragEnd(e) {
    var ghost = e.currentTarget && e.currentTarget.__tsDragGhost;
    if (ghost && ghost.parentNode) ghost.parentNode.removeChild(ghost);
    if (e.currentTarget) e.currentTarget.__tsDragGhost = null;
}

function isMergeFieldDrag(dt) {
    if (!dt || !dt.types) return false;
    var types = Array.from(dt.types);
    return types.indexOf(MERGE_FIELD_MIME) !== -1
        || types.indexOf(MERGE_FIELD_PAYLOAD_MIME) !== -1;
}

function readMergeFieldKey(dt) {
    if (!dt) return null;
    try {
        var env = dt.getData(MERGE_FIELD_PAYLOAD_MIME);
        if (env) {
            var parsed = JSON.parse(env);
            if (parsed && parsed.source === 'merge-fields-panel' && typeof parsed.key === 'string') {
                return parsed.key;
            }
        }
    } catch (_) { /* fall through */ }
    var dedicated = dt.getData(MERGE_FIELD_MIME);
    if (dedicated) return dedicated;
    if (isMergeFieldDrag(dt)) return dt.getData('text/plain');
    return null;
}

// Drop target = the DocumentEditor control itself (id="DocumentEditor"),
// NOT the surrounding tsCanvas wrapper. Listeners are registered on that
// single element only, so the drop is not accepted anywhere else (header,
// merge-fields panel, dialogs). The caret is placed at the drop point via
// placeCaretAtDropPoint before the field is inserted.
function wireCanvasDropTarget() {
    // The <ejs-documenteditorcontainer> tag helper emits a host <div
    // id="DocumentEditor"> that EJ2 appends the control into.
    var editorEl = document.getElementById('DocumentEditor');
    if (!editorEl) {
        // Fallback to the canvas wrapper if the control host isn't found.
        editorEl = document.getElementById('tsCanvas');
    }
    if (!editorEl) return;
    var target = editorEl;
    target.addEventListener('dragenter', function (e) {
        if (!isMergeFieldDrag(e.dataTransfer)) return;
        e.preventDefault();
        dragDepth++;
        if (!isDragOver) { isDragOver = true; target.classList.add('ts-viewer-canvas--drag-over'); }
    });
    target.addEventListener('dragover', function (e) {
        if (!isMergeFieldDrag(e.dataTransfer)) return;
        e.preventDefault();
        if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
    });
    target.addEventListener('dragleave', function (e) {
        if (!isMergeFieldDrag(e.dataTransfer)) return;
        dragDepth = Math.max(0, dragDepth - 1);
        if (dragDepth === 0) {
            isDragOver = false;
            target.classList.remove('ts-viewer-canvas--drag-over');
        }
    });
    target.addEventListener('drop', function (e) {
        if (!isMergeFieldDrag(e.dataTransfer)) return;
        e.preventDefault();
        e.stopPropagation();
        dragDepth = 0;
        isDragOver = false;
        target.classList.remove('ts-viewer-canvas--drag-over');
        var key = readMergeFieldKey(e.dataTransfer);
        if (!key) return;
        placeCaretAtDropPoint(e.clientX, e.clientY);
        try { insertMergeField(key); }
        catch (err) { console.error('Drop-insert failed:', err); }
    });
    // Belt-and-braces: block merge-field drops on any OTHER element (the
    // browser would otherwise let the native text drop land in inputs and
    // the page body). Only the DocumentEditor target above accepts them.
    document.addEventListener('dragover', function (e) {
        if (!isMergeFieldDrag(e.dataTransfer)) return;
        if (target.contains(e.target)) return;
        e.preventDefault();
        if (e.dataTransfer) e.dataTransfer.dropEffect = 'none';
    });
    document.addEventListener('drop', function (e) {
        if (!isMergeFieldDrag(e.dataTransfer)) return;
        if (target.contains(e.target)) return;
        e.preventDefault();
        e.stopPropagation();
    });
}

function placeCaretAtDropPoint(dropX, dropY) {
    try {
        var inst = getEditor();
        if (!inst) return;
        var de = inst.documentEditor;
        if (!de || !de.selection || typeof de.selection.select !== 'function') return;
        var rootEl = inst.element;
        if (!rootEl) return;
        var candidates = [
            '.e-de-viewer', '.e-de-page-content', '.e-de-page-container',
            '.e-de-scroll-container', '.e-documenteditor',
            '.e-documenteditor-content', '.e-documenteditor-container',
        ];
        var viewer = null;
        for (var i = 0; i < candidates.length; i++) {
            var el = rootEl.querySelector ? rootEl.querySelector(candidates[i]) : null;
            if (!el || !el.getBoundingClientRect) continue;
            var r = el.getBoundingClientRect();
            if (r.width > 0 && r.height > 0 && dropX >= r.left && dropX <= r.right &&
                dropY >= r.top && dropY <= r.bottom) {
                viewer = el;
                break;
            }
        }
        if (!viewer) {
            var all = rootEl.querySelectorAll ? rootEl.querySelectorAll('*') : [];
            var bestArea = -1;
            for (var j = 0; j < all.length; j++) {
                var e = all[j];
                if (!e || !e.getBoundingClientRect) continue;
                var rr = e.getBoundingClientRect();
                if (rr.width > 0 && rr.height > 0 && dropX >= rr.left && dropX <= rr.right &&
                    dropY >= rr.top && dropY <= rr.bottom) {
                    var a = rr.width * rr.height;
                    if (a > bestArea) { bestArea = a; viewer = e; }
                }
            }
        }
        if (!viewer) viewer = rootEl;
        if (viewer && viewer.getBoundingClientRect) {
            var rect = viewer.getBoundingClientRect();
            var localX = dropX - rect.left;
            var localY = dropY - rect.top;
            var sLeft = 0, sTop = 0;
            try {
                if (typeof viewer.scrollLeft === 'number' && viewer.scrollLeft !== 0) {
                    sLeft = viewer.scrollLeft;
                } else {
                    var p = viewer.parentElement;
                    while (p && !(p.scrollLeft || p.scrollTop)) p = p.parentElement;
                    if (p) { sLeft = p.scrollLeft || 0; sTop = p.scrollTop || 0; }
                }
            } catch (_) { /* ignore */ }
            var finalX = Math.max(0, localX + sLeft);
            var finalY = Math.max(0, localY + sTop);
            de.focusIn();
            de.selection.select({ x: finalX, y: finalY, extend: false });
        }
    } catch (selErr) {
        console.warn('Drop caret placement failed; using existing caret:', selErr);
    }
}

// ---------- Add Field dialog ----------
function openAddFieldDialog() {
    var overlay = document.getElementById('tsAddFieldOverlay');
    if (overlay) overlay.style.display = 'flex';
    var keyInput = document.getElementById('ts-af-key');
    if (keyInput) keyInput.focus();
}

function closeAddFieldDialog() {
    var overlay = document.getElementById('tsAddFieldOverlay');
    if (overlay) overlay.style.display = 'none';
    var err = document.getElementById('tsAddFieldError');
    if (err) { err.style.display = 'none'; err.textContent = ''; }
    var keyInput = document.getElementById('ts-af-key');
    if (keyInput) keyInput.value = '';
}

function submitAddField(e) {
    e.preventDefault();
    var keyInput = document.getElementById('ts-af-key');
    var key = (keyInput.value || '').trim();
    if (!key) return;
    var scopeEl = document.querySelector('input[name="ts-add-scope"]:checked');
    var scope = scopeEl ? scopeEl.value : 'template';
    var templateId = document.getElementById('tsTemplateId').value;
    var errEl = document.getElementById('tsAddFieldError');

    var fd = new FormData();
    fd.append('scope', scope);
    fd.append('key', key);
    if (scope === 'template') fd.append('templateId', templateId);

    fetch('/api/studio/mergefield', { method: 'POST', body: fd })
        .then(function (res) {
            if (!res.ok) return res.json().then(function (j) { throw new Error(j.error || 'Failed'); });
            return res.json();
        })
        .then(function () {
            // Append the new key into the panel so it shows immediately.
            var ul = document.getElementById('tsFieldChipList');
            if (ul) {
                var exists = Array.prototype.some.call(ul.querySelectorAll('.ts-field-chip-li'),
                    function (li) { return li.getAttribute('data-merge-field') === key; });
                if (!exists) {
                    var li = document.createElement('li');
                    li.className = 'ts-field-chip-li';
                    li.setAttribute('draggable', 'true');
                    li.setAttribute('data-merge-field', key);
                    li.addEventListener('dragstart', function (ev) { onChipDragStart(ev, key); });
                    li.addEventListener('dragend', function (ev) { onChipDragEnd(ev); });
                    var btn = document.createElement('button');
                    btn.type = 'button';
                    btn.className = 'ts-field-chip';
                    btn.innerHTML = '<span class="ts-field-chip-label">' + escapeHtml(key) + '</span>';
                    btn.addEventListener('click', function () { insertMergeField(key); });
                    li.appendChild(btn);
                    ul.appendChild(li);
                }
            }
            closeAddFieldDialog();
        })
        .catch(function (err) {
            if (errEl) { errEl.textContent = err.message || String(err); errEl.style.display = 'block'; }
        });
}

// ---------- Preview With Data (MailMerge) dialog ----------
// Flow: open dialog → Browse + select a JSON file (OK enables) → OK:
//   1. serialize the editor to SFDT, saveAsBlob('Docx') → Blob
//   2. Blob → base64 (documentData)
//   3. POST /api/DocumentEditor/MailMerge
//      { fileName, documentData (base64 docx), mailMergeData (JSON string) }
//   4. Server (DocIO) executes the merge, returns merged SFDT
//   5. de.open(sfdt) — the editor now previews the merged document.
//
// Reference merge-data JSON — shown in the collapsible "Reference JSON"
// example inside the dialog so the user can build a matching file.
// The root key is a collection name; GetJsonData takes the FIRST
// property's array as the merge record set.
var PREVIEW_EXAMPLE_JSON = '{\n' +
    '  "Organization": [\n' +
    '    {\n' +
    '      "OrgName": "ABC Foundation",\n' +
    '      "OrgAddress": "123 Main Street, New York, NY 10001",\n' +
    '      "DonorName": "John Smith",\n' +
    '      "DonorAddress": "45 Oak Street, New York, NY 10002",\n' +
    '      "DonationAmount": "$1,500.00",\n' +
    '      "DonationDate": "August 15, 2026"\n' +
    '    }\n' +
    '  ]\n' +
    '}';

// The selected JSON file's text content (NOT stored in a textarea —
// the dialog no longer has one). Set by the file picker's change
// handler; consumed by handlePreviewOk for the MailMerge call.
var tsPreviewJsonContent = '';

function openPreviewDialog() {
    var overlay = document.getElementById('tsPreviewModal');
    if (!overlay) return;
    overlay.style.display = 'flex';
    // Reset any previous file selection so OK starts disabled and only
    // enables after the user Browse-selects a JSON file.
    tsPreviewJsonContent = '';
    var fileInput = document.getElementById('tsPreviewJsonFile');
    if (fileInput) fileInput.value = '';
    var nameEl = document.getElementById('tsPreviewFileName');
    if (nameEl) nameEl.textContent = 'No file selected';
    // Render the reference JSON into the collapsible example block.
    var examplePre = document.getElementById('tsPreviewExamplePre');
    if (examplePre) examplePre.textContent = PREVIEW_EXAMPLE_JSON;
    setPreviewOkEnabled(false);
}

function closePreviewDialog() {
    var overlay = document.getElementById('tsPreviewModal');
    if (overlay) overlay.style.display = 'none';
}

// Enable/disable the OK button — it should only be clickable once a
// JSON file has been successfully read via Browse.
function setPreviewOkEnabled(enabled) {
    var okBtn = document.getElementById('tsPreviewOkBtn');
    if (!okBtn) return;
    okBtn.disabled = !enabled;
}

function showPreviewError(msg) {
    // The dialog no longer has an inline error element — surface errors
    // through the editor toast instead.
    showEditorToast('Preview: ' + (msg || String(msg)), 'error');
}
function hidePreviewError() { /* no inline error element — kept for compat */ }

// File picking: the EJ2 Browse button is a <button>, not a <label>, so
// native label→input forwarding is gone — trigger the hidden file input
// programmatically on click and let the 'change' handler load it.
function triggerBrowseJsonFile() {
    var fileInput = document.getElementById('tsPreviewJsonFile');
    if (fileInput) fileInput.click();
}

// File picker → load into the JSON textarea + show the filename.
document.addEventListener('DOMContentLoaded', function () {
    var fileInput = document.getElementById('tsPreviewJsonFile');
    if (fileInput) {
        fileInput.addEventListener('change', function () {
            var file = fileInput.files && fileInput.files[0];
            var nameEl = document.getElementById('tsPreviewFileName');
            if (!file) {
                if (nameEl) nameEl.textContent = 'No file selected';
                tsPreviewJsonContent = '';
                setPreviewOkEnabled(false);
                return;
            }
            if (nameEl) nameEl.textContent = file.name;
            var reader = new FileReader();
            reader.onload = function () {
                // Hold the file's JSON text in memory (no textarea
                // anymore) and enable OK only on a successful read.
                tsPreviewJsonContent = String(reader.result || '');
                setPreviewOkEnabled(!!tsPreviewJsonContent.trim());
            };
            reader.onerror = function () {
                tsPreviewJsonContent = '';
                setPreviewOkEnabled(false);
                showPreviewError('Could not read the selected JSON file.');
            };
            reader.readAsText(file);
        });
    }
    // EJ2 Browse button → hidden file input (onclick is also set in the
    // markup now; the listener here is a no-op if the element changed).
    var browseBtn = document.getElementById('tsPreviewBrowseBtn');
    if (browseBtn) browseBtn.addEventListener('click', triggerBrowseJsonFile);
    // Also seed the reference example block once on load.
    var examplePre = document.getElementById('tsPreviewExamplePre');
    if (examplePre) examplePre.textContent = PREVIEW_EXAMPLE_JSON;
    // Esc closes whichever studio dialog is open.
    document.addEventListener('keydown', function (e) {
        if (e.key !== 'Escape') return;
        closePreviewDialog();
        closeAddFieldDialog();
    });
    var pvOverlay = document.getElementById('tsPreviewModal');
    if (pvOverlay) {
        pvOverlay.addEventListener('click', function (e) {
            if (e.target === pvOverlay) closePreviewDialog();
        });
    }
});

function setPreviewBusy(busy) {
    var okBtn = document.getElementById('tsPreviewOkBtn');
    if (!okBtn) return;
    // Plain <button> (not ejs-button) — set DOM directly.
    okBtn.disabled = busy;
    okBtn.classList.toggle('ts-btn--busy', busy);
    okBtn.setAttribute('aria-busy', busy ? 'true' : 'false');
    okBtn.innerHTML = busy
        ? '<span class="e-icons e-check"></span> Merging…'
        : '<span class="e-icons e-check"></span> OK';
}

function blobToBase64(blob) {
    return new Promise(function (resolve, reject) {
        var reader = new FileReader();
        reader.onload = function () {
            // reader.result = "data:application/vnd...;base64,XXXX"
            var s = String(reader.result || '');
            var comma = s.indexOf(',');
            resolve(comma >= 0 ? s.substring(comma + 1) : s);
        };
        reader.onerror = function () { reject(new Error('Could not read the document blob.')); };
        reader.readAsDataURL(blob);
    });
}

function handlePreviewOk() {
    // The JSON comes from the file selected via Browse (held in
    // tsPreviewJsonContent — there is no textarea anymore).
    var mailMergeJson = (tsPreviewJsonContent || '').trim();
    if (!mailMergeJson) {
        showPreviewError('Please select a JSON file first.');
        return;
    }

    // Validate the JSON client-side so obvious syntax errors surface here
    // rather than as a 500 from the server.
    try { JSON.parse(mailMergeJson); }
    catch (parseErr) {
        showPreviewError('The selected file is not valid JSON: ' + parseErr.message);
        return;
    }

    var inst = getEditor();
    if (!inst) { showPreviewError('Editor is not ready.'); return; }
    var de = inst.documentEditor;
    if (!de) { showPreviewError('Editor is not ready.'); return; }

    var templateName = (document.getElementById('tsTemplateName').value || 'Document')
        .replace(/[^A-Za-z0-9-_]+/g, '_') || 'Document';

    if (typeof de.saveAsBlob !== 'function') {
        showPreviewError('This editor build does not support saveAsBlob.');
        return;
    }

    setPreviewBusy(true);
    showEditorToast('Executing mail merge…', 'info');

    // Step 1: get the current document as a .docx Blob.
    var blobPromise;
    try {
        // saveAsBlob takes a FormattedDocumentType value; 6 = Docx.
        blobPromise = de.saveAsBlob('Docx');
    } catch (blobErr) {
        setPreviewBusy(false);
        showPreviewError('Could not serialize the document: ' + (blobErr.message || blobErr));
        return;
    }
    if (!blobPromise || typeof blobPromise.then !== 'function') {
        setPreviewBusy(false);
        showPreviewError('Could not serialize the document.');
        return;
    }

    blobPromise
        .then(function (blob) {
            if (!blob || blob.size === 0) throw new Error('The document is empty.');
            return blobToBase64(blob);
        })
        .then(function (base64) {
            // Step 2: ship docx + JSON to the MailMerge endpoint.
            return fetch('/api/DocumentEditor/MailMerge', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    fileName: templateName + '.docx',
                    documentData: base64,
                    mailMergeData: mailMergeJson,
                }),
            })
            .then(function (res) {
                return res.text().then(function (body) {
                    if (!res.ok) {
                        var msg = body;
                        try { var ej = JSON.parse(body); if (ej.error) msg = ej.error; } catch (_) { }
                        throw new Error('MailMerge failed (HTTP ' + res.status + '): ' + (msg || 'no details'));
                    }
                    if (!body) throw new Error('MailMerge returned an empty response.');
                    var mergedSfdt;
                    try { mergedSfdt = JSON.parse(body); }
                    catch (_) { throw new Error('MailMerge returned an unexpected response.'); }
                    return mergedSfdt;
                });
            });
        })
        .then(function (mergedSfdt) {
            // Step 3: put the merged document back into the editor.
            // Gated so the merge's contentChange does NOT enable Save
            // and Publish (the merge isn't an authoring edit).
            openDocumentInEditor(de, mergedSfdt);
            de.isReadOnly = false;
            tsDirty = false;
            setSaveEnabled(false);
            closePreviewDialog();
            showEditorToast('Mail merge complete — previewing the merged document.', 'success');
        })
        .catch(function (err) {
            showPreviewError(err.message || String(err));
            showEditorToast('Preview failed: ' + (err.message || err), 'error');
        })
        .finally(function () { setPreviewBusy(false); });
}
