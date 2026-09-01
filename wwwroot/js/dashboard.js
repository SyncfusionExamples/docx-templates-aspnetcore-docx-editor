// dashboard.js — client-side wiring for the Index (dashboard) page.
//
// Two responsibilities:
//   1. Open/close the Create Template modal (server-rendered partial
//      _CreateTemplateModal.cshtml). The modal stays hidden by default
//      and is shown/hidden via the .style.display toggle below.
//   2. AJAX-submit the Create Template form to /Create (the page-model's
//      OnPost handler now returns JSON instead of a redirect). On
//      success the modal closes and the browser navigates to
//      /Editor?id=<newId>; on failure the error message is surfaced
//      inside the modal without a full page reload.
//
// Suggested-role chips inside the modal append their role to the
// Signer Roles input on click (mirrors the old React behaviour).
//
// The EJ2 TextBox search control is initialised by the <ejs-scripts>
// script manager in _Layout.cshtml — no JS init is needed here.

// ---------- Modal open / close ----------
function openCreateTemplateModal() {
    var modal = document.getElementById('tsCreateModal');
    if (!modal) return;
    modal.style.display = 'flex';
    // Focus the Name input so the user can type immediately.
    setTimeout(function () {
        var name = document.getElementById('ts-create-name');
        if (name) name.focus();
    }, 30);
}

function closeCreateTemplateModal() {
    var modal = document.getElementById('tsCreateModal');
    if (!modal) return;
    modal.style.display = 'none';
    // Clear any error message + reset the form so the next open is
    // clean (the user shouldn't see their last-attempt values again).
    var form = document.getElementById('tsCreateForm');
    if (form) form.reset();
    var err = document.getElementById('tsCreateError');
    if (err) { err.style.display = 'none'; err.textContent = ''; }
}

// Close on backdrop click.
document.addEventListener('DOMContentLoaded', function () {
    var modal = document.getElementById('tsCreateModal');
    if (modal) {
        modal.addEventListener('click', function (e) {
            if (e.target === modal) closeCreateTemplateModal();
        });
    }
    // Suggested-role chips append to the roles input.
    document.querySelectorAll('.ts-role-chip-suggestion').forEach(function (btn) {
        btn.addEventListener('click', function () {
            var role = btn.getAttribute('data-role');
            var input = document.getElementById('ts-create-roles');
            if (!input) return;
            var parts = input.value.split(',').map(function (s) { return s.trim(); }).filter(Boolean);
            if (parts.indexOf(role) === -1) parts.push(role);
            input.value = parts.join(', ');
        });
    });
    // Esc closes whichever modal is open.
    document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') {
            closeCreateTemplateModal();
            closeDeleteTemplateDialog();
        }
    });

    // Delete dialog backdrop click closes.
    var delModal = document.getElementById('tsDeleteModal');
    if (delModal) {
        delModal.addEventListener('click', function (e) {
            if (e.target === delModal) closeDeleteTemplateDialog();
        });
        // Enter confirms while the delete dialog is open.
        delModal.addEventListener('keydown', function (e) {
            if (e.key === 'Enter' && delModal.style.display !== 'none') {
                e.preventDefault();
                confirmDeleteTemplate();
            }
        });
    }
});

// ---------- AJAX form submit ----------
function submitCreateTemplate(e) {
    e.preventDefault();
    var form = document.getElementById('tsCreateForm');
    if (!form) return;
    var errEl = document.getElementById('tsCreateError');
    var submitBtn = document.getElementById('tsCreateSubmitBtn');
    // ejs-button host → disable via instance + inner <button>, swap label.
    var subInst = submitBtn ? submitBtn.ej2_instances && submitBtn.ej2_instances[0] : null;
    var subInner = submitBtn ? submitBtn.querySelector('button') : null;
    if (subInst) { subInst.disabled = true; subInst.content = 'Creating…'; if (subInst.dataBind) subInst.dataBind(); }
    if (subInner) subInner.disabled = true;
    if (errEl) { errEl.style.display = 'none'; errEl.textContent = ''; }

    // FormData handles both the text fields and the optional file
    // input (multipart). The /Create page-model's [BindProperty]
    // fields pick these up by name (Id, Type, Name, Description,
    // RolesCsv, BodyFile).
    var fd = new FormData(form);

    fetch('/Create', { method: 'POST', body: fd })
        .then(function (res) {
            // The handler returns JSON for both success (200) and
            // validation failure (400 BadRequest). Parse BOTH through
            // text() first — res.json() throws 'Unexpected end of JSON
            // input' when the body is empty (e.g. an antiforgery 400 or
            // a proxied error page), which masked the real error.
            return res.text().then(function (txt) {
                if (!txt) {
                    throw new Error('Server returned an empty response (HTTP ' + res.status + '). Please restart the app and try again.');
                }
                var j;
                try { j = JSON.parse(txt); }
                catch (_) { throw new Error('Unexpected server response (HTTP ' + res.status + ').'); }
                if (!res.ok || !j.ok) {
                    throw new Error(j.error || ('Request failed (' + res.status + ')'));
                }
                return j;
            });
        })
        .then(function (j) {
            // Success — close the modal + navigate to the editor.
            closeCreateTemplateModal();
            if (j.redirect) {
                window.location.href = j.redirect;
            } else {
                // Fallback: reload the dashboard so the new entry
                // shows in the table.
                window.location.reload();
            }
        })
        .catch(function (err) {
            if (errEl) {
                errEl.textContent = err.message || String(err);
                errEl.style.display = 'block';
            }
        })
        .finally(function () {
            if (subInst) { subInst.disabled = false; subInst.content = 'Create Template'; if (subInst.dataBind) subInst.dataBind(); }
            if (subInner) subInner.disabled = false;
        });
}

// ---------- Delete Template confirmation dialog ----------
// The trash button opens a styled modal (NOT the native confirm()).
// The pending form is captured on click; Confirm submits it
// natively (POST /Index?handler=Delete&id=… → entry + .docx are
// removed server-side → redirect back with a toast).
var tsPendingDeleteForm = null;

function openDeleteTemplateDialog(btn) {
    // btn is the ejs-button host <div> (or the inner button on some
    // callers) — resolve up to the host, then to the enclosing form.
    var host = btn;
    if (btn && btn.tagName === 'BUTTON' && btn.parentElement &&
        btn.parentElement.classList.contains('e-control-wrap')) {
        host = btn.parentElement;
    }
    var form = host ? host.closest('form') : null;
    // Fallback: read the data attr from whichever element has it.
    var name = (btn && btn.getAttribute && btn.getAttribute('data-template-name')) ||
               (host && host.getAttribute && host.getAttribute('data-template-name')) || 'this template';
    if (!form) return;
    tsPendingDeleteForm = form;
    var overlay = document.getElementById('tsDeleteModal');
    if (!overlay) { form.submit(); return; }
    var nameEl = document.getElementById('tsDeleteName');
    if (nameEl) nameEl.textContent = name;
    overlay.style.display = 'flex';
    // The Razor ejs-button renders a wrapper div; focus its inner button.
    var confirmBtn = document.getElementById('tsDeleteConfirmBtn');
    var inner = confirmBtn ? confirmBtn.querySelector('button') : null;
    if (inner) inner.focus();
}

function closeDeleteTemplateDialog() {
    var overlay = document.getElementById('tsDeleteModal');
    if (overlay) overlay.style.display = 'none';
    tsPendingDeleteForm = null;
}

function confirmDeleteTemplate() {
    if (!tsPendingDeleteForm) return;
    var btn = document.getElementById('tsDeleteConfirmBtn');
    if (btn) {
        // ejs-button host → disable via instance + inner <button>, swap label.
        var inst = btn.ej2_instances && btn.ej2_instances[0];
        var inner = btn.querySelector('button');
        if (inst) {
            inst.disabled = true;
            inst.content = 'Deleting…';
            if (inst.dataBind) inst.dataBind();
        }
        if (inner) inner.disabled = true;
    }
    tsPendingDeleteForm.submit();
}
