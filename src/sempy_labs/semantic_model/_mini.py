import json
import html as html_module
from uuid import UUID, uuid4
from typing import Optional
from IPython.display import display, HTML
from sempy._utils._log import log
from sempy_labs.tom import connect_semantic_model


@log
def mini_model_manager(dataset: str | UUID, workspace: Optional[str | UUID] = None):
    """
    Opens an interactive UI for creating and editing mini models (perspectives)
    within a semantic model.

    Parameters
    ----------
    dataset : str | uuid.UUID
        Name or ID of the semantic model.
    workspace : str | uuid.UUID, default=None
        The Fabric workspace name or ID.
        Defaults to None which resolves to the workspace of the attached lakehouse
        or if no lakehouse attached, resolves to the workspace of the notebook.
    """

    with connect_semantic_model(dataset=dataset, workspace=workspace) as tom:
        model_name = html_module.escape(str(tom._dataset_name))
        workspace_name = html_module.escape(str(tom._workspace_name))

        # for p in tom.all_partitions():
        #    if any(str(p.Mode.ModeType)) != 'DirectLake':
        #        print("This function is only available to semantic models with all tables/partitions in Direct Lake mode.")
        #        return

        # ── Collect model metadata ────────────────────────────────────
        tables = []
        for t in tom.model.Tables:
            columns = []
            for c in t.Columns:
                if str(c.Type) != "RowNumber":
                    columns.append(c.Name)
            measures = [m.Name for m in t.Measures]
            hierarchies = [h.Name for h in t.Hierarchies]
            tables.append(
                {
                    "name": t.Name,
                    "columns": sorted(columns),
                    "measures": sorted(measures),
                    "hierarchies": sorted(hierarchies),
                }
            )
        tables.sort(key=lambda x: x["name"])

        # ── Collect perspective membership ────────────────────────────
        perspectives = []
        for p in tom.model.Perspectives:
            members = {}
            for t in tom.model.Tables:
                t_cols = []
                for c in t.Columns:
                    if str(c.Type) != "RowNumber":
                        if tom.in_perspective(c, p.Name):
                            t_cols.append(c.Name)
                t_measures = [
                    m.Name for m in t.Measures if tom.in_perspective(m, p.Name)
                ]
                t_hierarchies = [
                    h.Name for h in t.Hierarchies if tom.in_perspective(h, p.Name)
                ]
                if t_cols or t_measures or t_hierarchies:
                    members[t.Name] = {
                        "columns": t_cols,
                        "measures": t_measures,
                        "hierarchies": t_hierarchies,
                    }
            perspectives.append({"name": p.Name, "members": members})

    # ── Render the UI ─────────────────────────────────────────────────
    _render_mini_model_ui(model_name, workspace_name, tables, perspectives)


def _render_mini_model_ui(model_name, workspace_name, tables, perspectives):
    """Renders the interactive mini model manager HTML UI."""

    uid = uuid4().hex[:8]
    perspectives_json = json.dumps(perspectives)

    # SVG data URI for select dropdown arrow (must be one line in CSS)
    _chevron_svg = (
        "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg'"
        " width='10' height='6' fill='none'%3E%3Cpath d='M1 1l4 4 4-4'"
        " stroke='%236e6e73' stroke-width='1.5' stroke-linecap='round'"
        " stroke-linejoin='round'/%3E%3C/svg%3E"
    )

    # ── CSS ───────────────────────────────────────────────────────────
    styles = f"""
    <style>
    .mm-{uid} {{
        --mm-accent: #0071e3;
        --mm-accent-hover: #0077ED;
        --mm-bg: #ffffff;
        --mm-bg-secondary: #f5f5f7;
        --mm-bg-tertiary: #fafafa;
        --mm-border: #e5e5e5;
        --mm-border-strong: #d2d2d7;
        --mm-text: #1d1d1f;
        --mm-text-secondary: #6e6e73;
        --mm-text-tertiary: #86868b;
        --mm-radius: 14px;
        --mm-radius-sm: 8px;
        --mm-radius-xs: 6px;
        font-family: -apple-system, BlinkMacSystemFont,
                     "Segoe UI", Helvetica, Arial, sans-serif;
        font-size: 14px;
        color: var(--mm-text);
        -webkit-font-smoothing: antialiased;
        max-width: 680px;
        margin: 0;
        padding: 0;
        line-height: 1.4;
    }}
    .mm-{uid} *,
    .mm-{uid} *::before,
    .mm-{uid} *::after {{
        box-sizing: border-box;
        margin: 0;
        padding: 0;
    }}
    .mm-{uid} .mm-container {{
        background: var(--mm-bg);
        border-radius: var(--mm-radius);
        box-shadow: 0 1px 3px rgba(0,0,0,0.06), 0 8px 24px rgba(0,0,0,0.08);
        overflow: hidden;
        border: 1px solid var(--mm-border);
    }}
    /* ── Header ── */
    .mm-{uid} .mm-header {{
        padding: 24px 28px 18px;
        background: var(--mm-bg);
    }}
    .mm-{uid} .mm-title {{
        font-size: 22px;
        font-weight: 700;
        letter-spacing: -0.02em;
        color: var(--mm-text);
        margin: 0 0 2px;
        line-height: 1.2;
    }}
    .mm-{uid} .mm-subtitle {{
        font-size: 13px;
        font-weight: 400;
        color: var(--mm-text-tertiary);
        margin: 0;
        line-height: 1.4;
    }}
    /* ── Form ── */
    .mm-{uid} .mm-form {{
        display: flex;
        gap: 12px;
        padding: 0 28px 18px;
        background: var(--mm-bg);
        align-items: flex-end;
        flex-wrap: wrap;
    }}
    .mm-{uid} .mm-field {{
        display: flex;
        flex-direction: column;
        gap: 5px;
        flex: 1 1 200px;
        min-width: 180px;
    }}
    .mm-{uid} .mm-label {{
        font-size: 11px;
        font-weight: 600;
        color: var(--mm-text-secondary);
        text-transform: uppercase;
        letter-spacing: 0.04em;
        line-height: 1.3;
    }}
    .mm-{uid} .mm-input,
    .mm-{uid} .mm-select {{
        display: block;
        width: 100%;
        height: 36px;
        padding: 0 12px;
        font-size: 13px;
        font-family: inherit;
        background: var(--mm-bg-secondary);
        border: 1px solid transparent;
        border-radius: var(--mm-radius-sm);
        color: var(--mm-text);
        outline: none;
        transition: border-color 0.2s ease, box-shadow 0.2s ease;
        line-height: 34px;
    }}
    .mm-{uid} .mm-input:hover,
    .mm-{uid} .mm-select:hover {{
        border-color: var(--mm-border-strong);
    }}
    .mm-{uid} .mm-input:focus,
    .mm-{uid} .mm-select:focus {{
        background: var(--mm-bg);
        border-color: var(--mm-accent);
        box-shadow: 0 0 0 3px rgba(0, 113, 227, 0.15);
    }}
    .mm-{uid} .mm-input::placeholder {{
        color: var(--mm-text-tertiary);
    }}
    .mm-{uid} .mm-select {{
        appearance: none;
        -webkit-appearance: none;
        cursor: pointer;
        padding-right: 32px;
        background-image: url("{_chevron_svg}");
        background-repeat: no-repeat;
        background-position: right 12px center;
        background-size: 10px 6px;
    }}
    .mm-{uid} .mm-field-row {{
        display: flex;
        gap: 8px;
        align-items: flex-end;
    }}
    .mm-{uid} .mm-field-row .mm-field {{ flex: 1 1 0; }}
    .mm-{uid} .mm-add-btn {{
        display: inline-flex;
        width: 36px;
        height: 36px;
        flex-shrink: 0;
        align-items: center;
        justify-content: center;
        background: var(--mm-bg-secondary);
        border: 1px solid transparent;
        border-radius: var(--mm-radius-sm);
        cursor: pointer;
        color: var(--mm-accent);
        font-size: 20px;
        font-weight: 300;
        line-height: 1;
        transition: background 0.2s ease, border-color 0.2s ease;
    }}
    .mm-{uid} .mm-add-btn:hover {{
        background: rgba(0, 113, 227, 0.06);
        border-color: var(--mm-accent);
    }}
    /* ── Toolbar ── */
    .mm-{uid} .mm-toolbar {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 8px 28px;
        background: var(--mm-bg-tertiary);
        border-top: 1px solid var(--mm-border);
        border-bottom: 1px solid var(--mm-border);
    }}
    .mm-{uid} .mm-search {{
        display: block;
        width: 200px;
        height: 30px;
        padding: 0 10px;
        font-size: 12px;
        font-family: inherit;
        background: var(--mm-bg);
        border: 1px solid var(--mm-border-strong);
        border-radius: var(--mm-radius-xs);
        color: var(--mm-text);
        outline: none;
        transition: border-color 0.2s ease, box-shadow 0.2s ease;
        line-height: 28px;
    }}
    .mm-{uid} .mm-search::placeholder {{
        color: var(--mm-text-tertiary);
    }}
    .mm-{uid} .mm-search:focus {{
        border-color: var(--mm-accent);
        box-shadow: 0 0 0 3px rgba(0, 113, 227, 0.15);
    }}
    .mm-{uid} .mm-bulk-actions {{
        display: flex;
        gap: 4px;
    }}
    .mm-{uid} .mm-link-btn {{
        display: inline-block;
        font-size: 12px;
        font-weight: 500;
        font-family: inherit;
        color: var(--mm-accent);
        background: none;
        border: none;
        cursor: pointer;
        padding: 4px 8px;
        border-radius: var(--mm-radius-xs);
        line-height: 1.4;
        transition: background 0.15s ease;
    }}
    .mm-{uid} .mm-link-btn:hover {{
        background: rgba(0, 113, 227, 0.06);
    }}
    /* ── Tree ── */
    .mm-{uid} .mm-tree {{
        max-height: 460px;
        overflow-y: auto;
        overflow-x: hidden;
    }}
    .mm-{uid} .mm-tree::-webkit-scrollbar {{ width: 6px; }}
    .mm-{uid} .mm-tree::-webkit-scrollbar-track {{
        background: transparent;
    }}
    .mm-{uid} .mm-tree::-webkit-scrollbar-thumb {{
        background: var(--mm-border-strong);
        border-radius: 3px;
    }}
    .mm-{uid} .mm-tg {{
        border-bottom: 1px solid var(--mm-border);
    }}
    .mm-{uid} .mm-tg:last-child {{ border-bottom: none; }}
    .mm-{uid} .mm-tr {{
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 9px 28px 9px 16px;
        cursor: pointer;
        user-select: none;
        transition: background 0.15s ease;
    }}
    .mm-{uid} .mm-tr:hover {{
        background: rgba(0, 0, 0, 0.025);
    }}
    .mm-{uid} .mm-disc {{
        display: inline-block;
        width: 16px;
        height: 16px;
        flex-shrink: 0;
        font-size: 10px;
        line-height: 16px;
        text-align: center;
        color: var(--mm-text-tertiary);
        transition: transform 0.2s ease;
    }}
    .mm-{uid} .mm-tg.mm-open > .mm-tr .mm-disc {{
        transform: rotate(90deg);
    }}
    .mm-{uid} .mm-name {{
        font-size: 13px;
        font-weight: 500;
        color: var(--mm-text);
        flex: 1;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        line-height: 1.4;
    }}
    .mm-{uid} .mm-badge {{
        display: inline-block;
        font-size: 11px;
        font-weight: 500;
        color: var(--mm-text-tertiary);
        background: var(--mm-bg-secondary);
        padding: 1px 7px;
        border-radius: 10px;
        flex-shrink: 0;
        line-height: 1.4;
    }}
    /* ── Children ── */
    .mm-{uid} .mm-kids {{
        display: none;
        padding: 0;
    }}
    .mm-{uid} .mm-tg.mm-open > .mm-kids {{
        display: block;
    }}
    .mm-{uid} .mm-sec {{
        font-size: 10px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: var(--mm-text-tertiary);
        padding: 8px 28px 3px 48px;
        line-height: 1.3;
    }}
    .mm-{uid} .mm-cr {{
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 5px 28px 5px 48px;
        cursor: pointer;
        user-select: none;
        transition: background 0.15s ease;
    }}
    .mm-{uid} .mm-cr:hover {{
        background: rgba(0, 0, 0, 0.025);
    }}
    .mm-{uid} .mm-cr .mm-name {{
        font-weight: 400;
        font-size: 13px;
    }}
    /* ── Checkbox ── */
    .mm-{uid} .mm-cb {{
        position: relative;
        display: inline-block;
        width: 16px;
        height: 16px;
        flex-shrink: 0;
        vertical-align: middle;
    }}
    .mm-{uid} .mm-cb input {{
        position: absolute;
        opacity: 0;
        width: 16px;
        height: 16px;
        margin: 0;
        padding: 0;
        cursor: pointer;
        z-index: 1;
    }}
    .mm-{uid} .mm-cb .mm-ck {{
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        border-radius: 4px;
        border: 1.5px solid var(--mm-border-strong);
        background: var(--mm-bg);
        transition: background 0.15s ease, border-color 0.15s ease;
        pointer-events: none;
    }}
    .mm-{uid} .mm-cb .mm-ck::after {{
        content: '';
        position: absolute;
        display: none;
    }}
    .mm-{uid} .mm-cb input:checked + .mm-ck {{
        background: var(--mm-accent);
        border-color: var(--mm-accent);
    }}
    .mm-{uid} .mm-cb input:checked + .mm-ck::after {{
        display: block;
        left: 4.5px;
        top: 1.5px;
        width: 4.5px;
        height: 8px;
        border: solid #fff;
        border-width: 0 1.8px 1.8px 0;
        transform: rotate(45deg);
    }}
    .mm-{uid} .mm-cb input.mm-ind + .mm-ck {{
        background: var(--mm-accent);
        border-color: var(--mm-accent);
    }}
    .mm-{uid} .mm-cb input.mm-ind + .mm-ck::after {{
        display: block;
        left: 3px;
        top: 6.5px;
        width: 8px;
        height: 0;
        border: solid #fff;
        border-width: 0 0 1.8px 0;
        transform: none;
    }}
    /* ── Footer ── */
    .mm-{uid} .mm-footer {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 14px 28px;
        background: var(--mm-bg-tertiary);
        border-top: 1px solid var(--mm-border);
    }}
    .mm-{uid} .mm-status {{
        font-size: 12px;
        color: var(--mm-text-tertiary);
        line-height: 1.4;
    }}
    .mm-{uid} .mm-save-btn {{
        display: inline-block;
        height: 34px;
        padding: 0 22px;
        font-size: 13px;
        font-weight: 600;
        font-family: inherit;
        color: #fff;
        background: var(--mm-accent);
        border: none;
        border-radius: var(--mm-radius-sm);
        cursor: pointer;
        line-height: 34px;
        transition: background 0.2s ease, box-shadow 0.2s ease;
    }}
    .mm-{uid} .mm-save-btn:hover {{
        background: var(--mm-accent-hover);
        box-shadow: 0 2px 8px rgba(0, 113, 227, 0.25);
    }}
    .mm-{uid} .mm-hidden {{ display: none !important; }}
    .mm-{uid} .mm-brand {{
        padding: 8px 28px;
        font-size: 11px;
        color: var(--mm-text-tertiary);
        text-align: right;
        border-top: 1px solid var(--mm-border);
        background: var(--mm-bg-tertiary);
        line-height: 1.4;
    }}
    .mm-{uid} .mm-brand a {{
        color: inherit;
        text-decoration: underline;
    }}
    </style>
    """

    # ── Build HTML ────────────────────────────────────────────────────
    h = []
    h.append(f'<div class="mm-{uid}">')
    h.append('<div class="mm-container">')

    # Header
    h.append('<div class="mm-header">')
    h.append(f'<div class="mm-title">{model_name}</div>')
    h.append(f'<div class="mm-subtitle">{workspace_name}</div>')
    h.append("</div>")

    # Form
    h.append('<div class="mm-form">')
    h.append('<div class="mm-field">')
    h.append('<label class="mm-label">Mini Model Name</label>')
    h.append(
        f'<input type="text" class="mm-input" id="mm-name-{uid}" '
        f'placeholder="Enter a name\u2026" />'
    )
    h.append("</div>")
    if perspectives:
        h.append('<div class="mm-field-row">')
        h.append('<div class="mm-field">')
        h.append('<label class="mm-label">Existing Mini Models</label>')
        h.append(
            f'<select class="mm-select" id="mm-perspective-{uid}" '
            f'onchange="mmLoad_{uid}(this.value)">'
        )
        h.append(
            '<option value="" disabled selected>'
            "Select a mini model\u2026</option>"
        )
        for p in perspectives:
            ne = html_module.escape(p["name"])
            h.append(f'<option value="{ne}">{ne}</option>')
        h.append("</select>")
        h.append("</div>")
        h.append(
            f'<button class="mm-add-btn" title="New mini model" '
            f'onclick="mmNew_{uid}()">+</button>'
        )
        h.append("</div>")
    h.append("</div>")

    # Toolbar
    h.append('<div class="mm-toolbar">')
    h.append(
        f'<input type="text" class="mm-search" id="mm-search-{uid}" '
        f'placeholder="Filter\u2026" '
        f'oninput="mmFilter_{uid}(this.value)" />'
    )
    h.append('<div class="mm-bulk-actions">')
    h.append(
        f'<button class="mm-link-btn" '
        f'onclick="mmAll_{uid}(true)">Select All</button>'
    )
    h.append(
        f'<button class="mm-link-btn" '
        f'onclick="mmAll_{uid}(false)">Clear All</button>'
    )
    h.append("</div></div>")

    # Tree
    h.append(f'<div class="mm-tree" id="mm-tree-{uid}">')
    for ti, t in enumerate(tables):
        tn = html_module.escape(t["name"])
        cnt = len(t["columns"]) + len(t["measures"]) + len(t["hierarchies"])
        h.append(
            f'<div class="mm-tg" data-mm-table="{tn}" '
            f'id="mm-tg-{uid}-{ti}">'
        )
        h.append(
            f'<div class="mm-tr" onclick="mmExpand_{uid}(this.parentElement)">'
        )
        h.append('<span class="mm-disc">\u25b8</span>')
        h.append(
            f'<label class="mm-cb" onclick="event.stopPropagation()">'
            f'<input type="checkbox" data-mm-role="table" '
            f'data-mm-table="{tn}" '
            f'onchange="mmTblCk_{uid}(this)" />'
            f'<span class="mm-ck"></span></label>'
        )
        h.append(f'<span class="mm-name">{tn}</span>')
        h.append(f'<span class="mm-badge">{cnt}</span>')
        h.append("</div>")
        h.append('<div class="mm-kids">')

        def _child_rows(role, items, label):
            if not items:
                return
            h.append(f'<div class="mm-sec">{label}</div>')
            for name in items:
                ne = html_module.escape(name)
                h.append(
                    f'<div class="mm-cr" '
                    f'onclick="mmRowCk_{uid}(this)">'
                    f'<label class="mm-cb" '
                    f'onclick="event.stopPropagation()">'
                    f'<input type="checkbox" data-mm-role="{role}" '
                    f'data-mm-table="{tn}" data-mm-name="{ne}" '
                    f'onchange="mmChildCk_{uid}(this)" />'
                    f'<span class="mm-ck"></span></label>'
                    f'<span class="mm-name">{ne}</span></div>'
                )

        _child_rows("column", t["columns"], "Columns")
        _child_rows("measure", t["measures"], "Measures")
        _child_rows("hierarchy", t["hierarchies"], "Hierarchies")
        h.append("</div></div>")

    h.append("</div>")

    # Footer
    h.append('<div class="mm-footer">')
    h.append(
        f'<div class="mm-status" id="mm-status-{uid}">'
        f"No selections yet</div>"
    )
    h.append(
        f'<button class="mm-save-btn" id="mm-save-{uid}" '
        f'onclick="mmSave_{uid}()">Save Mini Model</button>'
    )
    h.append("</div>")
    h.append(
        '<div class="mm-brand">Powered by '
        '<a href="https://github.com/microsoft/semantic-link-labs" '
        'target="_blank">Semantic Link Labs</a></div>'
    )
    h.append("</div></div>")

    # ── JavaScript ────────────────────────────────────────────────────
    script = f"""
    <script>
    (function() {{
        var P = {perspectives_json};
        var R = '.mm-{uid}';

        function root() {{ return document.querySelector(R); }}
        function $(id) {{ return document.getElementById(id); }}

        /* Expand / Collapse with animation */
        window.mmExpand_{uid} = function(g) {{
            g.classList.toggle('mm-open');
        }};

        /* Click child row to toggle its checkbox */
        window.mmRowCk_{uid} = function(row) {{
            var cb = row.querySelector('input[type="checkbox"]');
            if (cb) {{ cb.checked = !cb.checked; mmChildCk_{uid}(cb); }}
        }};

        /* Table checkbox → toggle all children */
        window.mmTblCk_{uid} = function(tcb) {{
            var g = tcb.closest('.mm-tg');
            var cbs = g.querySelectorAll('.mm-kids input[type="checkbox"]');
            tcb.classList.remove('mm-ind');
            cbs.forEach(function(c) {{ c.checked = tcb.checked; }});
            syncStatus();
        }};

        /* Child checkbox → update parent tri-state */
        window.mmChildCk_{uid} = function(cb) {{
            syncParent(cb.closest('.mm-tg'));
            syncStatus();
        }};

        function syncParent(g) {{
            var tcb = g.querySelector('input[data-mm-role="table"]');
            var cbs = g.querySelectorAll('.mm-kids input[type="checkbox"]');
            if (!cbs.length) {{ tcb.checked = false; tcb.classList.remove('mm-ind'); return; }}
            var n = 0;
            cbs.forEach(function(c) {{ if (c.checked) n++; }});
            if (n === 0) {{
                tcb.checked = false; tcb.classList.remove('mm-ind');
            }} else if (n === cbs.length) {{
                tcb.checked = true; tcb.classList.remove('mm-ind');
            }} else {{
                tcb.checked = false; tcb.classList.add('mm-ind');
            }}
        }}

        /* New mini model (+ button) */
        window.mmNew_{uid} = function() {{
            var sel = $('mm-perspective-{uid}');
            if (sel) sel.selectedIndex = 0;
            root().querySelectorAll('.mm-tree input[type="checkbox"]')
                .forEach(function(c) {{ c.checked = false; c.classList.remove('mm-ind'); }});
            $('mm-name-{uid}').value = '';
            $('mm-name-{uid}').focus();
            syncStatus();
        }};

        /* Load perspective */
        window.mmLoad_{uid} = function(v) {{
            if (!v) return;
            var r = root();
            r.querySelectorAll('.mm-tree input[type="checkbox"]')
                .forEach(function(c) {{ c.checked = false; c.classList.remove('mm-ind'); }});
            $('mm-name-{uid}').value = v;
            var pData = null;
            for (var i = 0; i < P.length; i++) {{
                if (P[i].name === v) {{ pData = P[i]; break; }}
            }}
            if (!pData) {{ syncStatus(); return; }}
            var m = pData.members;
            for (var tbl in m) {{
                var d = m[tbl];
                ['column','measure','hierarchy'].forEach(function(role) {{
                    var key = role + 's';
                    (d[key] || []).forEach(function(n) {{
                        var sel = 'input[data-mm-table="' +
                            CSS.escape(tbl) + '"][data-mm-role="' +
                            role + '"][data-mm-name="' +
                            CSS.escape(n) + '"]';
                        var cb = r.querySelector(sel);
                        if (cb) cb.checked = true;
                    }});
                }});
            }}
            r.querySelectorAll('.mm-tg').forEach(function(g) {{ syncParent(g); }});
            syncStatus();
        }};

        /* Select All / Clear All */
        window.mmAll_{uid} = function(on) {{
            root().querySelectorAll('.mm-tree input[type="checkbox"]')
                .forEach(function(c) {{
                    if (!c.closest('.mm-hidden')) {{
                        c.checked = on;
                        c.classList.remove('mm-ind');
                    }}
                }});
            syncStatus();
        }};

        /* Filter */
        window.mmFilter_{uid} = function(q) {{
            var r = root();
            q = q.toLowerCase();
            r.querySelectorAll('.mm-tg').forEach(function(g) {{
                var tn = (g.getAttribute('data-mm-table') || '').toLowerCase();
                var rows = g.querySelectorAll('.mm-cr');
                var any = false;
                rows.forEach(function(row) {{
                    var nm = row.querySelector('.mm-name');
                    var t = nm ? nm.textContent.toLowerCase() : '';
                    var ok = !q || t.indexOf(q) !== -1 || tn.indexOf(q) !== -1;
                    row.classList.toggle('mm-hidden', !ok);
                    if (ok) any = true;
                }});
                g.querySelectorAll('.mm-sec').forEach(function(sec) {{
                    var sib = sec.nextElementSibling;
                    var vis = false;
                    while (sib && sib.classList.contains('mm-cr')) {{
                        if (!sib.classList.contains('mm-hidden')) vis = true;
                        sib = sib.nextElementSibling;
                    }}
                    sec.classList.toggle('mm-hidden', !vis);
                }});
                g.classList.toggle('mm-hidden', !(!q || tn.indexOf(q) !== -1 || any));
                if (q && any) g.classList.add('mm-open');
            }});
        }};

        function syncStatus() {{
            var n = root().querySelectorAll(
                '.mm-kids input[type="checkbox"]:checked').length;
            $('mm-status-{uid}').textContent = n === 0
                ? 'No selections yet'
                : n + ' object' + (n !== 1 ? 's' : '') + ' selected';
        }}

        /* Save (placeholder) */
        window.mmSave_{uid} = function() {{
            var name = $('mm-name-{uid}').value.trim();
            if (!name) {{
                $('mm-name-{uid}').focus();
                $('mm-status-{uid}').textContent = 'Please enter a name';
                return;
            }}
            var r = root();
            var sel = {{}};
            r.querySelectorAll('.mm-kids input[type="checkbox"]:checked')
                .forEach(function(cb) {{
                    var tbl = cb.getAttribute('data-mm-table');
                    var role = cb.getAttribute('data-mm-role');
                    var obj = cb.getAttribute('data-mm-name');
                    if (!sel[tbl]) sel[tbl] = {{
                        columns: [], measures: [], hierarchies: []}};
                    sel[tbl][role + 's'].push(obj);
                }});
            var dd = $('mm-perspective-{uid}');
            window._mm_save_data_{uid} = {{
                name: name,
                selections: sel,
                isNew: !dd || dd.selectedIndex <= 0 || dd.value !== name
            }};
            $('mm-status-{uid}').textContent = 'Ready to save \\u201c' + name + '\\u201d';
        }};

        syncStatus();
    }})();
    </script>
    """

    display(HTML(styles + "\n".join(h) + script))
