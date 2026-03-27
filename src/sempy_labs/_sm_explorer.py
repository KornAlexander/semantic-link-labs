# Semantic Model Explorer tab for PBI Fixer.
# Provides a tree view of tables, columns, measures, hierarchies and
# calculation groups with DAX expression preview.

import ipywidgets as widgets

from sempy_labs._ui_components import (
    FONT_FAMILY,
    BORDER_COLOR,
    GRAY_COLOR,
    ICON_ACCENT,
    SECTION_BG,
    ICONS,
    EXPANDED,
    COLLAPSED,
    build_tree_items,
    create_three_panel_layout,
    status_html,
    set_status,
    placeholder_panel,
    panel_box,
)


def _load_model_data(dataset, workspace):
    """
    Connect to a semantic model (read-only) and pre-fetch all metadata
    into a plain Python dict.  The TOM connection is closed after loading.

    Returns
    -------
    dict  with structure:
        {
            "tables": {
                "<TableName>": {
                    "description": str,
                    "is_hidden": bool,
                    "type": str,          # "Table" | "CalculatedTable" | "CalculationGroup"
                    "columns": { "<ColName>": { "data_type": str, "is_hidden": bool, "expression": str|None, "type": str } },
                    "measures": { "<MeasureName>": { "expression": str, "format_string": str, "description": str, "display_folder": str } },
                    "hierarchies": { "<HierName>": { "levels": [str, ...] } },
                    "calc_items": { "<ItemName>": { "expression": str, "ordinal": int } },  # only for calc groups
                },
            },
        }
    """
    from sempy_labs.tom import connect_semantic_model

    model_data = {"tables": {}}

    with connect_semantic_model(dataset=dataset, readonly=True, workspace=workspace) as tm:
        for table in tm.model.Tables:
            t_name = table.Name
            t_info = {
                "description": str(table.Description or ""),
                "is_hidden": bool(table.IsHidden),
                "type": "Table",
                "columns": {},
                "measures": {},
                "hierarchies": {},
                "calc_items": {},
            }

            # Detect calculation groups
            is_calc_group = False
            try:
                if table.CalculationGroup is not None:
                    is_calc_group = True
                    t_info["type"] = "CalculationGroup"
            except Exception:
                pass

            if not is_calc_group:
                # Detect calculated tables
                try:
                    for p in table.Partitions:
                        src_type = str(p.SourceType)
                        if "Calculated" in src_type:
                            t_info["type"] = "CalculatedTable"
                        break
                except Exception:
                    pass

            # Columns
            for col in table.Columns:
                col_type = str(col.Type) if hasattr(col, "Type") else ""
                t_info["columns"][col.Name] = {
                    "data_type": str(col.DataType) if hasattr(col, "DataType") else "",
                    "is_hidden": bool(col.IsHidden),
                    "expression": str(col.Expression) if hasattr(col, "Expression") and col.Expression else None,
                    "type": col_type,
                }

            # Measures
            for m in table.Measures:
                t_info["measures"][m.Name] = {
                    "expression": str(m.Expression) if m.Expression else "",
                    "format_string": str(m.FormatString) if m.FormatString else "",
                    "description": str(m.Description) if m.Description else "",
                    "display_folder": str(m.DisplayFolder) if m.DisplayFolder else "",
                }

            # Hierarchies
            for h in table.Hierarchies:
                levels = []
                for lvl in h.Levels:
                    levels.append(str(lvl.Name))
                t_info["hierarchies"][h.Name] = {"levels": levels}

            # Calculation items (only for calc groups)
            if is_calc_group:
                try:
                    for ci in table.CalculationGroup.CalculationItems:
                        t_info["calc_items"][ci.Name] = {
                            "expression": str(ci.Expression) if ci.Expression else "",
                            "ordinal": int(ci.Ordinal) if hasattr(ci, "Ordinal") else 0,
                        }
                except Exception:
                    pass

            model_data["tables"][t_name] = t_info

    return model_data


def _build_tree(model_data, expanded_tables):
    """
    Build tree items from the pre-fetched model data dict.
    Only includes children of tables that are in ``expanded_tables``.

    Returns (options, key_map).
    """
    items = []
    for t_name in sorted(model_data["tables"]):
        t = model_data["tables"][t_name]
        t_type = t["type"]
        icon = "calc_group" if t_type == "CalculationGroup" else "table"
        is_expanded = t_name in expanded_tables
        marker = EXPANDED if is_expanded else COLLAPSED
        suffix = ""
        if t["is_hidden"]:
            suffix = " (hidden)"
        child_count = len(t["measures"]) + len(t["columns"]) + len(t["hierarchies"]) + len(t["calc_items"])
        items.append((0, icon, f"{marker} {t_name}{suffix}  [{child_count}]", f"table:{t_name}"))

        if not is_expanded:
            continue

        # Measures first (most commonly browsed)
        for m_name in sorted(t["measures"]):
            items.append((1, "measure", m_name, f"measure:{t_name}:{m_name}"))

        # Columns
        for c_name in sorted(t["columns"]):
            c = t["columns"][c_name]
            dt = c["data_type"]
            hidden = " (hidden)" if c["is_hidden"] else ""
            items.append((1, "column", f"{c_name} [{dt}]{hidden}", f"column:{t_name}:{c_name}"))

        # Hierarchies
        for h_name in sorted(t["hierarchies"]):
            h = t["hierarchies"][h_name]
            lvl_str = " → ".join(h["levels"])
            items.append((1, "hierarchy", f"{h_name}  ({lvl_str})", f"hierarchy:{t_name}:{h_name}"))

        # Calculation items (for calc groups)
        for ci_name in sorted(t["calc_items"], key=lambda n: t["calc_items"][n]["ordinal"]):
            items.append((1, "calc_item", ci_name, f"calc_item:{t_name}:{ci_name}"))

    return build_tree_items(items)


def _get_preview_text(model_data, key):
    """Return DAX/expression preview text for a selected tree node."""
    parts = key.split(":", 2)
    node_type = parts[0]

    if node_type == "measure":
        t_name, m_name = parts[1], parts[2]
        m = model_data["tables"][t_name]["measures"].get(m_name, {})
        return m.get("expression", "")

    if node_type == "column":
        t_name, c_name = parts[1], parts[2]
        c = model_data["tables"][t_name]["columns"].get(c_name, {})
        if c.get("expression"):
            return c["expression"]
        return ""

    if node_type == "calc_item":
        t_name, ci_name = parts[1], parts[2]
        ci = model_data["tables"][t_name]["calc_items"].get(ci_name, {})
        return ci.get("expression", "")

    return ""


def _prop_row(label, value):
    """Single property row as HTML table row."""
    return (
        f'<tr><td style="padding:3px 10px 3px 0; font-weight:600; color:#555; '
        f'white-space:nowrap; vertical-align:top;">{label}</td>'
        f'<td style="padding:3px 0; word-break:break-word;">{value}</td></tr>'
    )


def _props_table(rows_html):
    """Wrap property rows in a styled HTML table."""
    return (
        f'<table style="font-size:13px; font-family:{FONT_FAMILY}; '
        f'border-collapse:collapse; width:100%;">'
        f'{rows_html}</table>'
    )


def _get_properties_html(model_data, key):
    """Return properties HTML for the bottom-right panel."""
    parts = key.split(":", 2)
    node_type = parts[0]

    if node_type == "table":
        t_name = parts[1]
        t = model_data["tables"].get(t_name, {})
        rows = _prop_row("Name", t_name)
        rows += _prop_row("Type", t.get("type", ""))
        rows += _prop_row("Hidden", str(t.get("is_hidden", False)))
        if t.get("description"):
            rows += _prop_row("Description", t["description"])
        rows += _prop_row("Columns", str(len(t.get("columns", {}))))
        rows += _prop_row("Measures", str(len(t.get("measures", {}))))
        rows += _prop_row("Hierarchies", str(len(t.get("hierarchies", {}))))
        if t.get("calc_items"):
            rows += _prop_row("Calc Items", str(len(t["calc_items"])))
        return _props_table(rows)

    if node_type == "measure":
        t_name, m_name = parts[1], parts[2]
        m = model_data["tables"][t_name]["measures"].get(m_name, {})
        rows = _prop_row("Name", m_name)
        rows += _prop_row("Table", t_name)
        rows += _prop_row("Object Type", "Measure")
        if m.get("format_string"):
            rows += _prop_row("Format String", m["format_string"])
        if m.get("display_folder"):
            rows += _prop_row("Display Folder", m["display_folder"])
        if m.get("description"):
            rows += _prop_row("Description", m["description"])
        return _props_table(rows)

    if node_type == "column":
        t_name, c_name = parts[1], parts[2]
        c = model_data["tables"][t_name]["columns"].get(c_name, {})
        rows = _prop_row("Name", c_name)
        rows += _prop_row("Table", t_name)
        rows += _prop_row("Data Type", c.get("data_type", ""))
        rows += _prop_row("Column Type", c.get("type", ""))
        rows += _prop_row("Hidden", str(c.get("is_hidden", False)))
        if c.get("expression"):
            rows += _prop_row("Calculated", "Yes")
        return _props_table(rows)

    if node_type == "hierarchy":
        t_name, h_name = parts[1], parts[2]
        h = model_data["tables"][t_name]["hierarchies"].get(h_name, {})
        rows = _prop_row("Name", h_name)
        rows += _prop_row("Table", t_name)
        rows += _prop_row("Object Type", "Hierarchy")
        levels = h.get("levels", [])
        rows += _prop_row("Levels", str(len(levels)))
        for i, lvl in enumerate(levels, 1):
            rows += _prop_row(f"  Level {i}", lvl)
        return _props_table(rows)

    if node_type == "calc_item":
        t_name, ci_name = parts[1], parts[2]
        ci = model_data["tables"][t_name]["calc_items"].get(ci_name, {})
        rows = _prop_row("Name", ci_name)
        rows += _prop_row("Calc Group", t_name)
        rows += _prop_row("Object Type", "Calculation Item")
        rows += _prop_row("Ordinal", str(ci.get("ordinal", 0)))
        return _props_table(rows)

    return ""


def sm_explorer_tab(workspace_input=None, report_input=None):
    """
    Build the Semantic Model Explorer tab widget.

    Parameters
    ----------
    workspace_input : widgets.Text
        Shared workspace text input widget (reads .value on Load).
    report_input : widgets.Text
        Shared report/model text input widget (reads .value on Load).

    Returns
    -------
    widgets.VBox
    """
    # -- state --
    _model_data = {}
    _key_map = {}
    _expanded = set()  # set of table names currently expanded

    # -- Load button + status row --
    load_btn = widgets.Button(
        description="Load Model",
        button_style="primary",
        layout=widgets.Layout(width="110px"),
    )
    expand_btn = widgets.Button(
        description="Expand All",
        layout=widgets.Layout(width="100px"),
    )
    collapse_btn = widgets.Button(
        description="Collapse All",
        layout=widgets.Layout(width="100px"),
    )
    conn_status = status_html()
    load_row = widgets.HBox(
        [load_btn, expand_btn, collapse_btn, conn_status],
        layout=widgets.Layout(align_items="center", gap="8px", margin="0 0 8px 0"),
    )

    # -- tree --
    tree = widgets.Select(
        options=[],
        rows=28,
        layout=widgets.Layout(
            width="320px",
            height="500px",
            font_family="monospace",
        ),
    )

    def _refresh_tree(preserve_selection=None):
        nonlocal _key_map
        options, _key_map = _build_tree(_model_data, _expanded)
        tree.unobserve(on_select, names="value")
        tree.options = options
        if preserve_selection and preserve_selection in options:
            tree.value = preserve_selection
        else:
            tree.value = None
        tree.observe(on_select, names="value")

    # -- preview (top-right, editable) --
    preview = widgets.Textarea(
        value="Select a measure to view its DAX expression.",
        layout=widgets.Layout(
            width="100%",
            height="300px",
            font_family="monospace",
        ),
    )
    preview_label = widgets.HTML(
        value=f'<div style="font-size:12px; font-weight:600; color:{ICON_ACCENT}; '
        f'font-family:{FONT_FAMILY}; text-transform:uppercase; letter-spacing:0.5px; '
        f'margin-bottom:2px;">Expression</div>'
    )
    preview_box = panel_box([preview_label, preview], flex="1")

    # -- properties (bottom-right, dynamic HTML) --
    props_label = widgets.HTML(
        value=f'<div style="font-size:12px; font-weight:600; color:{ICON_ACCENT}; '
        f'font-family:{FONT_FAMILY}; text-transform:uppercase; letter-spacing:0.5px; '
        f'margin-bottom:2px;">Properties</div>'
    )
    props_html = widgets.HTML(
        value=f'<div style="padding:12px; color:{GRAY_COLOR}; font-size:13px; '
        f'font-family:{FONT_FAMILY}; font-style:italic;">Select an object to view properties</div>',
    )
    props_box = panel_box([props_label, props_html], flex="0 0 auto", min_height="150px")

    # -- three-panel layout --
    panels = create_three_panel_layout(tree, preview_box, props_box)

    # -- tree label --
    tree_header = widgets.HTML(
        value=f'<div style="font-size:12px; font-weight:600; color:{ICON_ACCENT}; '
        f'font-family:{FONT_FAMILY}; text-transform:uppercase; letter-spacing:0.5px; '
        f'margin-bottom:2px;">Model Objects</div>'
    )

    # -- handlers --
    def on_load(_):
        nonlocal _model_data, _key_map
        _expanded.clear()
        ws = workspace_input.value.strip() if workspace_input else None
        ws = ws or None
        ds = report_input.value.strip() if report_input else ""
        if not ds:
            set_status(conn_status, "Enter a report / semantic model name in the top bar.", "#ff3b30")
            return

        load_btn.disabled = True
        load_btn.description = "Loading…"
        set_status(conn_status, "Connecting…", GRAY_COLOR)

        try:
            _model_data = _load_model_data(dataset=ds, workspace=ws)
            _refresh_tree()

            n_tables = len(_model_data["tables"])
            n_measures = sum(len(t["measures"]) for t in _model_data["tables"].values())
            n_columns = sum(len(t["columns"]) for t in _model_data["tables"].values())
            set_status(
                conn_status,
                f"Loaded: {n_tables} tables, {n_columns} columns, {n_measures} measures",
                "#34c759",
            )
            preview.value = "Select a measure to view its DAX expression."
        except Exception as e:
            set_status(conn_status, f"Error: {e}", "#ff3b30")
        finally:
            load_btn.disabled = False
            load_btn.description = "Load Model"

    def on_select(change):
        selected = change.get("new")
        if not selected or selected not in _key_map:
            return
        key = _key_map[selected]
        # Toggle expand/collapse if a table row was clicked
        if key.startswith("table:"):
            t_name = key.split(":", 1)[1]
            if t_name in _expanded:
                _expanded.discard(t_name)
            else:
                _expanded.add(t_name)
            _refresh_tree(preserve_selection=selected)
        preview.value = _get_preview_text(_model_data, key)
        props_html.value = _get_properties_html(_model_data, key)

    def on_expand_all(_):
        if _model_data:
            _expanded.update(_model_data["tables"].keys())
            _refresh_tree()

    def on_collapse_all(_):
        _expanded.clear()
        if _model_data:
            _refresh_tree()

    load_btn.on_click(on_load)
    tree.observe(on_select, names="value")
    expand_btn.on_click(on_expand_all)
    collapse_btn.on_click(on_collapse_all)

    # -- assemble --
    return widgets.VBox(
        [load_row, tree_header, panels],
        layout=widgets.Layout(
            padding="12px",
            gap="4px",
        ),
    )
