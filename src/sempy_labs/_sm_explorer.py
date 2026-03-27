# Semantic Model Explorer tab for PBI Fixer.
# Provides a tree view of tables, columns, measures, hierarchies and
# calculation groups with DAX expression preview.

import ipywidgets as widgets
from typing import Optional
from uuid import UUID

from sempy_labs._ui_components import (
    FONT_FAMILY,
    BORDER_COLOR,
    GRAY_COLOR,
    ICON_ACCENT,
    SECTION_BG,
    ICONS,
    build_tree_items,
    create_three_panel_layout,
    create_connection_bar,
    input_label,
    status_html,
    set_status,
    placeholder_panel,
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


def _build_tree(model_data):
    """
    Build tree items from the pre-fetched model data dict.

    Returns (options, key_map)  where key_map values encode the node type
    and identity as  "type:table:name"  strings.
    """
    items = []
    for t_name in sorted(model_data["tables"]):
        t = model_data["tables"][t_name]
        t_type = t["type"]
        icon = "calc_group" if t_type == "CalculationGroup" else "table"
        suffix = ""
        if t["is_hidden"]:
            suffix = " (hidden)"
        items.append((0, icon, f"{t_name}{suffix}", f"table:{t_name}"))

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
    """Return preview text for a selected tree node."""
    parts = key.split(":", 2)
    node_type = parts[0]

    if node_type == "table":
        t_name = parts[1]
        t = model_data["tables"].get(t_name, {})
        lines = [f"Table: {t_name}"]
        lines.append(f"Type: {t.get('type', '')}")
        if t.get("description"):
            lines.append(f"Description: {t['description']}")
        lines.append(f"Columns: {len(t.get('columns', {}))}")
        lines.append(f"Measures: {len(t.get('measures', {}))}")
        lines.append(f"Hierarchies: {len(t.get('hierarchies', {}))}")
        if t.get("calc_items"):
            lines.append(f"Calculation Items: {len(t['calc_items'])}")
        return "\n".join(lines)

    if node_type == "measure":
        t_name, m_name = parts[1], parts[2]
        m = model_data["tables"][t_name]["measures"].get(m_name, {})
        lines = [f"// Measure: {m_name}"]
        lines.append(f"// Table: {t_name}")
        if m.get("description"):
            lines.append(f"// {m['description']}")
        if m.get("format_string"):
            lines.append(f"// Format: {m['format_string']}")
        if m.get("display_folder"):
            lines.append(f"// Folder: {m['display_folder']}")
        lines.append("")
        lines.append(m.get("expression", ""))
        return "\n".join(lines)

    if node_type == "column":
        t_name, c_name = parts[1], parts[2]
        c = model_data["tables"][t_name]["columns"].get(c_name, {})
        lines = [f"Column: {c_name}"]
        lines.append(f"Table: {t_name}")
        lines.append(f"Data Type: {c.get('data_type', '')}")
        lines.append(f"Column Type: {c.get('type', '')}")
        lines.append(f"Hidden: {c.get('is_hidden', False)}")
        if c.get("expression"):
            lines.append(f"\n// Calculated column expression:")
            lines.append(c["expression"])
        return "\n".join(lines)

    if node_type == "hierarchy":
        t_name, h_name = parts[1], parts[2]
        h = model_data["tables"][t_name]["hierarchies"].get(h_name, {})
        lines = [f"Hierarchy: {h_name}"]
        lines.append(f"Table: {t_name}")
        lines.append(f"Levels ({len(h.get('levels', []))}):")
        for i, lvl in enumerate(h.get("levels", []), 1):
            lines.append(f"  {i}. {lvl}")
        return "\n".join(lines)

    if node_type == "calc_item":
        t_name, ci_name = parts[1], parts[2]
        ci = model_data["tables"][t_name]["calc_items"].get(ci_name, {})
        lines = [f"// Calculation Item: {ci_name}"]
        lines.append(f"// Calculation Group: {t_name}")
        lines.append(f"// Ordinal: {ci.get('ordinal', 0)}")
        lines.append("")
        lines.append(ci.get("expression", ""))
        return "\n".join(lines)

    return ""


def sm_explorer_tab(
    workspace: Optional[str | UUID] = None,
    dataset: Optional[str | UUID] = None,
):
    """
    Build the Semantic Model Explorer tab widget.

    Parameters
    ----------
    workspace : str | uuid.UUID, default=None
        Pre-populated workspace name or ID.
    dataset : str | uuid.UUID, default=None
        Pre-populated dataset / semantic model name or ID.

    Returns
    -------
    widgets.VBox
    """
    # -- state --
    _model_data = {}
    _key_map = {}

    # -- connection bar --
    ws_input = widgets.Text(
        value=str(workspace) if workspace else "",
        placeholder="Leave empty for notebook workspace",
        layout=widgets.Layout(width="220px"),
    )
    ds_input = widgets.Text(
        value=str(dataset) if dataset else "",
        placeholder="Semantic model name or ID",
        layout=widgets.Layout(width="220px"),
    )
    load_btn = widgets.Button(
        description="Load",
        button_style="primary",
        layout=widgets.Layout(width="80px"),
    )
    conn_status = status_html()

    conn_bar = create_connection_bar(
        input_label("Workspace"),
        ws_input,
        input_label("Model"),
        ds_input,
        load_btn,
        conn_status,
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

    # -- preview (top-right) --
    preview = widgets.Textarea(
        value="Select a measure to view its DAX expression.",
        disabled=True,
        layout=widgets.Layout(
            width="100%",
            height="300px",
            font_family="monospace",
            border=f"1px solid {BORDER_COLOR}",
            border_radius="8px",
        ),
    )
    preview_label = widgets.HTML(
        value=f'<div style="font-size:12px; font-weight:600; color:{ICON_ACCENT}; '
        f'font-family:{FONT_FAMILY}; text-transform:uppercase; letter-spacing:0.5px; '
        f'margin-bottom:2px;">Preview</div>'
    )
    preview_box = widgets.VBox(
        [preview_label, preview],
        layout=widgets.Layout(flex="1"),
    )

    # -- properties placeholder (bottom-right) --
    props_label = widgets.HTML(
        value=f'<div style="font-size:12px; font-weight:600; color:{ICON_ACCENT}; '
        f'font-family:{FONT_FAMILY}; text-transform:uppercase; letter-spacing:0.5px; '
        f'margin-bottom:2px;">Properties</div>'
    )
    props_placeholder = placeholder_panel("Properties — coming soon", min_height="150px")
    props_box = widgets.VBox(
        [props_label, props_placeholder],
        layout=widgets.Layout(flex="0 0 auto"),
    )

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
        ws = ws_input.value.strip() or None
        ds = ds_input.value.strip()
        if not ds:
            set_status(conn_status, "Enter a semantic model name or ID.", "#ff3b30")
            return

        load_btn.disabled = True
        load_btn.description = "Loading…"
        set_status(conn_status, "Connecting…", GRAY_COLOR)

        try:
            _model_data = _load_model_data(dataset=ds, workspace=ws)
            options, _key_map = _build_tree(_model_data)
            tree.options = options
            tree.value = None

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
            load_btn.description = "Load"

    def on_select(change):
        selected = change.get("new")
        if not selected or selected not in _key_map:
            return
        key = _key_map[selected]
        text = _get_preview_text(_model_data, key)
        preview.value = text

    load_btn.on_click(on_load)
    tree.observe(on_select, names="value")

    # -- assemble --
    return widgets.VBox(
        [conn_bar, tree_header, panels],
        layout=widgets.Layout(
            padding="12px",
            gap="4px",
        ),
    )
