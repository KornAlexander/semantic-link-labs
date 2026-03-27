# Perspective Editor tab for PBI Fixer.
# Inspired by Michael Kovalsky's PerspectiveEditor.cs — ported to ipywidgets.
# Provides a tri-state checkbox tree to create/modify perspectives on a semantic model.

import ipywidgets as widgets

from sempy_labs._ui_components import (
    FONT_FAMILY,
    BORDER_COLOR,
    GRAY_COLOR,
    ICON_ACCENT,
    SECTION_BG,
    ICONS,
    status_html,
    set_status,
    panel_box,
)


def _load_perspective_data(dataset, workspace):
    """Load model objects and existing perspectives via TOM."""
    from sempy_labs.tom import connect_semantic_model

    data = {"tables": {}, "perspectives": []}

    with connect_semantic_model(dataset=dataset, readonly=True, workspace=workspace) as tm:
        for p in tm.model.Perspectives:
            data["perspectives"].append(str(p.Name))

        for table in tm.model.Tables:
            t_name = table.Name
            t_info = {
                "is_hidden": bool(table.IsHidden),
                "columns": {},
                "measures": {},
                "hierarchies": {},
            }
            for col in table.Columns:
                t_info["columns"][col.Name] = {"is_hidden": bool(col.IsHidden)}
            for m in table.Measures:
                t_info["measures"][m.Name] = {"is_hidden": bool(m.IsHidden)}
            for h in table.Hierarchies:
                t_info["hierarchies"][h.Name] = {"is_hidden": bool(h.IsHidden)}
            data["tables"][t_name] = t_info

    return data


def _load_perspective_members(dataset, workspace, persp_name):
    """Load which objects are in a given perspective."""
    from sempy_labs.tom import connect_semantic_model

    members = {}
    with connect_semantic_model(dataset=dataset, readonly=True, workspace=workspace) as tm:
        for table in tm.model.Tables:
            t_name = table.Name
            members[t_name] = {"columns": set(), "measures": set(), "hierarchies": set()}
            for col in table.Columns:
                if col.InPerspective[persp_name]:
                    members[t_name]["columns"].add(col.Name)
            for m in table.Measures:
                if m.InPerspective[persp_name]:
                    members[t_name]["measures"].add(m.Name)
            for h in table.Hierarchies:
                if h.InPerspective[persp_name]:
                    members[t_name]["hierarchies"].add(h.Name)
    return members


def perspective_editor_tab(workspace_input=None, report_input=None):
    """Build the Perspective Editor tab widget."""
    _data = {}
    _checkboxes = {}  # key -> checkbox widget
    _table_checkboxes = {}  # table_name -> checkbox widget

    load_btn = widgets.Button(description="Load Model", button_style="primary", layout=widgets.Layout(width="110px"))
    conn_status = status_html()

    mode_selector = widgets.RadioButtons(
        options=["Create New Perspective", "Modify Existing Perspective"],
        value="Create New Perspective",
        layout=widgets.Layout(width="auto"),
    )
    persp_dropdown = widgets.Dropdown(
        options=[],
        layout=widgets.Layout(width="250px"),
    )
    persp_dropdown.layout.display = "none"
    persp_name_input = widgets.Text(
        placeholder="Enter perspective name",
        layout=widgets.Layout(width="250px"),
    )

    save_btn = widgets.Button(
        description="Save Perspective",
        button_style="warning",
        disabled=True,
        layout=widgets.Layout(width="150px"),
    )
    save_status = status_html()

    def on_mode_change(change):
        if mode_selector.value == "Create New Perspective":
            persp_dropdown.layout.display = "none"
            persp_name_input.layout.display = ""
            persp_name_input.value = ""
            _clear_all_checkboxes()
        else:
            persp_dropdown.layout.display = ""
            persp_name_input.layout.display = "none"
            if persp_dropdown.value:
                _load_existing_perspective(persp_dropdown.value)

    mode_selector.observe(on_mode_change, names="value")

    def on_persp_selected(change):
        if change.get("new") and mode_selector.value == "Modify Existing Perspective":
            persp_name_input.value = change["new"]
            _load_existing_perspective(change["new"])

    persp_dropdown.observe(on_persp_selected, names="value")

    # Tree container (will be populated with checkboxes)
    tree_container = widgets.VBox(layout=widgets.Layout(
        max_height="500px",
        overflow_y="auto",
        border=f"1px solid {BORDER_COLOR}",
        border_radius="8px",
        padding="8px",
    ))

    select_all_btn = widgets.Button(description="Select All", layout=widgets.Layout(width="100px"))
    deselect_all_btn = widgets.Button(description="Deselect All", layout=widgets.Layout(width="100px"))

    load_row = widgets.HBox(
        [load_btn, select_all_btn, deselect_all_btn, conn_status],
        layout=widgets.Layout(align_items="center", gap="8px", margin="0 0 8px 0"),
    )

    def _clear_all_checkboxes():
        for cb in _checkboxes.values():
            cb.value = False
        for cb in _table_checkboxes.values():
            cb.value = False

    def _load_existing_perspective(persp_name):
        ws = workspace_input.value.strip() if workspace_input else None
        ws = ws or None
        ds = report_input.value.strip() if report_input else ""
        if not ds or not _data:
            return
        try:
            members = _load_perspective_members(ds, ws, persp_name)
            _clear_all_checkboxes()
            for t_name, mem in members.items():
                all_cols = set(_data["tables"].get(t_name, {}).get("columns", {}).keys())
                all_meas = set(_data["tables"].get(t_name, {}).get("measures", {}).keys())
                all_hier = set(_data["tables"].get(t_name, {}).get("hierarchies", {}).keys())
                total = len(all_cols) + len(all_meas) + len(all_hier)
                checked = len(mem["columns"]) + len(mem["measures"]) + len(mem["hierarchies"])
                if t_name in _table_checkboxes and total > 0 and checked == total:
                    _table_checkboxes[t_name].value = True
                for c_name in mem["columns"]:
                    k = f"col:{t_name}:{c_name}"
                    if k in _checkboxes:
                        _checkboxes[k].value = True
                for m_name in mem["measures"]:
                    k = f"mea:{t_name}:{m_name}"
                    if k in _checkboxes:
                        _checkboxes[k].value = True
                for h_name in mem["hierarchies"]:
                    k = f"hie:{t_name}:{h_name}"
                    if k in _checkboxes:
                        _checkboxes[k].value = True
        except Exception:
            pass

    def _build_checkbox_tree():
        _checkboxes.clear()
        _table_checkboxes.clear()
        rows = []

        for t_name in sorted(_data["tables"]):
            t = _data["tables"][t_name]
            hidden = " (hidden)" if t["is_hidden"] else ""
            t_cb = widgets.Checkbox(value=False, indent=False, layout=widgets.Layout(width="22px"))
            _table_checkboxes[t_name] = t_cb

            child_keys = []
            t_label = widgets.HTML(
                value=f'<span style="font-size:13px; font-weight:600; font-family:{FONT_FAMILY};">'
                f'{ICONS["table"]} {t_name}{hidden}</span>'
            )
            rows.append(widgets.HBox([t_cb, t_label], layout=widgets.Layout(align_items="center", gap="4px")))

            for c_name in sorted(t["columns"]):
                k = f"col:{t_name}:{c_name}"
                child_keys.append(k)
                cb = widgets.Checkbox(value=False, indent=False, layout=widgets.Layout(width="22px"))
                _checkboxes[k] = cb
                lbl = widgets.HTML(
                    value=f'<span style="font-size:12px; font-family:{FONT_FAMILY}; color:#555;">'
                    f'{ICONS["column"]} {c_name}</span>'
                )
                rows.append(widgets.HBox(
                    [cb, lbl],
                    layout=widgets.Layout(align_items="center", gap="4px", margin="0 0 0 28px"),
                ))

            for m_name in sorted(t["measures"]):
                k = f"mea:{t_name}:{m_name}"
                child_keys.append(k)
                cb = widgets.Checkbox(value=False, indent=False, layout=widgets.Layout(width="22px"))
                _checkboxes[k] = cb
                lbl = widgets.HTML(
                    value=f'<span style="font-size:12px; font-family:{FONT_FAMILY}; color:#555;">'
                    f'{ICONS["measure"]} {m_name}</span>'
                )
                rows.append(widgets.HBox(
                    [cb, lbl],
                    layout=widgets.Layout(align_items="center", gap="4px", margin="0 0 0 28px"),
                ))

            for h_name in sorted(t["hierarchies"]):
                k = f"hie:{t_name}:{h_name}"
                child_keys.append(k)
                cb = widgets.Checkbox(value=False, indent=False, layout=widgets.Layout(width="22px"))
                _checkboxes[k] = cb
                lbl = widgets.HTML(
                    value=f'<span style="font-size:12px; font-family:{FONT_FAMILY}; color:#555;">'
                    f'{ICONS["hierarchy"]} {h_name}</span>'
                )
                rows.append(widgets.HBox(
                    [cb, lbl],
                    layout=widgets.Layout(align_items="center", gap="4px", margin="0 0 0 28px"),
                ))

            # Table checkbox toggles all children
            _captured_keys = list(child_keys)
            def _make_table_toggle(keys):
                def _toggle(change):
                    for ck in keys:
                        if ck in _checkboxes:
                            _checkboxes[ck].value = change["new"]
                return _toggle
            t_cb.observe(_make_table_toggle(_captured_keys), names="value")

        tree_container.children = rows
        save_btn.disabled = False

    def on_load(_):
        nonlocal _data
        ws = workspace_input.value.strip() if workspace_input else None
        ws = ws or None
        ds = report_input.value.strip() if report_input else ""
        if not ds:
            set_status(conn_status, "Enter a semantic model name in the top bar.", "#ff3b30")
            return
        load_btn.disabled = True
        load_btn.description = "Loading\u2026"
        set_status(conn_status, "Connecting\u2026", GRAY_COLOR)
        try:
            _data = _load_perspective_data(ds, ws)
            _build_checkbox_tree()
            persp_dropdown.options = _data["perspectives"]
            n_t = len(_data["tables"])
            set_status(conn_status, f"Loaded: {n_t} tables, {len(_data['perspectives'])} perspectives", "#34c759")
        except Exception as e:
            set_status(conn_status, f"Error: {e}", "#ff3b30")
        finally:
            load_btn.disabled = False
            load_btn.description = "Load Model"

    def on_save(_):
        ws = workspace_input.value.strip() if workspace_input else None
        ws = ws or None
        ds = report_input.value.strip() if report_input else ""
        p_name = persp_name_input.value.strip()
        if not ds:
            set_status(save_status, "No model loaded.", "#ff3b30")
            return
        if not p_name:
            set_status(save_status, "Enter a perspective name.", "#ff3b30")
            return
        save_btn.disabled = True
        save_btn.description = "Saving\u2026"
        set_status(save_status, "Writing via XMLA\u2026", GRAY_COLOR)
        try:
            from sempy_labs.tom import connect_semantic_model
            with connect_semantic_model(dataset=ds, readonly=False, workspace=ws) as tm:
                if not any(p.Name == p_name for p in tm.model.Perspectives):
                    tm.model.AddPerspective(p_name)
                # Clear all
                for table in tm.model.Tables:
                    table.InPerspective[p_name] = False
                # Set checked items
                for key, cb in _checkboxes.items():
                    if not cb.value:
                        continue
                    parts = key.split(":", 2)
                    obj_type, t_name, obj_name = parts[0], parts[1], parts[2]
                    if obj_type == "col":
                        tm.model.Tables[t_name].Columns[obj_name].InPerspective[p_name] = True
                    elif obj_type == "mea":
                        tm.model.Tables[t_name].Measures[obj_name].InPerspective[p_name] = True
                    elif obj_type == "hie":
                        tm.model.Tables[t_name].Hierarchies[obj_name].InPerspective[p_name] = True
                tm.model.SaveChanges()
            set_status(save_status, f"\u2713 Perspective '{p_name}' saved.", "#34c759")
            if p_name not in _data["perspectives"]:
                _data["perspectives"].append(p_name)
                persp_dropdown.options = _data["perspectives"]
        except Exception as e:
            set_status(save_status, f"Error: {e}", "#ff3b30")
        finally:
            save_btn.disabled = False
            save_btn.description = "Save Perspective"

    def on_select_all(_):
        for cb in _checkboxes.values():
            cb.value = True
        for cb in _table_checkboxes.values():
            cb.value = True

    def on_deselect_all(_):
        _clear_all_checkboxes()

    load_btn.on_click(on_load)
    save_btn.on_click(on_save)
    select_all_btn.on_click(on_select_all)
    deselect_all_btn.on_click(on_deselect_all)

    header = widgets.HTML(
        value=f'<div style="font-size:12px; font-weight:600; color:{ICON_ACCENT}; font-family:{FONT_FAMILY}; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:2px;">Perspective Editor</div>'
    )
    config_row = widgets.VBox([
        mode_selector,
        widgets.HBox([persp_name_input, persp_dropdown], layout=widgets.Layout(gap="8px")),
    ], layout=widgets.Layout(gap="4px", margin="0 0 8px 0"))

    save_row = widgets.HBox([save_btn, save_status], layout=widgets.Layout(align_items="center", gap="8px", margin="8px 0 0 0"))

    return widgets.VBox(
        [load_row, header, config_row, tree_container, save_row],
        layout=widgets.Layout(padding="12px", gap="4px"),
    )