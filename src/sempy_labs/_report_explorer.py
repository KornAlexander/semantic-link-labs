# Report Explorer tab for PBI Fixer.
# Provides a tree view of report pages and visuals.

import ipywidgets as widgets

from sempy_labs._ui_components import (
    FONT_FAMILY,
    BORDER_COLOR,
    GRAY_COLOR,
    ICON_ACCENT,
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


def _load_report_data(report, workspace):
    """
    Connect to a report (read-only) and pre-fetch page/visual metadata
    into a plain Python dict.  The report connection is closed after loading.

    Returns
    -------
    dict  with structure:
        {
            "pages": {
                "<PageName>": {
                    "display_name": str,
                    "width": int,
                    "height": int,
                    "hidden": bool,
                    "visual_count": int,
                    "visuals": {
                        "<VisualName>": {
                            "type": str,
                            "display_type": str,
                            "x": int, "y": int,
                            "width": int, "height": int,
                            "hidden": bool,
                            "title": str,
                        },
                    },
                },
            },
            "format": str,  # "PBIR" | "PBIRLegacy" | ...
        }
    """
    from sempy_labs.report import connect_report

    report_data = {"pages": {}, "format": ""}

    with connect_report(report=report, readonly=True, workspace=workspace) as rw:
        report_data["format"] = str(getattr(rw, "format", ""))

        pages_df = rw.list_pages()
        visuals_df = rw.list_visuals()

        for _, row in pages_df.iterrows():
            p_name = str(row.get("Page Name", row.get("Page Display Name", "")))
            display_name = str(row.get("Page Display Name", p_name))
            p_info = {
                "display_name": display_name,
                "width": int(row.get("Width", 0)) if row.get("Width") else 0,
                "height": int(row.get("Height", 0)) if row.get("Height") else 0,
                "hidden": bool(row.get("Hidden", False)),
                "visual_count": int(row.get("Visual Count", 0)) if row.get("Visual Count") else 0,
                "visuals": {},
            }
            report_data["pages"][p_name] = p_info

        for _, row in visuals_df.iterrows():
            p_name = str(row.get("Page Name", row.get("Page Display Name", "")))
            if p_name not in report_data["pages"]:
                continue

            v_name = str(row.get("Visual Name", ""))
            v_type = str(row.get("Type", ""))
            display_type = str(row.get("Display Type", v_type))

            report_data["pages"][p_name]["visuals"][v_name] = {
                "type": v_type,
                "display_type": display_type,
                "x": int(row.get("X", 0)) if row.get("X") else 0,
                "y": int(row.get("Y", 0)) if row.get("Y") else 0,
                "width": int(row.get("Width", 0)) if row.get("Width") else 0,
                "height": int(row.get("Height", 0)) if row.get("Height") else 0,
                "hidden": bool(row.get("Hidden", False)),
                "title": str(row.get("Title", "")) if row.get("Title") else "",
            }

    return report_data


def _build_tree(report_data, expanded_pages):
    """
    Build tree items from the pre-fetched report data dict.
    Only includes visuals of pages that are in ``expanded_pages``.

    Returns (options, key_map).
    """
    items = []
    for p_name in report_data["pages"]:
        p = report_data["pages"][p_name]
        is_expanded = p_name in expanded_pages
        marker = EXPANDED if is_expanded else COLLAPSED
        hidden_suffix = " (hidden)" if p["hidden"] else ""
        v_count = len(p["visuals"])
        items.append(
            (0, "page", f"{marker} {p['display_name']}{hidden_suffix}  [{v_count} visuals]", f"page:{p_name}")
        )

        if not is_expanded:
            continue

        for v_name in sorted(p["visuals"]):
            v = p["visuals"][v_name]
            label = v["display_type"] or v["type"]
            if v["title"]:
                label = f"{label}: {v['title']}"
            if v["hidden"]:
                label += " (hidden)"
            items.append((1, "visual", label, f"visual:{p_name}:{v_name}"))

    return build_tree_items(items)


def report_explorer_tab(workspace_input=None, report_input=None):
    """
    Build the Report Explorer tab widget.

    Parameters
    ----------
    workspace_input : widgets.Text
        Shared workspace text input widget (reads .value on Load).
    report_input : widgets.Text
        Shared report text input widget (reads .value on Load).

    Returns
    -------
    widgets.VBox
    """
    # -- state --
    _report_data = {}
    _key_map = {}
    _expanded = set()  # set of page names currently expanded

    # -- Load button + status row --
    load_btn = widgets.Button(
        description="Load Report",
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
        options, _key_map = _build_tree(_report_data, _expanded)
        tree.unobserve(on_select, names="value")
        tree.options = options
        if preserve_selection and preserve_selection in options:
            tree.value = preserve_selection
        else:
            tree.value = None
        tree.observe(on_select, names="value")

    # -- preview placeholder (top-right) --
    preview_label = widgets.HTML(
        value=f'<div style="font-size:12px; font-weight:600; color:{ICON_ACCENT}; '
        f'font-family:{FONT_FAMILY}; text-transform:uppercase; letter-spacing:0.5px; '
        f'margin-bottom:2px;">Preview</div>'
    )
    preview_placeholder = widgets.HTML(
        value=f'<div style="padding:16px; color:{GRAY_COLOR}; font-size:13px; '
        f'font-family:{FONT_FAMILY}; text-align:center; '
        f'font-style:italic;">Preview \u2014 coming soon</div>',
    )
    preview_box = panel_box([preview_label, preview_placeholder], flex="1", min_height="250px")

    # -- properties placeholder (bottom-right) --
    props_label = widgets.HTML(
        value=f'<div style="font-size:12px; font-weight:600; color:{ICON_ACCENT}; '
        f'font-family:{FONT_FAMILY}; text-transform:uppercase; letter-spacing:0.5px; '
        f'margin-bottom:2px;">Properties</div>'
    )
    props_placeholder = widgets.HTML(
        value=f'<div style="padding:16px; color:{GRAY_COLOR}; font-size:13px; '
        f'font-family:{FONT_FAMILY}; text-align:center; '
        f'font-style:italic;">Properties \u2014 coming soon</div>',
    )
    props_box = panel_box([props_label, props_placeholder], flex="0 0 auto", min_height="150px")

    # -- three-panel layout --
    panels = create_three_panel_layout(tree, preview_box, props_box)

    # -- tree label --
    tree_header = widgets.HTML(
        value=f'<div style="font-size:12px; font-weight:600; color:{ICON_ACCENT}; '
        f'font-family:{FONT_FAMILY}; text-transform:uppercase; letter-spacing:0.5px; '
        f'margin-bottom:2px;">Report Structure</div>'
    )

    # -- handlers --
    def on_load(_):
        nonlocal _report_data, _key_map
        _expanded.clear()
        ws = workspace_input.value.strip() if workspace_input else None
        ws = ws or None
        rpt = report_input.value.strip() if report_input else ""
        if not rpt:
            set_status(conn_status, "Enter a report name in the top bar.", "#ff3b30")
            return

        load_btn.disabled = True
        load_btn.description = "Loading…"
        set_status(conn_status, "Connecting…", GRAY_COLOR)

        try:
            _report_data = _load_report_data(report=rpt, workspace=ws)
            _refresh_tree()

            n_pages = len(_report_data["pages"])
            n_visuals = sum(len(p["visuals"]) for p in _report_data["pages"].values())
            fmt = _report_data.get("format", "")
            fmt_str = f" ({fmt})" if fmt else ""
            set_status(
                conn_status,
                f"Loaded: {n_pages} pages, {n_visuals} visuals{fmt_str}",
                "#34c759",
            )
        except Exception as e:
            set_status(conn_status, f"Error: {e}", "#ff3b30")
        finally:
            load_btn.disabled = False
            load_btn.description = "Load Report"

    def on_select(change):
        selected = change.get("new")
        if not selected or selected not in _key_map:
            return
        key = _key_map[selected]
        if key.startswith("page:"):
            p_name = key.split(":", 1)[1]
            if p_name in _expanded:
                _expanded.discard(p_name)
            else:
                _expanded.add(p_name)
            _refresh_tree(preserve_selection=selected)

    def on_expand_all(_):
        if _report_data:
            _expanded.update(_report_data["pages"].keys())
            _refresh_tree()

    def on_collapse_all(_):
        _expanded.clear()
        if _report_data:
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
