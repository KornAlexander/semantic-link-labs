# Report Explorer tab for PBI Fixer.
# Provides a tree view of report pages and visuals.

import ipywidgets as widgets
from typing import Optional
from uuid import UUID

from sempy_labs._ui_components import (
    FONT_FAMILY,
    BORDER_COLOR,
    GRAY_COLOR,
    ICON_ACCENT,
    ICONS,
    build_tree_items,
    create_three_panel_layout,
    create_connection_bar,
    input_label,
    status_html,
    set_status,
    placeholder_panel,
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


def _build_tree(report_data):
    """
    Build tree items from the pre-fetched report data dict.

    Returns (options, key_map)  where key_map values encode the node type
    as  "type:page:visual"  strings.
    """
    items = []
    for p_name in report_data["pages"]:
        p = report_data["pages"][p_name]
        hidden_suffix = " (hidden)" if p["hidden"] else ""
        v_count = len(p["visuals"])
        items.append(
            (0, "page", f"{p['display_name']}{hidden_suffix}  [{v_count} visuals]", f"page:{p_name}")
        )

        for v_name in sorted(p["visuals"]):
            v = p["visuals"][v_name]
            label = v["display_type"] or v["type"]
            if v["title"]:
                label = f"{label}: {v['title']}"
            if v["hidden"]:
                label += " (hidden)"
            items.append((1, "visual", label, f"visual:{p_name}:{v_name}"))

    return build_tree_items(items)


def report_explorer_tab(
    workspace: Optional[str | UUID] = None,
    report: Optional[str | UUID] = None,
):
    """
    Build the Report Explorer tab widget.

    Parameters
    ----------
    workspace : str | uuid.UUID, default=None
        Pre-populated workspace name or ID.
    report : str | uuid.UUID, default=None
        Pre-populated report name or ID.

    Returns
    -------
    widgets.VBox
    """
    # -- state --
    _report_data = {}
    _key_map = {}

    # -- connection bar --
    ws_input = widgets.Text(
        value=str(workspace) if workspace else "",
        placeholder="Leave empty for notebook workspace",
        layout=widgets.Layout(width="220px"),
    )
    rpt_input = widgets.Text(
        value=str(report) if report else "",
        placeholder="Report name or ID",
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
        input_label("Report"),
        rpt_input,
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

    # -- preview placeholder (top-right) --
    preview_label = widgets.HTML(
        value=f'<div style="font-size:12px; font-weight:600; color:{ICON_ACCENT}; '
        f'font-family:{FONT_FAMILY}; text-transform:uppercase; letter-spacing:0.5px; '
        f'margin-bottom:2px;">Preview</div>'
    )
    preview_placeholder = placeholder_panel("Preview — coming soon", min_height="250px")
    preview_box = widgets.VBox(
        [preview_label, preview_placeholder],
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
        f'margin-bottom:2px;">Report Structure</div>'
    )

    # -- handlers --
    def on_load(_):
        nonlocal _report_data, _key_map
        ws = ws_input.value.strip() or None
        rpt = rpt_input.value.strip()
        if not rpt:
            set_status(conn_status, "Enter a report name or ID.", "#ff3b30")
            return

        load_btn.disabled = True
        load_btn.description = "Loading…"
        set_status(conn_status, "Connecting…", GRAY_COLOR)

        try:
            _report_data = _load_report_data(report=rpt, workspace=ws)
            options, _key_map = _build_tree(_report_data)
            tree.options = options
            tree.value = None

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
            load_btn.description = "Load"

    load_btn.on_click(on_load)

    # -- assemble --
    return widgets.VBox(
        [conn_bar, tree_header, panels],
        layout=widgets.Layout(
            padding="12px",
            gap="4px",
        ),
    )
