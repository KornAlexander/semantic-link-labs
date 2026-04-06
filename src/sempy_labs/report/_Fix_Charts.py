# Unified Chart Fixer — covers bar, column, line, and combo charts.
# Each chart type has its own checks config (axis semantics differ).
# Line charts keep the Y value axis visible.

from uuid import UUID
from typing import Optional
from sempy._utils._log import log
import sempy_labs._icons as icons
from sempy_labs.report._reportwrapper import connect_report


def _get_visual_property(visual: dict, object_name: str, property_name: str) -> str | None:
    obj_list = visual.get("visual", {}).get("objects", {}).get(object_name, [])
    if not obj_list:
        return None
    return obj_list[0].get("properties", {}).get(property_name, {}).get("expr", {}).get("Literal", {}).get("Value")


def _set_visual_property(visual: dict, object_name: str, property_name: str, value: str) -> None:
    objects = visual.setdefault("visual", {}).setdefault("objects", {})
    if object_name not in objects or not objects[object_name]:
        objects[object_name] = [{"properties": {}}]
    obj = objects[object_name][0]
    if "properties" not in obj:
        obj["properties"] = {}
    obj["properties"][property_name] = {"expr": {"Literal": {"Value": value}}}


# --- Per-type check configs ---
# Bar charts: value axis = horizontal (X), category axis = vertical (Y)
_BAR_CHECKS = [
    ("valueAxis",    "showAxisTitle", "false", "X axis title"),
    ("valueAxis",    "show",          "false", "X axis values"),
    ("categoryAxis", "showAxisTitle", "false", "Y axis title"),
    ("labels",       "show",          "true",  "Data labels"),
    ("valueAxis",    "gridlineShow",  "false", "Gridlines"),
]

# Column charts: category axis = horizontal (X), value axis = vertical (Y)
_COLUMN_CHECKS = [
    ("categoryAxis", "showAxisTitle", "false", "X axis title"),
    ("valueAxis",    "showAxisTitle", "false", "Y axis title"),
    ("valueAxis",    "show",          "false", "Y axis values"),
    ("labels",       "show",          "true",  "Data labels"),
    ("categoryAxis", "gridlineShow",  "false", "Gridlines"),
]

# Line charts: same layout as column but KEEP Y value axis visible
_LINE_CHECKS = [
    ("categoryAxis", "showAxisTitle", "false", "X axis title"),
    ("valueAxis",    "showAxisTitle", "false", "Y axis title"),
    ("labels",       "show",          "true",  "Data labels"),
    ("categoryAxis", "gridlineShow",  "false", "Gridlines"),
]

_TYPE_CHECKS = {
    "barChart":                        _BAR_CHECKS,
    "clusteredBarChart":               _BAR_CHECKS,
    "columnChart":                     _COLUMN_CHECKS,
    "clusteredColumnChart":            _COLUMN_CHECKS,
    "lineChart":                       _LINE_CHECKS,
    "lineClusteredColumnComboChart":   _LINE_CHECKS,
}

_ALL_CHART_TYPES = set(_TYPE_CHECKS.keys())

_TYPE_LABELS = {
    "barChart":                        "bar chart",
    "clusteredBarChart":               "clustered bar chart",
    "columnChart":                     "column chart",
    "clusteredColumnChart":            "clustered column chart",
    "lineChart":                       "line chart",
    "lineClusteredColumnComboChart":   "line/column combo chart",
}


@log
def fix_charts(
    report: str | UUID,
    page_name: Optional[str] = None,
    workspace: Optional[str | UUID] = None,
    scan_only: bool = False,
    chart_types: Optional[set[str]] = None,
) -> None:
    """
    Fixes chart visuals in a report by applying best-practice formatting.

    Covers bar, column, line, and combo chart types. Applies per-type rules:
    - Bar/Column: remove axis titles, remove value axis labels, add data labels, remove gridlines
    - Line/Combo: same but keeps the Y value axis visible

    Parameters
    ----------
    report : str | uuid.UUID
        Name or ID of the report.
    page_name : str, default=None
        The display name of the page to apply changes to.
        Defaults to None which applies changes to all pages.
    workspace : str | uuid.UUID, default=None
        The Fabric workspace name or ID.
        Defaults to None which resolves to the workspace of the attached lakehouse
        or if no lakehouse attached, resolves to the workspace of the notebook.
    scan_only : bool, default=False
        If True, only scans and reports issues without applying fixes.
    chart_types : set[str], default=None
        Subset of visual types to process. Defaults to all supported types.
    """

    target_types = (chart_types or _ALL_CHART_TYPES) & _ALL_CHART_TYPES

    with connect_report(report=report, workspace=workspace, readonly=scan_only, show_diffs=False) as rw:
        if rw.format != "PBIR":
            print(
                f"{icons.red_dot} Report '{rw._report_name}' is in '{rw.format}' format, not PBIR. "
                f"Run 'Upgrade to PBIR' first."
            )
            return

        paths_df = rw.list_paths()
        charts_found = 0
        charts_fixed = 0
        charts_need_fixing = 0

        page_id = rw.resolve_page_name(page_name) if page_name else None

        for file_path in paths_df["Path"]:
            if not file_path.endswith("/visual.json"):
                continue
            if page_id and f"/{page_id}/" not in file_path:
                continue

            visual = rw.get(file_path=file_path)
            vtype = visual.get("visual", {}).get("visualType")

            if vtype not in target_types:
                continue

            charts_found += 1
            checks = _TYPE_CHECKS[vtype]

            issues = [
                label
                for obj, prop, val, label in checks
                if _get_visual_property(visual, obj, prop) != val
            ]

            if not issues:
                if scan_only:
                    print(f"{icons.green_dot} {file_path} — {_TYPE_LABELS[vtype]} OK")
                continue

            charts_need_fixing += 1
            type_label = _TYPE_LABELS[vtype]

            if scan_only:
                print(f"{icons.yellow_dot} {file_path} — {type_label} needs fixing: {', '.join(issues)}")
                continue

            for obj, prop, val, _ in checks:
                _set_visual_property(visual, obj, prop, val)

            rw.update(file_path=file_path, payload=visual)
            charts_fixed += 1
            print(f"{icons.green_dot} Fixed {type_label} in {file_path}")

        type_desc = "chart" if len(target_types) > 2 else "/".join(sorted({_TYPE_LABELS[t] for t in target_types}))

        if charts_found == 0:
            print(f"{icons.info} No {type_desc}s found in the '{rw._report_name}' report.")
        elif scan_only:
            if charts_need_fixing == 0:
                print(f"\n{icons.green_dot} Scanned {charts_found} {type_desc}(s) — all have correct settings.")
            else:
                print(f"\n{icons.yellow_dot} Scanned {charts_found} {type_desc}(s) — {charts_need_fixing} need fixing.")
        elif charts_fixed == 0:
            print(f"{icons.info} Found {charts_found} {type_desc}(s) — all already have correct settings.")
        else:
            print(f"{icons.green_dot} Successfully fixed {charts_fixed} of {charts_found} {type_desc}(s).")


# Convenience wrappers for backward compatibility
def fix_barcharts(
    report: str | UUID,
    page_name: Optional[str] = None,
    workspace: Optional[str | UUID] = None,
    scan_only: bool = False,
) -> None:
    """Fixes bar chart visuals. Wrapper around fix_charts(chart_types=bar types)."""
    fix_charts(report=report, page_name=page_name, workspace=workspace, scan_only=scan_only,
               chart_types={"barChart", "clusteredBarChart"})


def fix_columncharts(
    report: str | UUID,
    page_name: Optional[str] = None,
    workspace: Optional[str | UUID] = None,
    scan_only: bool = False,
) -> None:
    """Fixes column chart visuals. Wrapper around fix_charts(chart_types=column types)."""
    fix_charts(report=report, page_name=page_name, workspace=workspace, scan_only=scan_only,
               chart_types={"columnChart", "clusteredColumnChart"})


def fix_linecharts(
    report: str | UUID,
    page_name: Optional[str] = None,
    workspace: Optional[str | UUID] = None,
    scan_only: bool = False,
) -> None:
    """Fixes line chart visuals. Keeps Y value axis visible. Wrapper around fix_charts(chart_types=line types)."""
    fix_charts(report=report, page_name=page_name, workspace=workspace, scan_only=scan_only,
               chart_types={"lineChart", "lineClusteredColumnComboChart"})
