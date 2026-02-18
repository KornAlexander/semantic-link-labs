from uuid import UUID
from typing import Optional
from sempy_labs._helper_functions import (
    resolve_workspace_name_and_id,
    resolve_item_name_and_id,
)
from sempy._utils._log import log
import sempy_labs._icons as icons
from sempy_labs.report._reportwrapper import connect_report


@log
def fix_pie_charts(
    report: str | UUID,
    target_visual_type: str = "clusteredBarChart",
    workspace: Optional[str | UUID] = None,
) -> None:
    """
    Replaces all pie chart visuals in a report with a specified visual type.

    This function scans through all visuals in a report and converts any pie charts
    to the target visual type (default: clustered bar chart).

    Parameters
    ----------
    report : str | uuid.UUID
        Name or ID of the report.
    target_visual_type : str, default="clusteredBarChart"
        The target visual type to replace pie charts with.
        Valid options: "clusteredBarChart", "barChart", "columnChart", etc.
    workspace : str | uuid.UUID, default=None
        The Fabric workspace name or ID.
        Defaults to None which resolves to the workspace of the attached lakehouse
        or if no lakehouse attached, resolves to the workspace of the notebook.

    Returns
    -------
    None
        This function does not return a value.
    """

    with connect_report(report=report, workspace=workspace, readonly=False, show_diffs=False) as rw:
        # Get all file paths in the report
        paths_df = rw.list_paths()
        pie_charts_replaced = 0

        for file_path in paths_df["Path"]:
            # Only process visual.json files
            if not file_path.endswith("/visual.json"):
                continue

            visual = rw.get(file_path=file_path)

            # Check if this is a pie chart
            if visual.get("visual", {}).get("visualType") == "pieChart":
                # Change the visual type to target type
                visual["visual"]["visualType"] = target_visual_type

                # Update the visual in the report
                rw.update(file_path=file_path, payload=visual)
                pie_charts_replaced += 1
                print(
                    f"{icons.green_dot} Replaced pie chart in {file_path} with {target_visual_type}"
                )

        if pie_charts_replaced == 0:
            print(
                f"{icons.info} No pie charts found in the '{rw._report_name}' report."
            )
        else:
            print(
                f"{icons.green_dot} Successfully replaced {pie_charts_replaced} pie chart(s) with {target_visual_type}."
            )
