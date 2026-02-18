import json
from uuid import UUID
from typing import Optional
from sempy._utils._log import log
import sempy_labs._icons as icons
from sempy_labs._helper_functions import (
    resolve_workspace_name_and_id,
    resolve_item_name_and_id,
    _conv_b64,
    _base_api,
)
from sempy_labs.report._reportwrapper import connect_report


@log
def get_thin_model_definition(
    report: str | UUID,
    workspace: Optional[str | UUID] = None,
) -> list[dict]:
    """
    Returns the full definition of a thin report (live-connected report).

    A thin report connects to a published semantic model via a live connection.
    Its definition files contain the connection details and any report-level measures.

    This is a wrapper function for the following API: `Items - Get Item Definition <https://learn.microsoft.com/rest/api/fabric/core/items/get-item-definition>`_.

    Service Principal Authentication is supported (see `here <https://github.com/microsoft/semantic-link-labs/blob/main/notebooks/Service%20Principal.ipynb>`_ for examples).

    Parameters
    ----------
    report : str | uuid.UUID
        Name or ID of the thin report.
    workspace : str | uuid.UUID, default=None
        The Fabric workspace name or ID.
        Defaults to None which resolves to the workspace of the attached lakehouse
        or if no lakehouse attached, resolves to the workspace of the notebook.

    Returns
    -------
    list of dict
        A list of definition file parts, each with keys ``file_name`` and ``content``.
    """

    with connect_report(report=report, workspace=workspace, readonly=True, show_diffs=False) as rw:
        paths_df = rw.list_paths()
        parts = []
        
        for file_path in paths_df["Path"]:
            content = rw.get(file_path=file_path)
            parts.append({
                "file_name": file_path,
                "content": content
            })
        
        return parts


@log
def set_thin_model_perspective(
    report: str | UUID,
    perspective: Optional[str] = None,
    workspace: Optional[str | UUID] = None,
) -> None:
    """
    Changes the model connection of a thin report to target a specific perspective.

    Appends ``Cube=<perspective>`` to the report's ``powerbi://`` connection string
    in ``definition.pbir``. Pass ``None`` to remove an existing perspective.

    This is a wrapper function for the following API: `Reports - Update Report Definition <https://learn.microsoft.com/rest/api/fabric/report/items/update-report-definition>`_.

    Service Principal Authentication is supported (see `here <https://github.com/microsoft/semantic-link-labs/blob/main/notebooks/Service%20Principal.ipynb>`_ for examples).

    Parameters
    ----------
    report : str | uuid.UUID
        Name or ID of the thin report.
    perspective : str, default=None
        The name of the perspective to connect to. Pass ``None`` to connect to the full model.
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
        # Get the definition.pbir file
        definition_pbir = rw.get(file_path="definition.pbir")
        
        # Update the perspective in the connection string
        conn = definition_pbir["datasetReference"]["byConnection"]["connectionString"]
        tokens = [
            t
            for t in conn.split(";")
            if t.strip() and not t.strip().lower().startswith("cube=")
        ]
        if perspective:
            tokens.append(f"Cube={perspective}")
        definition_pbir["datasetReference"]["byConnection"]["connectionString"] = ";".join(
            tokens
        )
        
        # Update the file
        rw.update(file_path="definition.pbir", payload=definition_pbir)
        
        # Show success message
        if perspective:
            print(
                f"{icons.green_dot} The '{rw._report_name}' report now connects to the '{perspective}' perspective."
            )
        else:
            print(
                f"{icons.green_dot} The '{rw._report_name}' report now connects to the full model."
            )
