import json
from uuid import UUID
from typing import Optional
from sempy._utils._log import log
import sempy_labs._icons as icons
from sempy_labs._helper_functions import (
    resolve_workspace_name_and_id,
    resolve_item_name_and_id,
    _decode_b64,
    _conv_b64,
    _base_api,
)


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
        A list of definition file parts, each with keys ``file_name``, ``content``, and ``is_binary``.
        Binary files have ``is_binary=True`` and content is truncated or marked to avoid large payloads.
    """

    workspace_name, workspace_id = resolve_workspace_name_and_id(workspace)
    _, report_id = resolve_item_name_and_id(
        item=report, type="Report", workspace=workspace_id
    )
    result = _base_api(
        request=f"/v1/workspaces/{workspace_id}/items/{report_id}/getDefinition",
        method="post",
        lro_return_json=True,
        status_codes=None,
        client="fabric_sp",
    )
    
    parts = []
    for p in result["definition"]["parts"]:
        try:
            content = _decode_b64(p["payload"])
            parts.append({
                "file_name": p["path"],
                "content": content,
                "is_binary": False
            })
        except UnicodeDecodeError:
            # Binary file - don't include full base64 payload
            parts.append({
                "file_name": p["path"],
                "content": f"[Binary file - {len(p['payload']) // 1024}KB base64 payload]",
                "is_binary": True,
                "payload_size_kb": len(p["payload"]) // 1024
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

    workspace_name, workspace_id = resolve_workspace_name_and_id(workspace)
    report_name, report_id = resolve_item_name_and_id(
        item=report, type="Report", workspace=workspace_id
    )

    result = _base_api(
        request=f"/v1/workspaces/{workspace_id}/items/{report_id}/getDefinition",
        method="post",
        lro_return_json=True,
        status_codes=None,
        client="fabric_sp",
    )

    new_parts = []
    for part in result["definition"]["parts"]:
        if part["path"] == "definition.pbir":
            content = json.loads(_decode_b64(part["payload"]))
            conn = content["datasetReference"]["byConnection"]["connectionString"]
            tokens = [
                t
                for t in conn.split(";")
                if t.strip() and not t.strip().lower().startswith("cube=")
            ]
            if perspective:
                tokens.append(f"Cube={perspective}")
            content["datasetReference"]["byConnection"]["connectionString"] = ";".join(
                tokens
            )
            new_parts.append(
                {
                    "path": part["path"],
                    "payload": _conv_b64(content),
                    "payloadType": "InlineBase64",
                }
            )
        else:
            new_parts.append(
                {
                    "path": part["path"],
                    "payload": part["payload"],
                    "payloadType": "InlineBase64",
                }
            )

    _base_api(
        request=f"/v1/workspaces/{workspace_id}/reports/{report_id}/updateDefinition",
        method="post",
        payload={"definition": {"parts": new_parts}},
        lro_return_status_code=True,
        status_codes=None,
        client="fabric_sp",
    )

    if perspective:
        print(
            f"{icons.green_dot} The '{report_name}' report now connects to the '{perspective}' perspective."
        )
    else:
        print(
            f"{icons.green_dot} The '{report_name}' report now connects to the full model."
        )
