# Fix measure format — standalone BPA fixer.
# Sets format string for measures without a format to #,0.

from typing import Optional
from uuid import UUID
from sempy._utils._log import log


@log
def fix_measure_format(
    dataset: str | UUID,
    workspace: Optional[str | UUID] = None,
    scan_only: bool = False,
) -> int:
    """
    Sets the format string of measures that have no format string to #,0.

    Parameters
    ----------
    dataset : str | UUID
        Name or ID of the semantic model.
    workspace : str | uuid.UUID, default=None
        The Fabric workspace name or ID.
    scan_only : bool, default=False
        If True, only reports what would be fixed without making changes.

    Returns
    -------
    int
        Number of items fixed.
    """
    from sempy_labs.tom import connect_semantic_model

    fixed = 0
    with connect_semantic_model(dataset=dataset, readonly=scan_only, workspace=workspace) as tom:
        for table in tom.model.Tables:
            for m in table.Measures:
                fmt = str(m.FormatString) if m.FormatString else ""
                if not fmt.strip():
                    if scan_only:
                        print(f"  Would fix: [{m.Name}] → #,0")
                    else:
                        m.FormatString = "#,0"
                        print(f"  Fixed: [{m.Name}] → #,0")
                    fixed += 1

    action = "Would fix" if scan_only else "Fixed"
    print(f"  {action} {fixed} measure format(s).")
    return fixed
