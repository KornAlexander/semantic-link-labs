# Fix measure descriptions — standalone BPA fixer.
# Sets measure description to its DAX expression when description is empty.

from typing import Optional
from uuid import UUID
from sempy._utils._log import log


@log
def fix_measure_descriptions(
    dataset: str | UUID,
    workspace: Optional[str | UUID] = None,
    scan_only: bool = False,
) -> int:
    """
    Sets the description of visible measures (that have no description) to their DAX expression.

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
                if not m.IsHidden and (not m.Description or str(m.Description).strip() == ""):
                    expr = str(m.Expression) if m.Expression else ""
                    if not expr:
                        continue
                    if scan_only:
                        print(f"  Would fix: [{m.Name}] — set description to DAX expression")
                    else:
                        m.Description = expr
                        print(f"  Fixed: [{m.Name}] — description set to DAX expression")
                    fixed += 1

    action = "Would fix" if scan_only else "Fixed"
    print(f"  {action} {fixed} measure description(s).")
    return fixed
