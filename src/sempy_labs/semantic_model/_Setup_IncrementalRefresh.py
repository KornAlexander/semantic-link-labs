# Setup Incremental Refresh — standalone fixer.
# Auto-detects the first DateTime column and configures incremental refresh.

from typing import Optional
from uuid import UUID
from datetime import datetime, timedelta
import re


def setup_incremental_refresh(
    dataset: str,
    table_name: str,
    workspace: Optional[str | UUID] = None,
    column_name: Optional[str] = None,
    rolling_window_years: int = 3,
    incremental_days: int = 30,
    only_refresh_complete_days: bool = False,
    scan_only: bool = False,
):
    """
    Sets up an incremental refresh policy on a table.

    Auto-detects the first DateTime column if column_name is not specified.

    Parameters
    ----------
    dataset : str
        Name of the semantic model.
    table_name : str
        Name of the table to configure.
    workspace : str | uuid.UUID, default=None
        The Fabric workspace name or ID.
    column_name : str, default=None
        The DateTime column to use. If None, auto-detects the first DateTime column.
    rolling_window_years : int, default=3
        Number of years to keep in the rolling window.
    incremental_days : int, default=30
        Number of days to refresh incrementally.
    only_refresh_complete_days : bool, default=False
        If True, only refresh complete days (offset -1).
    scan_only : bool, default=False
        If True, only reports what would be configured without making changes.
    """
    from sempy_labs.tom import connect_semantic_model

    with connect_semantic_model(dataset=dataset, readonly=scan_only, workspace=workspace) as tom:
        import Microsoft.AnalysisServices.Tabular as TOM

        t = tom.model.Tables[table_name]

        # Auto-detect date column
        date_col = column_name
        if date_col is None:
            for col in t.Columns:
                if col.DataType == TOM.DataType.DateTime:
                    date_col = col.Name
                    break
            if date_col is None:
                print(f"  No DateTime column found in '{table_name}'. Cannot set up incremental refresh.")
                return

        # Verify column type
        c = t.Columns[date_col]
        if c.DataType != TOM.DataType.DateTime:
            print(f"  Column '{date_col}' is not DateTime ({c.DataType}). Skipping.")
            return

        # Check if already has refresh policy
        try:
            if t.RefreshPolicy is not None:
                print(f"  Table '{table_name}' already has an incremental refresh policy. Skipping.")
                return
        except Exception:
            pass

        # Dates
        end_date = datetime.now().strftime("%m/%d/%Y")
        start_date = (datetime.now() - timedelta(days=rolling_window_years * 365)).strftime("%m/%d/%Y")

        if scan_only:
            print(f"  Would set up incremental refresh on '{table_name}':")
            print(f"    Column: [{date_col}]")
            print(f"    Rolling window: {rolling_window_years} year(s)")
            print(f"    Incremental refresh: {incremental_days} day(s)")
            print(f"    Only complete days: {only_refresh_complete_days}")
            print(f"    Range: {start_date} - {end_date}")
            return

        import System

        # Inline incremental refresh setup (bypasses tom.add_incremental_refresh_policy
        # which has a bug using p.Expression instead of p.Source.Expression)
        for idx, p in enumerate(t.Partitions):
            if p.SourceType != TOM.PartitionSourceType.M:
                print(f"  Partition '{p.Name}' is not M-partition ({p.SourceType}). Skipping table.")
                return
            if idx == 0:
                text = p.Source.Expression.rstrip()
                pattern = r"in\s*[^ ]*"
                matches = list(re.finditer(pattern, text))
                if not matches:
                    print(f"  Could not parse M-partition expression for '{table_name}'. Skipping.")
                    return
                last_match = matches[-1]
                text_before = text[:last_match.start()]
                obj = text[text.rfind(" ") + 1:]
                end_expr = (
                    f'#"Filtered Rows IR" = Table.SelectRows({obj}, '
                    f'each [{date_col}] >= RangeStart and [{date_col}] <= RangeEnd)\n'
                    f'#"Filtered Rows IR"'
                )
                p.Source.Expression = text_before + end_expr

        # Add RangeStart / RangeEnd expressions
        date_fmt = "%m/%d/%Y"
        ds = datetime.strptime(start_date, date_fmt)
        de = datetime.strptime(end_date, date_fmt)
        tom.add_expression(
            name="RangeStart",
            expression=f'datetime({ds.year}, {ds.month}, {ds.day}, 0, 0, 0) meta [IsParameterQuery=true, Type="DateTime", IsParameterQueryRequired=true]',
        )
        tom.add_expression(
            name="RangeEnd",
            expression=f'datetime({de.year}, {de.month}, {de.day}, 0, 0, 0) meta [IsParameterQuery=true, Type="DateTime", IsParameterQueryRequired=true]',
        )

        # Set refresh policy
        rp = TOM.BasicRefreshPolicy()
        rp.IncrementalPeriods = incremental_days
        rp.IncrementalGranularity = System.Enum.Parse(TOM.RefreshGranularityType, "Day")
        rp.RollingWindowPeriods = rolling_window_years
        rp.RollingWindowGranularity = System.Enum.Parse(TOM.RefreshGranularityType, "Year")
        rp.SourceExpression = t.Partitions[0].Source.Expression
        if only_refresh_complete_days:
            rp.IncrementalPeriodsOffset = -1
        t.RefreshPolicy = rp

        print(f"  \u2713 Incremental refresh configured on '{table_name}'[{date_col}]")
        print(f"    Rolling window: {rolling_window_years} year(s), Refresh: {incremental_days} day(s)")
