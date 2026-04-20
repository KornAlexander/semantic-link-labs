import re
import sempy.fabric as fabric
from uuid import UUID
from typing import Optional
from sempy_labs.tom import connect_semantic_model
from sempy_labs._helper_functions import (
    _create_spark_session,
)
from sempy_labs.directlake._sources import get_direct_lake_sources
import sempy_labs._icons as icons


def normalize_filter(f, alias=None):
    """Normalize a user-supplied filter expression.

    When ``alias`` is provided, bracketed column references ``[Col]`` are
    rewritten as ```alias`.`Col``` so the filter is unambiguous in a
    multi-table (joined) query. When no alias is given the brackets are
    simply stripped to preserve the original single-table behavior.
    """

    # Remove extra whitespace
    f = f.strip()

    # Replace == with =
    f = re.sub(r"==", "=", f)

    # Square bracketed column references
    if alias:
        f = re.sub(r"\[(.*?)\]", rf"`{alias}`.`\1`", f)
    else:
        f = re.sub(r"\[(.*?)\]", r"\1", f)

    # Convert double quotes to single quotes
    f = re.sub(r'"(.*?)"', r"'\1'", f)

    return f


def _table_ref(schema_name, entity_name):
    """Return a fully-qualified, safely quoted table reference."""

    entity_sql = f"`{entity_name}`" if " " in entity_name else entity_name
    if schema_name:
        return f"{schema_name}.{entity_sql}"
    return entity_sql


def create_mlvs_based_on_filters(
    dataset: str | UUID,
    mini_model_name: str,
    filters: dict,
    workspace: Optional[str | UUID] = None,
):
    """Create materialized lake views for a filtered subset of a Direct Lake
    semantic model.

    For each table listed in ``filters`` a materialized lake view is
    created (or replaced) in the ``mini_model_name`` schema with the
    supplied filter applied. In addition, filters are propagated through
    the model's many-to-one single-direction relationships: any table on
    the "many" side of such a relationship will have the filters of its
    related "one" side tables applied via the appropriate joins. This
    propagation follows relationship chains transitively.

    Example
    -------
    ::

        filters = {
            "Customer": "City = 'San Isidro'",
            "Sales":    "SaleKey > 100",
        }

    If ``Sales`` has a many-to-one OneDirection relationship to
    ``Customer``, the resulting ``Sales`` materialized view will include
    both the ``SaleKey > 100`` predicate and a join to ``Customer``
    constrained to ``City = 'San Isidro'``.
    """

    queries = {}

    sources = get_direct_lake_sources(dataset=dataset, workspace=workspace)

    with connect_semantic_model(dataset=dataset, workspace=workspace) as tom:
        # Validation
        if any(p for p in tom.all_partitions() if str(p.Mode) != "DirectLake"):
            print(
                f"{icons.red_dot} Only DirectLake partitions are supported for filtering."
            )
            return

        if len(sources) > 1:
            print("Multiple DirectLake sources are not supported for filtering.")
            return
        item_id = next(s.get("itemId") for s in sources)
        if item_id != fabric.get_lakehouse_id():
            print("The Direct Lake source must be from the same lakehouse.")
            return

        # Map of table_name -> (schema_name, entity_name)
        table_sources = {}
        for p in tom.all_partitions():
            tn = p.Parent.Name
            if tn not in table_sources:
                table_sources[tn] = (p.Source.SchemaName, p.Source.EntityName)

        # Build directed edges representing filter propagation via
        # many-to-one OneDirection relationships. A filter on the "one"
        # side flows to the "many" side, so we store edges as
        # many_table -> [(one_table, many_column, one_column), ...].
        m2o_edges: dict[str, list[tuple[str, str, str]]] = {}
        for r in tom.model.Relationships:
            if str(r.CrossFilteringBehavior) != "OneDirection":
                continue
            from_card = str(r.FromCardinality)
            to_card = str(r.ToCardinality)
            if from_card == "Many" and to_card == "One":
                many_t, many_c = r.FromTable.Name, r.FromColumn.Name
                one_t, one_c = r.ToTable.Name, r.ToColumn.Name
            elif from_card == "One" and to_card == "Many":
                many_t, many_c = r.ToTable.Name, r.ToColumn.Name
                one_t, one_c = r.FromTable.Name, r.FromColumn.Name
            else:
                continue
            m2o_edges.setdefault(many_t, []).append((one_t, many_c, one_c))

        # Determine, for every table in the model, which filtered tables
        # are reachable via m2o edges (and the join tree leading there).
        all_tables = {t.Name for t in tom.model.Tables}
        filtered_tables = set(filters.keys())

        for base_table in all_tables:
            # BFS from base_table through m2o edges, recording a single
            # parent per discovered table so we can reconstruct a tree.
            parents: dict[str, Optional[tuple[str, str, str]]] = {base_table: None}
            queue = [base_table]
            while queue:
                node = queue.pop(0)
                for one_t, many_c, one_c in m2o_edges.get(node, []):
                    if one_t not in parents:
                        parents[one_t] = (node, many_c, one_c)
                        queue.append(one_t)

            reachable_filtered = [
                t for t in parents if t in filtered_tables and t != base_table
            ]
            has_own_filter = base_table in filtered_tables

            if not has_own_filter and not reachable_filtered:
                continue

            if base_table not in table_sources:
                # Table not backed by a Direct Lake partition; skip.
                continue

            # Walk back from each reachable filtered table to collect the
            # joins we need. Keyed by the joined (child) table so each
            # intermediate is joined at most once.
            joins: dict[str, tuple[str, str, str]] = {}
            for ft in reachable_filtered:
                node = ft
                while node != base_table and node not in joins:
                    parent_info = parents[node]
                    if parent_info is None:
                        break
                    joins[node] = parent_info
                    node = parent_info[0]

            # Assemble the SQL.
            base_schema, base_entity = table_sources[base_table]
            sql = (
                f"SELECT `{base_table}`.* FROM "
                f"{_table_ref(base_schema, base_entity)} AS `{base_table}`"
            )

            # Emit joins in topological order (parents before children)
            # so every alias is in scope when referenced.
            emitted: set[str] = set()
            remaining = dict(joins)
            while remaining:
                progress = False
                for child, (parent, many_c, one_c) in list(remaining.items()):
                    if parent == base_table or parent in emitted:
                        if child not in table_sources:
                            # Should not happen for direct-lake models,
                            # but guard regardless.
                            remaining.pop(child)
                            continue
                        child_schema, child_entity = table_sources[child]
                        sql += (
                            f" INNER JOIN {_table_ref(child_schema, child_entity)} "
                            f"AS `{child}` ON `{parent}`.`{many_c}` = "
                            f"`{child}`.`{one_c}`"
                        )
                        emitted.add(child)
                        remaining.pop(child)
                        progress = True
                if not progress:
                    # Defensive: shouldn't happen with a valid tree.
                    break

            where_parts = []
            if has_own_filter:
                where_parts.append(
                    f"({normalize_filter(filters[base_table], alias=base_table)})"
                )
            for ft in reachable_filtered:
                where_parts.append(f"({normalize_filter(filters[ft], alias=ft)})")

            sql += " WHERE " + " AND ".join(where_parts)
            queries[base_table] = sql

    # Build materialized views for the filtered (and propagated) tables
    spark = _create_spark_session()
    spark.sql(f"CREATE SCHEMA IF NOT EXISTS {mini_model_name}")
    for table_name, query in queries.items():
        name = f"{mini_model_name}.{table_name}"
        print(
            f"{icons.in_progress} Creating the '{name}' materialized lake view for the '{table_name}'."
        )
        spark.sql(f"DROP MATERIALIZED LAKE VIEW IF EXISTS {name}")
        spark.sql(f"CREATE MATERIALIZED LAKE VIEW {name} AS {query}")
        print(f"{icons.green_dot} Created the '{name}' materialized view.")
