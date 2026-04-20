import re
import sempy.fabric as fabric
from uuid import UUID
from typing import Optional
from sempy_labs.tom import connect_semantic_model
from sempy_labs._helper_functions import (
    _create_spark_session,
)
from sempy_labs.directlake._sources import get_direct_lake_sources

def normalize_filter(f):
    # Remove extra whitespace
    f = f.strip()
    
    # Replace == with =
    f = re.sub(r'==', '=', f)
    
    # Remove square brackets around column names
    f = re.sub(r'\[(.*?)\]', r'\1', f)
    
    # Convert double quotes to single quotes
    f = re.sub(r'"(.*?)"', r"'\1'", f)
    
    return f


def filter_model(dataset: str | UUID, workspace: Optional[str | UUID] = None, mini_model_name: str, filters: dict):

    """

    Example:

        filters = {
        "dimension_city": "City = 'San Isidro' ",
        "fact_sale": "SaleKey > 100 ", 
    }
    
    """

    queries = {}

    sources = get_direct_lake_sources(dataset=dataset, workspace=workspace)

    with connect_semantic_model(dataset='mastermodel', workspace=None) as tom:
        # Validation
        if any(p for p in tom.all_partitions() if str(p.Mode) != 'DirectLake'):
            print("Only DirectLake partitions are supported for filtering.")
            return

        if len(sources) > 1:
            print("Multiple DirectLake sources are not supported for filtering.")
            return
        item_id = next(s.get('itemId') for s in sources)
        if item_id != fabric.get_lakehouse_id():
            print("The Direct Lake source must be from the same lakehouse.")
            return

        # Collect relationship info
        rels = []
        for r in tom.model.Relationships:
            rels.append({
                "from_table": r.FromTable,
                "from_column": r.FromColumn,
                "to_table": r.ToTable,
                "to_column": r.ToColumn,
                "to_cardinality": r.ToCardinality,
                "from_cardinality": r.FromCardinality,
                "crossfilteringbehavior": r.CrossFilteringBehavior,
            })
    
        # Formulate queries based on filters and relationships
        for table_name, filter_value in filters.items():
            where_clause = normalize_filter(filter_value)
            schema_name, entity_name = next(
            ((p.Source.SchemaName, p.Source.EntityName) for p in tom.all_partitions() if p.Parent.Name == table_name),
            (None, None)
        )

            # If table name has spaces, quote it
            table_sql = f"`{entity_name}`" if " " in entity_name else entity_name
        
            if schema_name:
                query = f"SELECT * FROM {schema_name}.{table_sql} WHERE {where_clause}"
            else:
                query = f"SELECT * FROM {table_sql} WHERE {where_clause}"
            queries[table_name] = query

    # Build materialized views for the filtered tables
    spark = _create_spark_session()
    spark.sql(f'CREATE SCHEMA IF NOT EXISTS {mini_model_name}')
    for table_name, query in queries.items():
        name = f'{mini_model_name}.{table_name}'
        spark.sql(f"DROP MATERIALIZED LAKE VIEW IF EXISTS {name}")
        spark.sql(f"CREATE MATERIALIZED LAKE VIEW {name} AS {query}")
        