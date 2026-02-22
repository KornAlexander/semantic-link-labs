# Semantic Link Labs — Agent Memory

## Project Basics
- **Package:** `semantic-link-labs` v0.13.1
- **Python:** `>=3.10,<3.12` — 3.12 is NOT supported (hard upper bound)
- **Source layout:** `src/sempy_labs/`
- **Main exports file:** `src/sempy_labs/__init__.py`

## Critical Patterns

### All public functions MUST have:
- `@log` decorator from `from sempy._utils._log import log`
- numpydoc-style docstring with Parameters + Returns sections

### `_base_api` return type matrix (common mistake area):
- Default → `Response` object → **must call `.json()`** to get dict
- `uses_pagination=True` → `list[dict]` → iterate directly, each item has `.get("value", [])`
- `lro_return_json=True` → `dict` → use directly
- `lro_return_status_code=True` → `int`

### Standard patterns:
```python
(workspace_name, workspace_id) = resolve_workspace_name_and_id(workspace)
df = _create_dataframe(columns={"Id": "string", "Name": "string"})
```

## Tooling
- **Formatter:** `black src/sempy_labs tests` (88 char default line length)
- **Linter:** `flake8 src/sempy_labs tests` (max-line-length = 200 in pyproject.toml)
- **Type checker:** `mypy src/sempy_labs`
- **Tests:** `pytest -s tests/` — NOTE: tests/ directory is currently **empty** (no test files committed)
- **Docs:** `sphinx-apidoc -f -o docs/source src/sempy_labs/ && cd docs && make html`

## Key Files
- `src/sempy_labs/_helper_functions.py` — ~2860 lines, ~100 utility functions (central hub)
- `src/sempy_labs/_icons.py` — icons, constants, SKU mappings, item type dicts
- `src/sempy_labs/_utils.py` — `item_types` and `items` dicts (~40 Fabric item types)
- `src/sempy_labs/tom/_model.py` — TOMWrapper class (full TOM implementation)

## Submodules in src/sempy_labs/
Core: `tom/`, `admin/`, `report/`, `lakehouse/`, `directlake/`, `migration/`
Item-specific (each with `__init__.py`): `semantic_model/`, `workspace/`, `warehouse/`,
`notebook/`, `lakehouse/`, `environment/`, `dataflow/`, `data_pipeline/`, `deployment_pipeline/`,
`gateway/`, `git/`, `graph/`, `eventhouse/`, `eventstream/`, `kql_database/`, `kql_queryset/`,
`kql_dashboard/`, `graphql/`, `ml_experiment/`, `ml_model/`, `spark/`, `sql_endpoint/`,
`sql_database/`, `mirrored_database/`, `mirrored_warehouse/`, `mirrored_azure_databricks_catalog/`,
`managed_private_endpoint/`, `rti/`, `connection/`, `daxlib/`, `theme/`, `variable_library/`,
`warehouse_snapshot/`, `operations_agent/`, `apache_airflow_job/`, `event_schema_set/`,
`external_data_share/`, `mounted_data_factory/`, `snowflake_database/`, `surge_protection/`,
`graph_model/`, `ml_model/`

## Skills (always invoke before coding)
See `.claude/skills/` — use `Skill` tool to invoke. Key ones:
- `add-function` — before adding any new function
- `rest-api-patterns` — before any API wrapper work
- `planning-with-files` — for complex/multi-step tasks (>5 tool calls)
- `code-style` — before running black/flake8
- `write-tests` / `run-tests` — for test work
- `tom-operations` — for TOM/semantic model changes
- `build-docs` — for documentation work

## Details file
See `memory/repo-details.md` for extended notes.
