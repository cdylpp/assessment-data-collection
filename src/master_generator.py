from __future__ import annotations

import argparse
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Any, DefaultDict, Dict, Iterable, List, Mapping, MutableMapping, Optional, Sequence, Tuple

from openpyxl import Workbook, load_workbook

from template_generator import (
    DEFAULT_CONFIG_PATH,
    HEADER_FILL,
    HEADER_FONT,
    LoadedTemplateConfig,
    TemplateGenerationRequest,
    load_generation_inputs,
    load_yaml,
    resolve_path,
)


@dataclass(frozen=True)
class MasterGenerationRequest:
    config_path: Path
    workbook_paths: Sequence[Path]
    output_path: Optional[Path] = None


@dataclass(frozen=True)
class LoadedMasterConfig:
    template: LoadedTemplateConfig
    master_path: Path
    master_doc: Dict[str, Any]


@dataclass
class CandidateAccumulator:
    source_values: DefaultDict[Tuple[str, str], List[Any]] = field(
        default_factory=lambda: defaultdict(list)
    )
    metric_values: DefaultDict[str, List[Any]] = field(
        default_factory=lambda: defaultdict(list)
    )


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Aggregate one or more node workbooks into the master workbook."
    )
    parser.add_argument(
        "--config",
        default=DEFAULT_CONFIG_PATH,
        help="Path to canonical config file.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output xlsx path (default: workbooks/MasterWorkbook_v{version}.xlsx).",
    )
    parser.add_argument(
        "workbooks",
        nargs="+",
        help="One or more node workbook xlsx files to aggregate.",
    )
    return parser.parse_args(argv)


def request_from_namespace(args: argparse.Namespace) -> MasterGenerationRequest:
    return MasterGenerationRequest(
        config_path=Path(args.config).resolve(),
        workbook_paths=[Path(value).resolve() for value in args.workbooks],
        output_path=Path(args.output).resolve() if args.output else None,
    )


def validation_contract(config_doc: Mapping[str, Any]) -> Mapping[str, Any]:
    validation = config_doc.get("validation", {})
    if not isinstance(validation, dict):
        raise ValueError("config.yaml 'validation' must be a mapping")
    return validation


def validate_supported_value(
    *,
    label: str,
    value: Any,
    supported_values: Iterable[Any],
) -> None:
    if value not in set(supported_values):
        raise ValueError(
            "{0} uses unsupported value '{1}'".format(label, value)
        )


def validate_duplicate_rule(
    *,
    label: str,
    rule: Mapping[str, Any],
    validation: Mapping[str, Any],
) -> None:
    mode = rule.get("mode", "error_on_conflict")
    validate_supported_value(
        label="{0}.mode".format(label),
        value=mode,
        supported_values=validation.get("supported_duplicate_resolution_modes", []),
    )
    if mode == "aggregate":
        aggregate_function = rule.get("aggregate_function")
        if not aggregate_function:
            raise ValueError(
                "{0} requires aggregate_function when mode is 'aggregate'".format(label)
            )
        validate_supported_value(
            label="{0}.aggregate_function".format(label),
            value=aggregate_function,
            supported_values=validation.get("supported_aggregate_functions", []),
        )


def validate_null_rule(
    *,
    label: str,
    rule: Mapping[str, Any],
    validation: Mapping[str, Any],
) -> None:
    output_value = rule.get("on_all_inputs_null", "null")
    validate_supported_value(
        label="{0}.on_all_inputs_null".format(label),
        value=output_value,
        supported_values=validation.get("supported_null_outputs", []),
    )


def validate_master_config(loaded: LoadedMasterConfig) -> None:
    validation = validation_contract(loaded.template.config_doc)
    master_doc = loaded.master_doc

    columns = master_doc.get("columns")
    if not isinstance(columns, list) or not columns:
        raise ValueError("master-config.yaml must contain a non-empty 'columns' list")

    defaults = master_doc.get("defaults", {})
    if not isinstance(defaults, dict):
        raise ValueError("master-config.yaml 'defaults' must be a mapping")

    default_duplicate_rule = defaults.get("duplicate_metric_rule", {})
    if not isinstance(default_duplicate_rule, dict):
        raise ValueError(
            "master-config.yaml 'defaults.duplicate_metric_rule' must be a mapping"
        )
    validate_duplicate_rule(
        label="defaults.duplicate_metric_rule",
        rule=default_duplicate_rule,
        validation=validation,
    )

    default_null_rule = defaults.get("null_rule", {})
    if not isinstance(default_null_rule, dict):
        raise ValueError("master-config.yaml 'defaults.null_rule' must be a mapping")
    validate_null_rule(
        label="defaults.null_rule",
        rule=default_null_rule,
        validation=validation,
    )

    seen_column_ids = set()
    for column in columns:
        if not isinstance(column, dict):
            raise ValueError("Each master-config column must be a mapping")

        column_id = column.get("column_id")
        if not column_id:
            raise ValueError("Master column missing column_id")
        if column_id in seen_column_ids:
            raise ValueError("Duplicate master column_id: {0}".format(column_id))
        seen_column_ids.add(column_id)

        source = column.get("source")
        if not isinstance(source, dict):
            raise ValueError(
                "Master column '{0}' is missing a valid source mapping".format(column_id)
            )

        transform = column.get("transform", {})
        if not isinstance(transform, dict):
            raise ValueError(
                "Master column '{0}' transform must be a mapping".format(column_id)
            )

        transform_op = transform.get("op", "identity")
        validate_supported_value(
            label="columns.{0}.transform.op".format(column_id),
            value=transform_op,
            supported_values=validation.get("supported_transform_ops", []),
        )

        source_kind = source.get("kind")
        if source_kind == "metric":
            metric_ids = source.get("metric_ids")
            if not isinstance(metric_ids, list) or not metric_ids:
                raise ValueError(
                    "Master column '{0}' metric source must define metric_ids".format(
                        column_id
                    )
                )
            for metric_id in metric_ids:
                if metric_id not in loaded.template.metrics_by_id:
                    raise ValueError(
                        "Master column '{0}' references unknown metric_id '{1}'".format(
                            column_id, metric_id
                        )
                    )
        elif source_kind in {"roster_field", "meta_field", "ingestion_metadata"}:
            if not source.get("field"):
                raise ValueError(
                    "Master column '{0}' source kind '{1}' requires a field".format(
                        column_id, source_kind
                    )
                )
        else:
            raise ValueError(
                "Master column '{0}' uses unsupported source kind '{1}'".format(
                    column_id, source_kind
                )
            )

        duplicate_rule = transform.get("duplicate_metric_rule", default_duplicate_rule)
        if not isinstance(duplicate_rule, dict):
            raise ValueError(
                "Master column '{0}' duplicate_metric_rule must be a mapping".format(
                    column_id
                )
            )
        validate_duplicate_rule(
            label="columns.{0}.transform.duplicate_metric_rule".format(column_id),
            rule=duplicate_rule,
            validation=validation,
        )

        null_rule = column.get("null_rule", default_null_rule)
        if not isinstance(null_rule, dict):
            raise ValueError(
                "Master column '{0}' null_rule must be a mapping".format(column_id)
            )
        validate_null_rule(
            label="columns.{0}.null_rule".format(column_id),
            rule=null_rule,
            validation=validation,
        )


def load_master_generation_inputs(
    request: MasterGenerationRequest,
) -> LoadedMasterConfig:
    if not request.workbook_paths:
        raise ValueError("At least one workbook path is required")

    template = load_generation_inputs(
        TemplateGenerationRequest(
            config_path=request.config_path,
            entry_rows=1,
        )
    )

    files = template.config_doc.get("files", {})
    if not isinstance(files, dict):
        raise ValueError("config.yaml 'files' must be a mapping")

    master_candidate = files.get("master")
    if not master_candidate:
        raise ValueError("config.yaml 'files.master' is required for master generation")
    master_path = resolve_path(template.config_path, master_candidate)
    master_doc = load_yaml(master_path)

    loaded = LoadedMasterConfig(
        template=template,
        master_path=master_path,
        master_doc=master_doc,
    )
    validate_master_config(loaded)
    return loaded


def default_output_path(loaded: LoadedMasterConfig) -> Path:
    return (
        loaded.template.config_path.parent.parent
        / "workbooks"
        / "MasterWorkbook_v{0}.xlsx".format(loaded.master_doc.get("version", "unknown"))
    ).resolve()


def normalize_null_like(value: Any, rule: Mapping[str, Any]) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        candidate = value.strip() if rule.get("trim_strings", True) else value
        if candidate == "" and rule.get("blank_inputs_are_null", True):
            return None
        if candidate in set(rule.get("null_literals", [])):
            return None
        return candidate
    return value


def normalize_time_string(raw: str) -> Optional[float]:
    stripped = raw.strip()
    if not stripped:
        return None

    parts = stripped.split(":")
    try:
        if len(parts) == 2:
            minutes = int(parts[0])
            seconds = float(parts[1])
            return minutes * 60 + seconds
        if len(parts) == 3:
            hours = int(parts[0])
            minutes = int(parts[1])
            seconds = float(parts[2])
            return hours * 3600 + minutes * 60 + seconds
    except ValueError:
        return None
    return None


def normalize_timed_value(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, timedelta):
        return value.total_seconds()
    if isinstance(value, datetime):
        return (
            value.hour * 3600
            + value.minute * 60
            + value.second
            + value.microsecond / 1_000_000.0
        )
    if isinstance(value, time):
        return (
            value.hour * 3600
            + value.minute * 60
            + value.second
            + value.microsecond / 1_000_000.0
        )
    if isinstance(value, (int, float)):
        return float(value) * 86400 if abs(float(value)) <= 2 else float(value)
    if isinstance(value, str):
        parsed = normalize_time_string(value)
        if parsed is not None:
            return parsed
    raise ValueError("Unsupported timed value '{0}'".format(value))


def normalize_numeric_value(value: Any, *, integer: bool) -> Optional[Any]:
    if value is None:
        return None
    if isinstance(value, bool):
        numeric = int(value)
    elif isinstance(value, (int, float)):
        numeric = float(value)
    elif isinstance(value, str):
        numeric = float(value.strip())
    else:
        raise ValueError("Unsupported numeric value '{0}'".format(value))

    if integer:
        return int(round(numeric))
    return numeric


def normalize_metric_value(
    *,
    raw_value: Any,
    metric: Mapping[str, Any],
    null_rule: Mapping[str, Any],
) -> Any:
    value = normalize_null_like(raw_value, null_rule)
    if value is None:
        return None

    metric_type = metric.get("type")
    if metric_type == "timed":
        return normalize_timed_value(value)
    if metric_type == "integer":
        return normalize_numeric_value(value, integer=True)
    if metric_type == "numeric":
        return normalize_numeric_value(value, integer=False)
    if metric_type in {"categorical", "text"}:
        return str(value)
    return value


def read_meta_sheet(wb: Any) -> Dict[str, Any]:
    if "META" not in wb.sheetnames:
        raise ValueError("Workbook is missing required META sheet")

    ws = wb["META"]
    meta = {}
    row_index = 1
    while True:
        key = ws.cell(row=row_index, column=1).value
        if key in (None, ""):
            break
        meta[str(key)] = ws.cell(row=row_index, column=2).value
        row_index += 1
    return meta


def validate_node_workbook_meta(
    *,
    workbook_path: Path,
    meta: Mapping[str, Any],
    loaded: LoadedMasterConfig,
) -> None:
    expected_registry = loaded.template.config_doc.get("registry_name")
    if meta.get("registry_name") != expected_registry:
        raise ValueError(
            "Workbook {0} uses registry_name '{1}', expected '{2}'".format(
                workbook_path,
                meta.get("registry_name"),
                expected_registry,
            )
        )

    expected_version = str(loaded.template.config_doc.get("version"))
    actual_version = str(meta.get("config_version"))
    if actual_version != expected_version:
        raise ValueError(
            "Workbook {0} uses config_version '{1}', expected '{2}'".format(
                workbook_path,
                actual_version,
                expected_version,
            )
        )


def configured_roster_fields(loaded: LoadedMasterConfig) -> List[str]:
    sheet_contract = loaded.template.config_doc.get("sheet_contract", {})
    configured = sheet_contract.get("locked_left_columns", ["uid", "first", "last"])
    return ["uid", "first", "last"] + [
        field for field in configured if field not in {"uid", "first", "last"}
    ]


def sheet_row_range(ws: Any, start_row: int) -> range:
    return range(start_row, ws.max_row + 1)


def read_candidate_cell(ws: Any, row_index: int, column_index: int) -> Any:
    return ws.cell(row=row_index, column=column_index).value


def source_key(kind: str, field_name: str) -> Tuple[str, str]:
    return (kind, field_name)


def accumulate_source_value(
    *,
    bucket: CandidateAccumulator,
    kind: str,
    field_name: str,
    value: Any,
    null_rule: Mapping[str, Any],
) -> None:
    bucket.source_values[source_key(kind, field_name)].append(
        normalize_null_like(value, null_rule)
    )


def metric_columns_for_sheet(
    *,
    ws: Any,
    metric_id_row: int,
    metrics_by_id: Mapping[str, Mapping[str, Any]],
) -> Dict[int, str]:
    metric_columns = {}
    for col_idx in range(1, ws.max_column + 1):
        value = ws.cell(row=metric_id_row, column=col_idx).value
        if value in metrics_by_id:
            metric_columns[col_idx] = str(value)
    return metric_columns


def roster_columns_for_sheet(
    *,
    ws: Any,
    metric_id_row: int,
    roster_fields: Sequence[str],
) -> Dict[str, int]:
    roster_columns = {}
    roster_field_set = set(roster_fields)
    for col_idx in range(1, ws.max_column + 1):
        value = ws.cell(row=metric_id_row, column=col_idx).value
        if value in roster_field_set:
            roster_columns[str(value)] = col_idx
    return roster_columns


def ingest_node_workbook(
    *,
    workbook_path: Path,
    loaded: LoadedMasterConfig,
    buckets: MutableMapping[Tuple[str, str, str], CandidateAccumulator],
) -> None:
    workbook = load_workbook(workbook_path, data_only=True)
    try:
        meta = read_meta_sheet(workbook)
        validate_node_workbook_meta(
            workbook_path=workbook_path,
            meta=meta,
            loaded=loaded,
        )

        default_null_rule = loaded.master_doc.get("defaults", {}).get("null_rule", {})
        roster_fields = configured_roster_fields(loaded)
        sheet_contract = loaded.template.config_doc.get("sheet_contract", {})
        metric_id_row = int(sheet_contract.get("metric_id_row", 2))
        first_candidate_row = int(sheet_contract.get("first_candidate_row", 3))

        system_sheets = {"META", "ROSTER", "LOOKUPS", "MASTER"}
        for sheet_name in workbook.sheetnames:
            if sheet_name in system_sheets:
                continue

            ws = workbook[sheet_name]
            metric_columns = metric_columns_for_sheet(
                ws=ws,
                metric_id_row=metric_id_row,
                metrics_by_id=loaded.template.metrics_by_id,
            )
            if not metric_columns:
                continue

            roster_columns = roster_columns_for_sheet(
                ws=ws,
                metric_id_row=metric_id_row,
                roster_fields=roster_fields,
            )
            missing_roster_fields = [
                field for field in ("uid", "first", "last") if field not in roster_columns
            ]
            if missing_roster_fields:
                raise ValueError(
                    "Workbook {0} sheet '{1}' is missing roster machine headers for {2}".format(
                        workbook_path,
                        sheet_name,
                        ", ".join(missing_roster_fields),
                    )
                )

            for row_index in sheet_row_range(ws, first_candidate_row):
                uid_value = normalize_null_like(
                    read_candidate_cell(ws, row_index, roster_columns["uid"]),
                    default_null_rule,
                )
                if uid_value is None:
                    continue

                first_value = read_candidate_cell(ws, row_index, roster_columns["first"])
                last_value = read_candidate_cell(ws, row_index, roster_columns["last"])
                key = (
                    str(uid_value),
                    str(meta.get("block_number", "")),
                    str(meta.get("fiscal_year", "")),
                )
                bucket = buckets.setdefault(key, CandidateAccumulator())

                accumulate_source_value(
                    bucket=bucket,
                    kind="roster_field",
                    field_name="uid",
                    value=uid_value,
                    null_rule=default_null_rule,
                )
                accumulate_source_value(
                    bucket=bucket,
                    kind="roster_field",
                    field_name="first",
                    value=first_value,
                    null_rule=default_null_rule,
                )
                accumulate_source_value(
                    bucket=bucket,
                    kind="roster_field",
                    field_name="last",
                    value=last_value,
                    null_rule=default_null_rule,
                )
                accumulate_source_value(
                    bucket=bucket,
                    kind="meta_field",
                    field_name="block_number",
                    value=meta.get("block_number"),
                    null_rule=default_null_rule,
                )
                accumulate_source_value(
                    bucket=bucket,
                    kind="meta_field",
                    field_name="fiscal_year",
                    value=meta.get("fiscal_year"),
                    null_rule=default_null_rule,
                )
                accumulate_source_value(
                    bucket=bucket,
                    kind="ingestion_metadata",
                    field_name="source_workbook",
                    value=workbook_path.name,
                    null_rule=default_null_rule,
                )

                for col_idx, metric_id in metric_columns.items():
                    raw_value = read_candidate_cell(ws, row_index, col_idx)
                    normalized_value = normalize_metric_value(
                        raw_value=raw_value,
                        metric=loaded.template.metrics_by_id[metric_id],
                        null_rule=default_null_rule,
                    )
                    bucket.metric_values[metric_id].append(normalized_value)
    finally:
        workbook.close()


def render_null_output(rule: Mapping[str, Any]) -> Any:
    mode = rule.get("on_all_inputs_null", "null")
    if mode in (None, "null"):
        return None
    if mode == "blank":
        return ""
    if mode == "zero":
        return 0
    if mode == "literal":
        return rule.get("literal", "")
    raise ValueError("Unsupported null output mode '{0}'".format(mode))


def dedupe_preserve_order(values: Sequence[Any]) -> List[Any]:
    seen = []
    for value in values:
        if value not in seen:
            seen.append(value)
    return seen


def apply_aggregate_function(values: Sequence[Any], aggregate_function: str) -> Any:
    if aggregate_function == "count":
        return len(values)

    numeric_values = [float(value) for value in values]
    if aggregate_function == "average":
        return sum(numeric_values) / len(numeric_values)
    if aggregate_function == "sum":
        return sum(numeric_values)
    if aggregate_function == "min":
        return min(numeric_values)
    if aggregate_function == "max":
        return max(numeric_values)
    raise ValueError(
        "Unsupported aggregate function '{0}'".format(aggregate_function)
    )


def resolve_values(
    *,
    values: Sequence[Any],
    duplicate_rule: Mapping[str, Any],
    null_rule: Mapping[str, Any],
    column_id: str,
) -> Any:
    non_null = [value for value in values if value is not None]
    if not non_null:
        return render_null_output(null_rule)

    distinct_values = dedupe_preserve_order(non_null)
    if len(distinct_values) == 1:
        return distinct_values[0]

    mode = duplicate_rule.get("mode", "error_on_conflict")
    if mode == "error_on_conflict":
        raise ValueError(
            "Column '{0}' encountered conflicting values: {1}".format(
                column_id, distinct_values
            )
        )
    if mode == "first_non_null":
        return non_null[0]
    if mode == "last_non_null":
        return non_null[-1]
    if mode == "concat_distinct":
        separator = duplicate_rule.get("separator", "; ")
        return separator.join(str(value) for value in distinct_values)
    if mode == "aggregate":
        return apply_aggregate_function(
            distinct_values, duplicate_rule.get("aggregate_function", "")
        )

    raise ValueError(
        "Column '{0}' uses unsupported duplicate mode '{1}'".format(column_id, mode)
    )


def evaluate_column(
    *,
    bucket: CandidateAccumulator,
    column: Mapping[str, Any],
    default_duplicate_rule: Mapping[str, Any],
    default_null_rule: Mapping[str, Any],
) -> Any:
    source = column.get("source", {})
    transform = column.get("transform", {})
    duplicate_rule = transform.get("duplicate_metric_rule", default_duplicate_rule)
    null_rule = column.get("null_rule", default_null_rule)
    column_id = str(column.get("column_id", "unknown"))

    source_kind = source.get("kind")
    if source_kind == "metric":
        values = []
        for metric_id in source.get("metric_ids", []):
            values.extend(bucket.metric_values.get(metric_id, []))
    else:
        field_name = str(source.get("field"))
        values = bucket.source_values.get(source_key(source_kind, field_name), [])

    return resolve_values(
        values=values,
        duplicate_rule=duplicate_rule,
        null_rule=null_rule,
        column_id=column_id,
    )


def build_master_rows(loaded: LoadedMasterConfig) -> List[OrderedDict[str, Any]]:
    buckets: MutableMapping[Tuple[str, str, str], CandidateAccumulator] = OrderedDict()
    for workbook_path in loaded_request_workbooks(loaded):
        ingest_node_workbook(
            workbook_path=workbook_path,
            loaded=loaded,
            buckets=buckets,
        )

    columns = loaded.master_doc.get("columns", [])
    defaults = loaded.master_doc.get("defaults", {})
    default_duplicate_rule = defaults.get("duplicate_metric_rule", {})
    default_null_rule = defaults.get("null_rule", {})

    rows = []
    for bucket in buckets.values():
        row = OrderedDict()
        for column in columns:
            row[str(column.get("header", column.get("column_id")))] = evaluate_column(
                bucket=bucket,
                column=column,
                default_duplicate_rule=default_duplicate_rule,
                default_null_rule=default_null_rule,
            )
        rows.append(row)
    return rows


def loaded_request_workbooks(loaded: LoadedMasterConfig) -> Sequence[Path]:
    workbook_paths = loaded.master_doc.get("_request_workbook_paths")
    if not isinstance(workbook_paths, list):
        raise ValueError("Internal master generation state is missing workbook paths")
    return [Path(value) for value in workbook_paths]


def attach_request_workbooks(
    loaded: LoadedMasterConfig, workbook_paths: Sequence[Path]
) -> LoadedMasterConfig:
    master_doc = dict(loaded.master_doc)
    master_doc["_request_workbook_paths"] = [str(path) for path in workbook_paths]
    return LoadedMasterConfig(
        template=loaded.template,
        master_path=loaded.master_path,
        master_doc=master_doc,
    )


def build_master_workbook(
    *,
    loaded: LoadedMasterConfig,
    rows: Sequence[OrderedDict[str, Any]],
) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "MASTER"

    headers = [column.get("header", column.get("column_id")) for column in loaded.master_doc["columns"]]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    for row_idx, row in enumerate(rows, start=2):
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row.get(header))

    return wb


def generate_master_workbook(request: MasterGenerationRequest) -> Path:
    loaded = attach_request_workbooks(
        load_master_generation_inputs(request),
        workbook_paths=request.workbook_paths,
    )
    rows = build_master_rows(loaded)
    workbook = build_master_workbook(loaded=loaded, rows=rows)
    output_path = (
        request.output_path.resolve() if request.output_path else default_output_path(loaded)
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return output_path


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    output_path = generate_master_workbook(request_from_namespace(args))
    print("Master workbook generated: {0}".format(output_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
