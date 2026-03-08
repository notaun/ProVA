"""
core.py — Minimal data validation & normalization for ProVA (V2).

Responsibilities:
- Validate structural assumptions
- Enforce explicit schemas
- Provide low-level utilities (IDs, hashes)

Non-responsibilities:
- Cleaning dirty data
- Guessing column meaning
- Feature engineering
- KPI computation
"""

from __future__ import annotations
from typing import Dict, List, Optional
from dataclasses import dataclass
import pandas as pd
import hashlib
import uuid


# -----------------------------
# Errors
# -----------------------------
class SchemaError(ValueError):
    pass


# -----------------------------
# Schema definition
# -----------------------------
@dataclass(frozen=True)
class Schema:
    """
    Explicit schema contract.

    columns:
        dict[column_name -> pandas dtype or callable validator]
    """
    columns: Dict[str, object]


# -----------------------------
# Validation
# -----------------------------
def validate_schema(df: pd.DataFrame, schema: Schema) -> None:
    if df is None or df.empty:
        raise SchemaError("DataFrame is empty")

    for col, rule in schema.columns.items():
        if col not in df.columns:
            raise SchemaError(f"Missing required column: '{col}'")

        series = df[col]

        # dtype check
        if isinstance(rule, str):
            if not pd.api.types.is_dtype_equal(series.dtype, rule):
                raise SchemaError(
                    f"Column '{col}' expected dtype '{rule}', got '{series.dtype}'"
                )

        # callable validator
        elif callable(rule):
            if not rule(series):
                raise SchemaError(f"Column '{col}' failed validation rule")

        # ignore unknown rule types


# -----------------------------
# Utilities
# -----------------------------
def ensure_record_id(df: pd.DataFrame, column: str = "RecordID") -> pd.DataFrame:
    if column in df.columns and df[column].notna().all():
        return df

    df = df.copy()
    df[column] = [str(uuid.uuid4()) for _ in range(len(df))]
    return df


def make_row_hash(row: pd.Series) -> str:
    parts = []
    for k in sorted(row.index):
        v = row.get(k)
        parts.append(str(v) if pd.notna(v) else "<NA>")
    base = "|".join(parts)
    return hashlib.sha256(base.encode("utf-8")).hexdigest()


def add_row_hash(df: pd.DataFrame, column: str = "RowHash") -> pd.DataFrame:
    df = df.copy()
    df[column] = df.apply(make_row_hash, axis=1)
    return df


# -----------------------------
# Convenience
# -----------------------------
def assert_columns(df: pd.DataFrame, required: List[str]):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise SchemaError(f"Missing required columns: {missing}")
