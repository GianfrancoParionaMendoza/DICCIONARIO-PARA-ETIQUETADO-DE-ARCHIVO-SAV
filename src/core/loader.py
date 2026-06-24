from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import polars as pl
import pyreadstat


CHUNK_SIZE = 50_000


@dataclass
class SavFile:
    path: Path
    df: pl.DataFrame
    columns: list[str]
    column_labels: dict[str, str]
    value_labels: dict[str, dict]
    format_strings: dict = field(default_factory=dict)
    row_count: int = field(init=False)

    def __post_init__(self):
        self.row_count = len(self.df)


def _rename_duplicate_columns(df) -> None:
    """Rename duplicate column names in-place by appending _2, _3, etc."""
    seen: dict[str, int] = {}
    new_cols = []
    for col in df.columns:
        if col in seen:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_cols.append(col)
    df.columns = new_cols


def _downcast_floats_to_int(df: pl.DataFrame, fmt_strings: dict | None = None) -> pl.DataFrame:
    """Cast Float columns whose values are all whole numbers (or null) to Int64.
    For all-null columns, uses SPSS format metadata (e.g. 'F8.0' = integer)."""
    _int_fmt = re.compile(r'[Ff]\d+\.0$')
    casts = []
    for col in df.columns:
        if df[col].dtype not in (pl.Float32, pl.Float64):
            continue
        s = df[col].drop_nulls()
        if s.is_empty():
            fmt = (fmt_strings or {}).get(col, "")
            if fmt and _int_fmt.match(fmt):
                casts.append(pl.col(col).cast(pl.Int64))
            continue
        try:
            if s.is_nan().any():
                continue
            if (s == s.cast(pl.Int64).cast(df[col].dtype)).all():
                casts.append(pl.col(col).cast(pl.Int64))
        except Exception:
            pass
    return df.with_columns(casts) if casts else df


def load_sav(path: str | Path) -> SavFile:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {path}")

    frames: list[pl.DataFrame] = []
    meta = None

    for chunk_df, chunk_meta in pyreadstat.read_file_in_chunks(
        pyreadstat.read_sav,
        str(path),
        chunksize=CHUNK_SIZE,
    ):
        _rename_duplicate_columns(chunk_df)
        frames.append(pl.from_pandas(chunk_df))
        if meta is None:
            meta = chunk_meta
        del chunk_df

    df = pl.concat(frames) if frames else pl.DataFrame()
    del frames

    fmt_strings = getattr(meta, 'variable_format_strings', None) if meta else None
    df = _downcast_floats_to_int(df, fmt_strings)

    return SavFile(
        path=path,
        df=df,
        columns=list(df.columns),
        column_labels=meta.column_names_to_labels if meta else {},
        value_labels=meta.variable_value_labels if meta else {},
        format_strings=fmt_strings or {},
    )
