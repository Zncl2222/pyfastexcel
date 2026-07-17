from __future__ import annotations

import math
import os
import struct
from typing import Any

import msgspec

WIRE_MAGIC = b'PFX2'
WIRE_VERSION = 2
WIRE_ENV_VAR = 'PYFASTEXCEL_WIRE'
MAX_WIRE_METADATA_BYTES = 64 << 20
_MSGPACK_MIN_INT = -(1 << 63)
_MSGPACK_MAX_INT = (1 << 64) - 1


class _UseLegacyJSON(Exception):
    """Signal that a cell needs the legacy JSON value semantics."""


class _RowNeedsCare(Exception):
    """Signal that a row cannot take the fast encode path."""


def use_json_wire() -> bool:
    """Return whether the human-readable legacy wire was explicitly requested."""
    return os.getenv(WIRE_ENV_VAR, '').strip().lower() in {'json', 'v1-json'}


def encode_json_payload(export_data: dict[str, Any]) -> bytes:
    """Encode the complete legacy payload."""
    return msgspec.json.encode(export_data)


def _normalize_scalar(value: Any) -> Any:
    """Normalize the scalar subset shared exactly by JSON and PFX2."""
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if isinstance(value, int):
        if _MSGPACK_MIN_INT <= value <= _MSGPACK_MAX_INT:
            return value
        raise _UseLegacyJSON
    if value is None or isinstance(value, (bool, float, str)):
        return value

    # msgspec JSON has established public behavior for values such as bytes
    # (base64), dates and container objects. Fall back as one complete payload
    # instead of maintaining a second, subtly different conversion table.
    raise _UseLegacyJSON


def _encode_no_style_row(row: Any) -> Any:
    """Copy only rows that contain a non-finite float."""
    for index, value in enumerate(row):
        normalized = _normalize_scalar(value)
        if normalized is value:
            continue
        encoded_row = list(row)
        encoded_row[index] = normalized
        for tail_index in range(index + 1, len(encoded_row)):
            encoded_row[tail_index] = _normalize_scalar(encoded_row[tail_index])
        return encoded_row
    return row


def _fast_no_style_row(row: Any) -> Any:
    """Pass through a row of plain scalars without per-value function calls.

    Raises _RowNeedsCare for anything the tight type dispatch does not cover,
    so ``_encode_no_style_row`` keeps the exact legacy semantics for rare rows.
    """
    isfinite = math.isfinite
    for value in row:
        value_type = type(value)
        if value_type is int:
            if _MSGPACK_MIN_INT <= value <= _MSGPACK_MAX_INT:
                continue
            raise _UseLegacyJSON
        if value_type is str or value is None or value_type is bool:
            continue
        if value_type is float:
            if isfinite(value):
                continue
            raise _RowNeedsCare
        raise _RowNeedsCare
    return row


def _fast_styled_row(row: Any, style_ids: dict[str, int]) -> list[Any]:
    """Encode the common case of a row of well-formed ``(value, style)`` cells.

    Exact-type dispatch keeps this loop cheap; any irregular cell shape,
    subclassed value, unknown style, or out-of-range integer raises so the
    caller can retry with ``_encode_styled_row`` and preserve its exact error
    and fallback behavior.
    """
    encoded_row = []
    append = encoded_row.append
    isfinite = math.isfinite
    for cell in row:
        if type(cell) is not tuple and type(cell) is not list:  # noqa: E721
            raise _RowNeedsCare
        value, style_name = cell
        value_type = type(value)
        if value_type is int:
            if not (_MSGPACK_MIN_INT <= value <= _MSGPACK_MAX_INT):
                raise _UseLegacyJSON
        elif value_type is float:
            if not isfinite(value):
                value = None
        elif value_type is not str and value is not None and value_type is not bool:
            raise _RowNeedsCare
        append((value, style_ids[style_name]))
    return encoded_row


def _encode_styled_row(row: Any, style_ids: dict[str, int]) -> list[Any]:
    encoded_row = []
    for cell in row:
        if cell is None:
            encoded_row.append(None)
            continue
        if not isinstance(cell, (list, tuple)):
            raise TypeError('Styled cell data should be a two-element list or tuple.')
        if len(cell) == 0:
            encoded_row.append(())
            continue
        if len(cell) != 2:
            raise ValueError('Styled cell data should contain exactly value and style.')

        style_name = cell[1]
        try:
            style_id = style_ids[style_name]
        except (KeyError, TypeError) as exc:
            raise ValueError(f'Style {style_name!r} is not registered in this workbook.') from exc
        encoded_row.append((_normalize_scalar(cell[0]), style_id))
    return encoded_row


def encode_v2_payload(export_data: dict[str, Any]) -> bytes:
    """Encode the version-2 metadata + row-stream framing.

    Layout::

        PFX2 | uint64(metadata length, big endian) | metadata JSON | msgpack rows...

    Each row is one complete msgpack object. ``row_counts`` and ``sheet_order``
    make additional per-row framing unnecessary.
    """
    sheet_order = list(export_data['sheet_order'])
    style_names = list(export_data['style'])
    style_ids = {name: index for index, name in enumerate(style_names)}

    metadata = dict(export_data)
    metadata_content: dict[str, Any] = {}
    row_counts: list[int] = []

    for sheet_name in sheet_order:
        sheet = export_data['content'][sheet_name]
        sheet_metadata = dict(sheet)
        rows = sheet.get('Data', [])
        sheet_metadata['Data'] = []
        metadata_content[sheet_name] = sheet_metadata
        row_counts.append(len(rows))

    # Rows are encoded before the metadata so each sheet's byte offset into
    # the row stream can ride along; Go uses the offsets to decode and write
    # multiple sheets concurrently.
    row_stream = bytearray()
    sheet_offsets: list[int] = []
    encoder = msgspec.msgpack.Encoder()
    encode_into = encoder.encode_into
    for sheet_name in sheet_order:
        sheet_offsets.append(len(row_stream))
        sheet = export_data['content'][sheet_name]
        no_style = bool(sheet.get('NoStyle', False))
        for row in sheet.get('Data', []):
            # The tight loops cover well-formed scalar rows; anything unusual
            # retries through the careful encoders, which own the exact error
            # messages and the legacy-JSON fallback semantics.
            if no_style:
                try:
                    encoded_row = _fast_no_style_row(row)
                except _RowNeedsCare:
                    encoded_row = _encode_no_style_row(row)
            else:
                try:
                    encoded_row = _fast_styled_row(row, style_ids)
                except (_RowNeedsCare, TypeError, ValueError, KeyError):
                    encoded_row = _encode_styled_row(row, style_ids)
            encode_into(encoded_row, row_stream, -1)

    metadata['content'] = metadata_content
    metadata['_pyfastexcel_wire'] = {
        'version': WIRE_VERSION,
        'style_names': style_names,
        'row_counts': row_counts,
        'sheet_offsets': sheet_offsets,
    }
    metadata_bytes = msgspec.json.encode(metadata)
    if len(metadata_bytes) > MAX_WIRE_METADATA_BYTES:
        raise _UseLegacyJSON

    payload = bytearray(WIRE_MAGIC)
    payload.extend(struct.pack('>Q', len(metadata_bytes)))
    payload.extend(metadata_bytes)
    payload.extend(row_stream)
    return bytes(payload)


def encode_payload(export_data: dict[str, Any], *, force_json: bool = False) -> bytes:
    """Encode an export payload, honoring the JSON debugging escape hatch."""
    if force_json or use_json_wire():
        return encode_json_payload(export_data)
    try:
        return encode_v2_payload(export_data)
    except _UseLegacyJSON:
        return encode_json_payload(export_data)
