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

    metadata['content'] = metadata_content
    metadata['_pyfastexcel_wire'] = {
        'version': WIRE_VERSION,
        'style_names': style_names,
        'row_counts': row_counts,
    }
    metadata_bytes = msgspec.json.encode(metadata)
    if len(metadata_bytes) > MAX_WIRE_METADATA_BYTES:
        raise _UseLegacyJSON

    payload = bytearray(WIRE_MAGIC)
    payload.extend(struct.pack('>Q', len(metadata_bytes)))
    payload.extend(metadata_bytes)
    encoder = msgspec.msgpack.Encoder()
    for sheet_name in sheet_order:
        sheet = export_data['content'][sheet_name]
        no_style = bool(sheet.get('NoStyle', False))
        for row in sheet.get('Data', []):
            encoder.encode_into(
                _encode_no_style_row(row) if no_style else _encode_styled_row(row, style_ids),
                payload,
                -1,
            )
    return bytes(payload)


def encode_payload(export_data: dict[str, Any], *, force_json: bool = False) -> bytes:
    """Encode an export payload, honoring the JSON debugging escape hatch."""
    if force_json or use_json_wire():
        return encode_json_payload(export_data)
    try:
        return encode_v2_payload(export_data)
    except _UseLegacyJSON:
        return encode_json_payload(export_data)
