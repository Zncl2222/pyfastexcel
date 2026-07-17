"""Regression tests for the fast encode/build paths and zip-level control.

The fast paths must stay observably identical to the careful implementations:
same bytes on the wire, same errors, same JSON fallback decisions.
"""

from __future__ import annotations

import os
import struct
import subprocess
import sys
import zipfile
from pathlib import Path

import msgspec
import pytest

import pyfastexcel.driver as driver_module
from pyfastexcel import CustomStyle, StreamWriter, set_zip_compression_level
from pyfastexcel.utils import set_custom_style
from pyfastexcel.wire import WIRE_MAGIC, _encode_no_style_row, _encode_styled_row, encode_v2_payload

ROOT = Path(__file__).resolve().parents[1]


class StyledWriter(StreamWriter):
    bold = CustomStyle(font_bold=True)
    red = CustomStyle(font_color='FF0000')


def careful_reference_payload(export_data) -> bytes:
    """The fast-path-free encoder: careful per-cell functions for every row."""
    sheet_order = list(export_data['sheet_order'])
    style_names = list(export_data['style'])
    style_ids = {name: index for index, name in enumerate(style_names)}

    metadata = dict(export_data)
    metadata_content = {}
    row_counts = []
    for sheet_name in sheet_order:
        sheet = export_data['content'][sheet_name]
        sheet_metadata = dict(sheet)
        rows = sheet.get('Data', [])
        sheet_metadata['Data'] = []
        metadata_content[sheet_name] = sheet_metadata
        row_counts.append(len(rows))

    row_stream = bytearray()
    sheet_offsets = []
    encoder = msgspec.msgpack.Encoder()
    for sheet_name in sheet_order:
        sheet_offsets.append(len(row_stream))
        sheet = export_data['content'][sheet_name]
        no_style = bool(sheet.get('NoStyle', False))
        for row in sheet.get('Data', []):
            encoder.encode_into(
                _encode_no_style_row(row) if no_style else _encode_styled_row(row, style_ids),
                row_stream,
                -1,
            )

    metadata['content'] = metadata_content
    metadata['_pyfastexcel_wire'] = {
        'version': 2,
        'style_names': style_names,
        'row_counts': row_counts,
        'sheet_offsets': sheet_offsets,
    }
    metadata_bytes = msgspec.json.encode(metadata)

    payload = bytearray(WIRE_MAGIC)
    payload.extend(struct.pack('>Q', len(metadata_bytes)))
    payload.extend(metadata_bytes)
    payload.extend(row_stream)
    return bytes(payload)


def test_fast_styled_encoding_matches_careful_encoder():
    writer = StyledWriter()
    for row in range(50):
        for col in range(8):
            writer.row_append(row * 8 + col, style='bold' if col % 2 else 'red')
        writer.create_row()

    worksheet = writer.workbook[writer.sheet]
    worksheet._data.append(
        [
            (1.5, 'bold'),
            (float('inf'), 'red'),
            (float('nan'), 'bold'),
            (float('-inf'), 'red'),
            None,
            (),
            (True, 'bold'),
            (False, 'red'),
            (None, 'bold'),
            ('text', 'DEFAULT_STYLE'),
            ('=SUM(A1:A2)', 'bold'),
            (-(1 << 62), 'red'),
            ((1 << 63) + 5, 'bold'),
        ]
    )
    export_data = writer._build_export_data()
    assert encode_v2_payload(export_data) == careful_reference_payload(export_data)


def test_fast_no_style_encoding_matches_careful_encoder():
    plain_rows = [
        ['a', 'b', 'c'],
        [1, 2.5, None, True, False],
        [float('inf'), float('nan'), 3],
        [1 << 63, -(1 << 63)],
    ]
    writer = StyledWriter()
    worksheet = writer.workbook[writer.sheet]
    worksheet._data = [list(row) for row in plain_rows]
    worksheet._sheet['NoStyle'] = True
    export_data = writer._build_export_data()
    export_data['content'][writer.sheet]['NoStyle'] = True
    assert encode_v2_payload(export_data) == careful_reference_payload(export_data)


def test_styled_cells_with_wrong_shapes_still_raise():
    writer = StyledWriter()
    writer.row_append(1, style='bold')
    writer.create_row()
    worksheet = writer.workbook[writer.sheet]

    worksheet._data.append([(1, 'bold', 'extra')])
    with pytest.raises(ValueError, match='exactly value and style'):
        encode_v2_payload(writer._build_export_data())

    worksheet._data[-1] = [(1, 'not-a-registered-style')]
    with pytest.raises(ValueError, match='not registered'):
        encode_v2_payload(writer._build_export_data())

    worksheet._data[-1] = ['bare string cell']
    with pytest.raises(TypeError, match='two-element'):
        encode_v2_payload(writer._build_export_data())


def test_row_append_unknown_style_raises_immediately():
    writer = StyledWriter()
    with pytest.raises(ValueError, match='not found'):
        writer.row_append(1, style='missing_style')


def test_row_append_accepts_styles_registered_after_first_rows():
    writer = StyledWriter()
    writer.row_append(1, style='bold')
    set_custom_style('late_style', CustomStyle(font_size=19))
    writer.row_append(2, style='late_style')
    writer.create_row()
    assert writer.workbook[writer.sheet].data[0] == [(1, 'bold'), (2, 'late_style')]


def test_append_row_with_per_column_styles_matches_row_append():
    per_cell = StyledWriter()
    for col, style in enumerate(('bold', 'red', 'DEFAULT_STYLE')):
        per_cell.row_append(col, style=style)
    per_cell.create_row()

    batched = StyledWriter()
    batched.append_row([0, 1, 2], style=['bold', 'red', 'DEFAULT_STYLE'])

    per_cell_rows = [list(map(tuple, row)) for row in per_cell.workbook[per_cell.sheet].data]
    batched_rows = [list(map(tuple, row)) for row in batched.workbook[batched.sheet].data]
    assert per_cell_rows == batched_rows


def test_append_rows_with_per_column_styles_and_custom_style_objects():
    style_object = CustomStyle(font_size=22)
    writer = StyledWriter()
    writer.append_rows([[1, 2], [3, 4]], style=['bold', style_object])
    rows = writer.workbook[writer.sheet].data
    assert rows[0][0] == (1, 'bold')
    generated_name = rows[0][1][1]
    assert rows[1] == ((3, 'bold'), (4, generated_name))
    assert writer.style.get_registered_style(generated_name) is style_object


def test_append_row_style_count_mismatch_raises():
    writer = StyledWriter()
    with pytest.raises(ValueError, match='2 values but 3 styles'):
        writer.append_row([1, 2], style=['bold', 'red', 'bold'])


def test_append_row_per_column_styles_reject_style_kwargs():
    writer = StyledWriter()
    with pytest.raises(ValueError, match='cannot be combined'):
        writer.append_row([1], style=['bold'], font_size=20)


def test_set_zip_compression_level_validation(monkeypatch):
    monkeypatch.setattr(driver_module, '_NATIVE_EXPORT_STARTED', False)
    monkeypatch.delenv('PYFASTEXCEL_ZIP_LEVEL', raising=False)

    with pytest.raises(ValueError, match='Invalid zip compression level'):
        set_zip_compression_level(0)
    with pytest.raises(ValueError, match='Invalid zip compression level'):
        set_zip_compression_level(10)

    set_zip_compression_level(6)
    assert os.environ['PYFASTEXCEL_ZIP_LEVEL'] == '6'
    set_zip_compression_level(None)
    assert 'PYFASTEXCEL_ZIP_LEVEL' not in os.environ

    monkeypatch.setattr(driver_module, '_NATIVE_EXPORT_STARTED', True)
    with pytest.raises(RuntimeError, match='before the first workbook export'):
        set_zip_compression_level(6)


_SUBPROCESS_EXPORT = '''
import sys
sys.path.insert(0, {root!r})
from pyfastexcel import CustomStyle, StreamWriter

class W(StreamWriter):
    bold = CustomStyle(font_bold=True)

writer = W()
for row in range(200):
    for col in range(10):
        writer.row_append(row * 10 + col, style='bold')
    writer.create_row()
writer.save({path!r})
'''


def _export_in_subprocess(tmp_path: Path, name: str, env_extra: dict[str, str]) -> Path:
    output = tmp_path / name
    env = dict(os.environ, **env_extra)
    completed = subprocess.run(
        [sys.executable, '-c', _SUBPROCESS_EXPORT.format(root=str(ROOT), path=str(output))],
        capture_output=True,
        text=True,
        env=env,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr
    if 'built without fast-zip support' in completed.stderr:
        pytest.skip('native library built without the pfx_fastzip tag')
    return output


def test_zip_level_produces_equivalent_readable_workbook(tmp_path):
    default_path = _export_in_subprocess(tmp_path, 'default.xlsx', {})
    fast_path = _export_in_subprocess(tmp_path, 'fast.xlsx', {'PYFASTEXCEL_ZIP_LEVEL': '6'})

    with zipfile.ZipFile(default_path) as default_zip, zipfile.ZipFile(fast_path) as fast_zip:
        assert default_zip.testzip() is None
        assert fast_zip.testzip() is None
        default_names = sorted(info.filename for info in default_zip.infolist())
        fast_names = sorted(info.filename for info in fast_zip.infolist())
        assert default_names == fast_names
        for entry in default_names:
            assert default_zip.read(entry) == fast_zip.read(entry)
