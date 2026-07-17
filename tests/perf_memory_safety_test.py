from __future__ import annotations

import base64
import ctypes
import io
import struct
import zipfile
from concurrent.futures import ThreadPoolExecutor

import msgspec
import pytest

import pyfastexcel.wire as wire_module
from pyfastexcel import CustomStyle, StreamWriter, Workbook
from pyfastexcel.driver import NativeExcelClient
from pyfastexcel.manager import StyleManager
from pyfastexcel.utils import set_custom_style, validate_and_register_style
from pyfastexcel.wire import WIRE_MAGIC, encode_payload, encode_v2_payload


class FakeCFunction:
    def __init__(self, callback):
        self.callback = callback
        self.argtypes = None
        self.restype = None

    def __call__(self, *args):
        return self.callback(*args)


class FakeNativeLibrary:
    def __init__(self, *, version: int = 1, raw_output: bytes = b'xlsx'):
        self.buffers = []
        self.freed = []
        self.payloads = []
        self.paths = []
        self.FreeCPointer = FakeCFunction(self._free)
        self.Export = FakeCFunction(self._legacy_export)
        self.raw_output = raw_output
        if version >= 2:
            self.GetABIVersion = FakeCFunction(lambda: version)
            self.ExportV2 = FakeCFunction(self._export_v2)
            self.ExportToFileV2 = FakeCFunction(self._export_to_file_v2)

    @staticmethod
    def _pointer_value(pointer):
        return pointer.value if hasattr(pointer, 'value') else pointer

    def _keep_buffer(self, data: bytes) -> int:
        buffer = ctypes.create_string_buffer(data)
        self.buffers.append(buffer)
        return ctypes.addressof(buffer)

    def _free(self, pointer, _debug):
        self.freed.append(self._pointer_value(pointer))

    def _legacy_export(self, payload, _catch_panic):
        self.payloads.append(bytes(payload))
        return self._keep_buffer(base64.b64encode(self.raw_output))

    @staticmethod
    def _read_payload(payload, payload_length):
        address = payload.value if hasattr(payload, 'value') else payload
        return bytes((ctypes.c_ubyte * payload_length).from_address(address))

    def _export_v2(self, payload, payload_length, _catch_panic, output_length, _error):
        self.payloads.append(self._read_payload(payload, payload_length))
        output_length._obj.value = len(self.raw_output)
        return self._keep_buffer(self.raw_output)

    def _export_to_file_v2(self, payload, payload_length, path, _catch_panic, _error):
        self.payloads.append(self._read_payload(payload, payload_length))
        self.paths.append(bytes(path))
        return 0


def _decode_v2_metadata(payload: bytes):
    assert payload[:4] == WIRE_MAGIC
    metadata_length = struct.unpack('>Q', payload[4:12])[0]
    metadata_end = 12 + metadata_length
    return msgspec.json.decode(payload[12:metadata_end]), payload[metadata_end:]


def test_late_process_registration_reaches_existing_workbook_and_stream_writer():
    workbook = Workbook()
    writer = StreamWriter()
    late_style = CustomStyle(font_bold=True)

    set_custom_style('late_style', late_style)
    workbook['Sheet1']['A1'] = ('workbook', 'late_style')
    writer.row_append('writer', style='late_style')

    assert workbook.style.get_registered_style('late_style') is late_style
    assert writer.ws.data == []
    writer.create_row()
    assert writer.ws.data == [[('writer', 'late_style')]]
    assert workbook._build_export_data()['style']['late_style']['Font']['Bold'] is True


def test_workbook_local_styles_are_isolated_serially_and_concurrently():
    def build(color: str):
        workbook = Workbook()
        style = CustomStyle(font_color=color)
        workbook['Sheet1']['A1'] = ('value', style)
        export_data = workbook._build_export_data()
        return workbook, export_data['style']['Custom Style 0']['Font']['Color']

    colors = [f'{index:06X}' for index in range(16)]
    with ThreadPoolExecutor(max_workers=8) as executor:
        results = list(executor.map(build, colors))

    assert [result[1] for result in results] == colors
    assert all(result[0].style._STYLE_ID == 1 for result in results)
    assert list(StyleManager.REGISTERED_STYLES) == ['DEFAULT_STYLE']


def test_real_native_exports_keep_concurrent_workbook_styles_isolated():
    colors = [f'{index + 1:06X}' for index in range(12)]

    def export(color: str) -> bool:
        workbook = Workbook()
        workbook['Sheet1']['A1'] = ('value', CustomStyle(font_color=color))
        with zipfile.ZipFile(io.BytesIO(workbook.read_lib_and_create_excel())) as archive:
            styles_xml = archive.read('xl/styles.xml').decode('utf-8')
        return f'rgb="FF{color}"' in styles_xml

    with ThreadPoolExecutor(max_workers=6) as executor:
        assert all(executor.map(export, colors))


def test_standalone_generated_process_style_names_are_atomic():
    styles = [CustomStyle(font_size=index + 1) for index in range(32)]

    with ThreadPoolExecutor(max_workers=8) as executor:
        names = list(executor.map(validate_and_register_style, styles))

    assert len(set(names)) == len(styles)
    assert set(names) == {f'Custom Style {index}' for index in range(len(styles))}
    assert all(StyleManager.REGISTERED_STYLES[name] is style for name, style in zip(names, styles))


def test_style_override_removes_stale_process_reverse_mapping():
    old_style = CustomStyle(font_color='111111')
    new_style = CustomStyle(font_color='222222')
    set_custom_style('same_name', old_style)
    set_custom_style('same_name', new_style)

    assert old_style not in StyleManager._STYLE_NAME_MAP
    assert StyleManager._STYLE_NAME_MAP[new_style] == 'same_name'


def test_unique_local_style_registration_is_incremental_and_overrides_rebuild(monkeypatch):
    manager = StyleManager()
    original_rebuild = manager._rebuild_registered_styles
    rebuild_calls = 0

    def rebuild_spy():
        nonlocal rebuild_calls
        rebuild_calls += 1
        original_rebuild()

    monkeypatch.setattr(manager, '_rebuild_registered_styles', rebuild_spy)
    local_styles = [CustomStyle(font_size=index + 1) for index in range(64)]

    for index, style in enumerate(local_styles):
        manager.register_style(f'local_{index}', style)

    assert rebuild_calls == 0
    assert all(
        manager.REGISTERED_STYLES[f'local_{index}'] is style
        for index, style in enumerate(local_styles)
    )
    assert all(
        manager.get_style_name(style) == f'local_{index}'
        for index, style in enumerate(local_styles)
    )

    replacement = CustomStyle(font_color='ABCDEF')
    manager.register_style('local_7', replacement)

    assert rebuild_calls == 1
    assert manager.get_registered_style('local_7') is replacement
    assert manager.get_style_name(replacement) == 'local_7'
    assert manager.get_style_name(local_styles[7]) is None

    late_process_style = CustomStyle(font_color='123456')
    set_custom_style('late_process_style', late_process_style)

    assert manager.get_registered_style('late_process_style') is late_process_style
    assert manager.get_style_name(late_process_style) == 'late_process_style'
    assert manager.get_registered_style('local_7') is replacement
    assert manager.get_style_name(local_styles[7]) is None
    assert rebuild_calls == 2


def test_workbook_local_class_and_generated_styles_win_process_name_collisions():
    process_class_style = CustomStyle(font_color='111111')
    process_generated_style = CustomStyle(font_color='222222')
    set_custom_style('collision', process_class_style)
    set_custom_style('Custom Style 0', process_generated_style)

    class CollidingWriter(StreamWriter):
        collision = CustomStyle(font_color='AAAAAA')

    writer = CollidingWriter()
    generated_style = CustomStyle(font_color='BBBBBB')
    writer.row_append('class', style='collision')
    writer.row_append('generated', style=generated_style)
    writer.create_row()

    style_map = writer._build_export_data()['style']

    assert style_map['collision']['Font']['Color'] == 'AAAAAA'
    assert style_map['Custom Style 0']['Font']['Color'] == 'BBBBBB'


def test_style_modifier_cache_uses_resolved_name_and_canonical_kwargs():
    writer = StreamWriter()
    decimal = CustomStyle(number_format='0.00', protect=False)
    percent = CustomStyle(number_format='0.00%', protect=True)

    writer.row_append(1, decimal, font_bold=True, font_color='AABBCC')
    writer.row_append(2, percent, font_color='AABBCC', font_bold=True)
    first_name, second_name = writer._row_list[0][1], writer._row_list[1][1]

    assert first_name != second_name
    assert writer.style.REGISTERED_STYLES[first_name].number_format == '0.00'
    assert writer.style.REGISTERED_STYLES[second_name].number_format == '0.00%'
    assert writer.style.REGISTERED_STYLES[second_name].protection.locked is True

    writer.row_append(3, decimal, font_color='AABBCC', font_bold=True)
    assert writer._row_list[2][1] == first_name


def test_style_modifier_cache_observes_base_style_mutation():
    writer = StreamWriter()
    base = CustomStyle(font_color='111111')

    writer.row_append(1, base, font_bold=True)
    first_name = writer._row_list[-1][1]
    base.set_custom_style(font_color='222222')
    writer.row_append(2, base, font_bold=True)
    second_name = writer._row_list[-1][1]

    assert first_name != second_name
    assert writer.style.REGISTERED_STYLES[first_name].font.color == '111111'
    assert writer.style.REGISTERED_STYLES[second_name].font.color == '222222'


def test_append_row_and_append_rows_preserve_public_cell_tuples():
    writer = StreamWriter()
    writer.append_row([1, 2])
    writer.append_rows([[3, 4], [5, None]])

    assert writer.ws.data == [
        ((1, 'DEFAULT_STYLE'), (2, 'DEFAULT_STYLE')),
        ((3, 'DEFAULT_STYLE'), (4, 'DEFAULT_STYLE')),
        ((5, 'DEFAULT_STYLE'), ('None', 'DEFAULT_STYLE')),
    ]


def test_v2_wire_framing_style_ids_and_public_data_compatibility():
    workbook = Workbook()
    style = CustomStyle(font_bold=True)
    set_custom_style('bold', style)
    worksheet = workbook['Sheet1']
    worksheet[0] = [('styled', 'bold'), 2]
    public_data_before = [row.copy() for row in worksheet.data]

    export_data = workbook._build_export_data()
    payload = encode_v2_payload(export_data)
    metadata, row_payload = _decode_v2_metadata(payload)

    assert metadata['_pyfastexcel_wire'] == {
        'version': 2,
        'style_names': ['DEFAULT_STYLE', 'bold'],
        'row_counts': [1],
        'sheet_offsets': [0],
    }
    assert metadata['content']['Sheet1']['Data'] == []
    assert row_payload == msgspec.msgpack.encode([('styled', 1), (2, 0)])
    assert worksheet.data == public_data_before
    assert worksheet.data == [[('styled', 'bold'), (2, 'DEFAULT_STYLE')]]


def test_v2_wire_keeps_no_style_rows_unmodified():
    workbook = Workbook(plain_data=[[1, 'two'], [None, 4]])
    export_data = workbook._build_export_data()
    payload = encode_v2_payload(export_data)
    metadata, row_payload = _decode_v2_metadata(payload)

    assert metadata['_pyfastexcel_wire']['row_counts'] == [2]
    assert row_payload == (msgspec.msgpack.encode([1, 'two']) + msgspec.msgpack.encode([None, 4]))


def test_v2_wire_matches_json_non_finite_float_semantics_without_mutating_data():
    workbook = Workbook()
    styled = workbook['Sheet1']
    styled[0] = [float('nan'), float('inf'), float('-inf'), 1.5]
    plain = workbook.create_sheet(
        'Plain',
        plain_data=[[float('nan'), float('inf'), float('-inf'), 2.5]],
    )
    export_data = workbook._build_export_data()

    payload = encode_v2_payload(export_data)
    metadata, row_payload = _decode_v2_metadata(payload)

    assert metadata['_pyfastexcel_wire']['row_counts'] == [1, 1]
    assert row_payload == (
        msgspec.msgpack.encode([(None, 0), (None, 0), (None, 0), (1.5, 0)])
        + msgspec.msgpack.encode([None, None, None, 2.5])
    )
    json_content = msgspec.json.decode(encode_payload(export_data, force_json=True))['content']
    assert json_content['Sheet1']['Data'] == [
        [
            [None, 'DEFAULT_STYLE'],
            [None, 'DEFAULT_STYLE'],
            [None, 'DEFAULT_STYLE'],
            [1.5, 'DEFAULT_STYLE'],
        ]
    ]
    assert json_content['Plain']['Data'] == [[None, None, None, 2.5]]
    assert styled.data[0][0][0] != styled.data[0][0][0]
    assert styled.data[0][1][0] == float('inf')
    assert styled.data[0][2][0] == float('-inf')
    assert plain.data[0][0] != plain.data[0][0]
    assert plain.data[0][1] == float('inf')
    assert plain.data[0][2] == float('-inf')


@pytest.mark.parametrize('value', [b'abc', bytearray(b'abc'), memoryview(b'abc'), 2**100])
def test_v2_negotiation_falls_back_for_legacy_only_cell_values(value):
    workbook = Workbook()
    workbook['Sheet1']['A1'] = (value, 'DEFAULT_STYLE')

    payload = encode_payload(workbook._build_export_data())

    assert not payload.startswith(WIRE_MAGIC)
    decoded_value = msgspec.json.decode(payload)['content']['Sheet1']['Data'][0][0][0]
    if isinstance(value, int):
        assert decoded_value == value
    else:
        assert decoded_value == 'YWJj'
    assert workbook['Sheet1']['A1'][0] is value


def test_v2_negotiation_falls_back_when_metadata_exceeds_decoder_limit(monkeypatch):
    workbook = Workbook()
    workbook['Sheet1']['A1'] = 'value'
    monkeypatch.setattr(wire_module, 'MAX_WIRE_METADATA_BYTES', 1)

    payload = encode_payload(workbook._build_export_data())

    assert not payload.startswith(WIRE_MAGIC)


@pytest.mark.parametrize('wire_name', ['json', 'v1-json'])
def test_json_wire_escape_hatch(monkeypatch, wire_name):
    workbook = Workbook()
    workbook['Sheet1']['A1'] = 'value'
    export_data = workbook._build_export_data()
    monkeypatch.setenv('PYFASTEXCEL_WIRE', wire_name)

    payload = encode_payload(export_data)

    assert not payload.startswith(WIRE_MAGIC)
    assert msgspec.json.decode(payload)['content']['Sheet1']['Data'] == [
        [['value', 'DEFAULT_STYLE']]
    ]


def test_legacy_abi_forces_json_before_export(monkeypatch):
    library = FakeNativeLibrary(version=1, raw_output=b'legacy-xlsx')
    workbook = Workbook()
    workbook['Sheet1']['A1'] = 'value'
    monkeypatch.setattr(workbook, '_read_lib', lambda _path: library)

    assert workbook.read_lib_and_create_excel() == b'legacy-xlsx'
    assert library.payloads[0].startswith(b'{')
    assert not library.payloads[0].startswith(WIRE_MAGIC)
    assert len(library.freed) == 1


def test_legacy_pointer_is_freed_when_python_postprocessing_raises(monkeypatch):
    library = FakeNativeLibrary(version=1)
    client = NativeExcelClient(library)

    def fail_decode(_payload):
        raise RuntimeError('decode failed')

    monkeypatch.setattr('pyfastexcel.driver.base64.b64decode', fail_decode)
    with pytest.raises(RuntimeError, match='decode failed'):
        client.export_bytes(b'{}', 1)

    assert len(library.freed) == 1


def test_v2_raw_bytes_and_pointer_cleanup(monkeypatch):
    library = FakeNativeLibrary(version=2, raw_output=b'PK\x00binary')
    client = NativeExcelClient(library)

    assert client.export_bytes(b'PFX2\x00payload', 1) == b'PK\x00binary'
    assert library.payloads == [b'PFX2\x00payload']
    assert len(library.freed) == 1

    monkeypatch.setattr(
        'pyfastexcel.driver.ctypes.string_at',
        lambda *_args: (_ for _ in ()).throw(RuntimeError('copy failed')),
    )
    with pytest.raises(RuntimeError, match='copy failed'):
        client.export_bytes(b'PFX2payload', 1)
    assert len(library.freed) == 2


def test_v2_error_pointer_is_reported_and_freed():
    library = FakeNativeLibrary(version=2)

    def fail_export(_payload, _length, _catch_panic, _output_length, error_pointer):
        address = library._keep_buffer(b'native failure')
        ctypes.cast(error_pointer, ctypes.POINTER(ctypes.c_char_p))[0] = ctypes.c_char_p(address)
        return None

    library.ExportV2 = FakeCFunction(fail_export)
    client = NativeExcelClient(library)

    with pytest.raises(RuntimeError, match='native failure'):
        client.export_bytes(b'PFX2payload', 1)

    assert library.freed == [ctypes.addressof(library.buffers[-1])]


def test_v2_rejects_non_null_zero_length_output_and_frees_it():
    library = FakeNativeLibrary(version=2, raw_output=b'')
    client = NativeExcelClient(library)

    with pytest.raises(RuntimeError, match='empty workbook'):
        client.export_bytes(b'PFX2payload', 1)

    assert len(library.freed) == 1


def test_save_uses_v2_direct_file_for_unicode_arbitrary_extension(monkeypatch):
    library = FakeNativeLibrary(version=2)
    workbook = Workbook()
    workbook['Sheet1']['A1'] = 'value'
    monkeypatch.setattr(workbook, '_read_lib', lambda _path: library)

    workbook.save('報表.data')

    assert library.paths == ['報表.data'.encode()]
    assert library.payloads[0].startswith(WIRE_MAGIC)


def test_repeated_path_save_includes_mutations_instead_of_stale_cached_bytes(tmp_path):
    workbook = Workbook()
    first_path = tmp_path / 'first.xlsx'
    second_path = tmp_path / 'second.xlsx'
    workbook['Sheet1']['A1'] = 'first'
    workbook.save(str(first_path))

    workbook['Sheet1']['A1'] = 'second'
    workbook.save(str(second_path))

    with zipfile.ZipFile(first_path) as archive:
        first_xml = archive.read('xl/worksheets/sheet1.xml')
    with zipfile.ZipFile(second_path) as archive:
        second_xml = archive.read('xl/worksheets/sheet1.xml')
    assert b'<t>first</t>' in first_xml
    assert b'<t>second</t>' in second_xml


def test_save_rejects_embedded_nul_before_native_path_truncation(monkeypatch, tmp_path):
    library = FakeNativeLibrary(version=2)
    workbook = Workbook()
    monkeypatch.setattr(workbook, '_read_lib', lambda _path: library)
    intended = tmp_path / 'intended'

    with pytest.raises(ValueError, match='embedded null byte'):
        workbook.save(f'{intended}\x00suffix')

    assert not intended.exists()
    assert library.paths == []
