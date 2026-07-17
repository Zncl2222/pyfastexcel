from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import Any, Optional

from .style import CustomStyle
from .utils import validate_and_register_style
from .workbook import Workbook
from .worksheet import WorkSheet


class StreamWriter(Workbook):
    """
    A class for writing data to Excel files with or without custom styles.
    """

    def __init__(self, data: Optional[list[dict[str, str]]] = None):
        super().__init__()
        self._row_list = []
        self.data = data
        self._collections = self._get_style_collections()
        self._cache: dict[tuple[Any, ...], str] = {}
        # Style names that _resolve_style has already validated. Styles are
        # only ever registered or overridden, never removed, so a validated
        # name stays resolvable for the lifetime of this writer.
        self._validated_style_names: set[str] = {'DEFAULT_STYLE'}

    @property
    def wb(self) -> StreamWriter:
        return self

    @property
    def ws(self) -> WorkSheet:
        return self.workbook[self.sheet]

    @staticmethod
    def _canonicalize_cache_value(value: Any) -> Any:
        """Create a stable, hashable representation for nested style kwargs."""
        if isinstance(value, Mapping):
            return tuple(
                sorted(
                    (str(key), StreamWriter._canonicalize_cache_value(item))
                    for key, item in value.items()
                )
            )
        if isinstance(value, (list, tuple)):
            return tuple(StreamWriter._canonicalize_cache_value(item) for item in value)
        if isinstance(value, (set, frozenset)):
            return frozenset(StreamWriter._canonicalize_cache_value(item) for item in value)
        if hasattr(value, 'model_dump'):
            return StreamWriter._canonicalize_cache_value(value.model_dump())
        try:
            hash(value)
        except TypeError:
            return (type(value).__qualname__, repr(value))
        return value

    @classmethod
    def _style_fingerprint(cls, style: CustomStyle) -> tuple[Any, ...]:
        """Snapshot fields that affect a cloned style's serialized output."""
        return (
            cls._canonicalize_cache_value(style.font),
            cls._canonicalize_cache_value(style.fill),
            cls._canonicalize_cache_value(style.ali),
            cls._canonicalize_cache_value(style.border),
            cls._canonicalize_cache_value(style.protection),
            style.number_format,
        )

    def _resolve_style(self, style: str | CustomStyle, kwargs: dict[str, Any]) -> Any:
        """Resolve a public style input to a workbook-local style name."""
        if isinstance(style, str) and style == 'DEFAULT_STYLE' and not kwargs:
            return style

        if isinstance(style, CustomStyle):
            style_instance = style
            style_name = self.style.get_style_name(style_instance)
            if style_name is None:
                style_name = validate_and_register_style(style_instance, self.style)
        elif isinstance(style, str):
            style_name = style
            style_instance = self._collections.get(style_name)
            if style_instance is None:
                style_instance = self.style.get_registered_style(style_name)
            if style_instance is None:
                raise ValueError(f'Style {style_name} not found !')
            if not kwargs:
                return style_name
        else:
            # Preserve the historical behavior for callers that bypass the type
            # hint. The encoder will report unsupported style values as before.
            return style

        if not kwargs:
            return style_name

        cache_key = (
            style_name,
            self._style_fingerprint(style_instance),
            self._canonicalize_cache_value(kwargs),
        )
        if cache_key in self._cache:
            return self._cache[cache_key]

        new_style = style_instance.clone_and_modify(**kwargs)
        new_style_name = validate_and_register_style(new_style, self.style)
        self._cache[cache_key] = new_style_name
        return new_style_name

    def row_append(
        self,
        value: Any,
        style: str | CustomStyle = 'DEFAULT_STYLE',
        **kwargs,
    ) -> None:
        """
        Appends a value to the row list.

        Args:
            value (Any): The value to be appended.
            style (str | CustomStyle): The style of the value, can be either
                a style name or a CustomStyle object.
            **kwargs: Additional keyword arguments to modify the style.
        """
        if kwargs or type(style) is not str or style not in self._validated_style_names:  # noqa: E721
            style = self._resolve_style(style, kwargs)
            if not kwargs and type(style) is str:  # noqa: E721
                self._validated_style_names.add(style)
        if not isinstance(value, (int, float, str)):
            value = f'{value}'
        elif isinstance(value, float):
            value = float(value)
        self._row_list.append((value, style))

    def row_append_list(
        self,
        value: list[Any],
        style: str | CustomStyle = 'DEFAULT_STYLE',
        create_row: bool = False,
        **kwargs,
    ) -> None:
        """
        Appends a value to the row list.

        Args:
            value (list[Any]): The value to be appended.
            style (str | CustomStyle): The style of the value, can be either
                a style name or a CustomStyle object.
            create_row (bool): Whether to create row.
            **kwargs: Additional keyword arguments to modify the style.
        """
        if kwargs or type(style) is not str or style not in self._validated_style_names:  # noqa: E721
            style = self._resolve_style(style, kwargs)
            if not kwargs and type(style) is str:  # noqa: E721
                self._validated_style_names.add(style)
        value = tuple(
            (
                float(x) if isinstance(x, float) else x if isinstance(x, (int, str)) else f'{x}',
                style,
            )
            for x in value
        )

        if create_row:
            self.workbook[self.sheet].data.append(value)
        else:
            self._row_list.extend(value)

    def append_row(
        self,
        values: Iterable[Any],
        style: str | CustomStyle | list[str | CustomStyle] = 'DEFAULT_STYLE',
        **kwargs,
    ) -> None:
        """Append one complete row without changing the existing row APIs.

        ``style`` may also be a list or tuple with one style per column, which
        is substantially faster than one ``row_append`` call per cell when
        styles vary across columns.
        """
        if isinstance(style, (list, tuple)):
            if kwargs:
                raise ValueError('Per-column styles cannot be combined with style kwargs.')
            self.workbook[self.sheet].data.append(self._pair_row_with_styles(values, style))
            return
        self.row_append_list(values, style=style, create_row=True, **kwargs)

    def append_rows(
        self,
        rows: Iterable[Iterable[Any]],
        style: str | CustomStyle | list[str | CustomStyle] = 'DEFAULT_STYLE',
        **kwargs,
    ) -> None:
        """Append multiple complete rows using a shared style or per-column styles."""
        if isinstance(style, (list, tuple)):
            if kwargs:
                raise ValueError('Per-column styles cannot be combined with style kwargs.')
            resolved = self._resolve_style_row(style)
            data = self.workbook[self.sheet].data
            for row in rows:
                data.append(self._pair_row_with_resolved(row, resolved))
            return
        for row in rows:
            self.append_row(row, style=style, **kwargs)

    def _resolve_style_row(self, styles: list[str | CustomStyle]) -> tuple[str, ...]:
        """Resolve one style per column to validated workbook-local names."""
        resolved = []
        for style in styles:
            if type(style) is str and style in self._validated_style_names:  # noqa: E721
                resolved.append(style)
                continue
            name = self._resolve_style(style, {})
            if type(name) is str:  # noqa: E721
                self._validated_style_names.add(name)
            resolved.append(name)
        return tuple(resolved)

    def _pair_row_with_styles(
        self,
        values: Iterable[Any],
        styles: list[str | CustomStyle],
    ) -> tuple[tuple[Any, str], ...]:
        return self._pair_row_with_resolved(values, self._resolve_style_row(styles))

    def _pair_row_with_resolved(
        self,
        values: Iterable[Any],
        resolved: tuple[str, ...],
    ) -> tuple[tuple[Any, str], ...]:
        normalized = [
            x if isinstance(x, (int, str)) else float(x) if isinstance(x, float) else f'{x}'
            for x in values
        ]
        if len(normalized) != len(resolved):
            raise ValueError(
                f'Row has {len(normalized)} values but {len(resolved)} styles were given.',
            )
        return tuple(zip(normalized, resolved))

    def create_row(self):
        """
        Creates a row in the Excel data, and clean the current _row_list.
        """
        self.workbook[self.sheet].data.append(self._row_list)
        self._row_list = []
