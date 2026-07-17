from __future__ import annotations

import logging
import threading
from pathlib import Path
from typing import Any

from .logformatter import formatter, log_warning
from .style import CustomStyle

BASE_DIR = Path(__file__).resolve().parent

logger = logging.getLogger(__name__)
style_formatter = logging.StreamHandler()
style_formatter.setFormatter(formatter)

logger.addHandler(style_formatter)
logger.propagate = False


class StyleManager:
    """
    A class to set custom styles for Excel files.

    ### Methods:
        set_custom_style(cls, name: str, custom_style: CustomStyle): Set custom style
        by register method.
        _get_style_collections(): Gets collections of custom styles.
        _get_default_style(): Gets the default style.
        _update_style_map(style_name: str, custom_style: CustomStyle): Updates
            the style map.
        _get_font_style(style: CustomStyle): Gets the font style.
        _get_fill_style(style: CustomStyle): Gets the fill style.
        _get_border_style(style: CustomStyle): Gets the border style.
        _get_alignment_style(style: CustomStyle): Gets the alignment style.
        _get_protection_style(style: CustomStyle): Gets the protection style.
    """

    # ``set_custom_style`` is a process-level convenience API.  Workbooks never
    # mutate these dictionaries directly; every StyleManager instance maintains
    # a local overlay and lazily synchronizes process defaults by version.
    DEFAULT_STYLE = CustomStyle()
    REGISTERED_STYLES = {'DEFAULT_STYLE': DEFAULT_STYLE}
    _STYLE_NAME_MAP = {DEFAULT_STYLE: 'DEFAULT_STYLE'}
    _STYLE_ID = 0
    _style_map = {}
    _REGISTRY_LOCK = threading.RLock()
    _REGISTRY_VERSION = 0

    def __init__(self) -> None:
        """Initialize a workbook-local view of the process style defaults."""
        self._process_styles: dict[str, CustomStyle] = {}
        self._local_styles: dict[str, CustomStyle] = {}
        self._process_version = -1
        self._STYLE_ID = 0
        self._style_map: dict[str, dict[str, Any]] = {}
        self.REGISTERED_STYLES: dict[str, CustomStyle] = {}
        self._STYLE_NAME_MAP: dict[CustomStyle, str] = {}
        self.sync_defaults()

    @classmethod
    def set_custom_style(cls, name: str, custom_style: CustomStyle):
        """Register a process default used by existing and future workbooks."""
        with cls._REGISTRY_LOCK:
            registered_styles = dict(cls.REGISTERED_STYLES)
            if registered_styles.get(name):
                log_warning(
                    logger,
                    f'{name} has already existed. Overiding the style settings.',
                )
            registered_styles[name] = custom_style
            cls.REGISTERED_STYLES = registered_styles
            cls._STYLE_NAME_MAP = {
                style: style_name for style_name, style in registered_styles.items()
            }
            cls._REGISTRY_VERSION += 1

    @classmethod
    def register_generated_default_style(cls, custom_style: CustomStyle) -> str:
        """Atomically register an auto-named process default for compatibility."""
        with cls._REGISTRY_LOCK:
            style_name = f'Custom Style {cls._STYLE_ID}'
            cls._STYLE_ID += 1
            cls.set_custom_style(style_name, custom_style)
            return style_name

    @classmethod
    def reset_style_configs(cls):
        """
        Explicitly reset process defaults.

        This entry point is retained for compatibility and test isolation.  A
        workbook save intentionally never calls it.
        """
        with cls._REGISTRY_LOCK:
            cls.REGISTERED_STYLES = {'DEFAULT_STYLE': cls.DEFAULT_STYLE}
            cls._STYLE_NAME_MAP = {cls.DEFAULT_STYLE: 'DEFAULT_STYLE'}
            cls._STYLE_ID = 0
            cls._style_map = {}
            cls._REGISTRY_VERSION += 1

    def sync_defaults(self) -> None:
        """Lazily merge process defaults into this workbook's local overlay."""
        cls = type(self)
        if self._process_version == cls._REGISTRY_VERSION:
            return
        with cls._REGISTRY_LOCK:
            if self._process_version == cls._REGISTRY_VERSION:
                return
            process_styles = dict(cls.REGISTERED_STYLES)
            process_version = cls._REGISTRY_VERSION

        self._process_styles = process_styles
        self._process_version = process_version
        self._rebuild_registered_styles()

    def _rebuild_registered_styles(self) -> None:
        registered_styles = dict(self._process_styles)
        registered_styles.update(self._local_styles)
        self.REGISTERED_STYLES = registered_styles
        self._STYLE_NAME_MAP = {style: name for name, style in registered_styles.items()}

    def register_style(self, name: str, custom_style: CustomStyle) -> str:
        """Register a style only for this workbook."""
        self.sync_defaults()
        if self._local_styles.get(name) is custom_style:
            return name
        if name in self.REGISTERED_STYLES:
            log_warning(
                logger,
                f'{name} has already existed. Overriding the style settings.',
            )
            self._local_styles[name] = custom_style
            self._rebuild_registered_styles()
            return name

        # A unique local name is appended after both the process snapshot and
        # prior local styles, so the merged dictionaries can be updated in O(1).
        # Full rebuilds are reserved for name overrides and process syncs,
        # where stale reverse-map entries may need to be removed.
        self._local_styles[name] = custom_style
        self.REGISTERED_STYLES[name] = custom_style
        self._STYLE_NAME_MAP[custom_style] = name
        return name

    def register_generated_style(self, custom_style: CustomStyle) -> str:
        """Register an automatically named style in this workbook."""
        style_name = f'Custom Style {self._STYLE_ID}'
        self._STYLE_ID += 1
        return self.register_style(style_name, custom_style)

    def get_style_name(self, custom_style: CustomStyle) -> str | None:
        self.sync_defaults()
        return self._STYLE_NAME_MAP.get(custom_style)

    def get_registered_style(self, name: str) -> CustomStyle | None:
        self.sync_defaults()
        return self.REGISTERED_STYLES.get(name)

    def begin_style_build(self) -> None:
        """Start an atomic, repeatable serialization build for this workbook."""
        self.sync_defaults()
        self._style_map = {}

    def _get_default_style(self) -> dict[str, dict[str, Any] | str]:
        """
        Gets the default style.

        Returns:
            dict[str, dict[str, Any] | str]: A dictionary containing the
                default style settings.
        """
        return {
            'Font': {},
            'Fill': {},
            'Border': {},
            'Alignment': {},
            'Protection': {},
            'CustomNumFmt': 'general',
        }

    def _update_style_map(self, style_name: str, custom_style: CustomStyle) -> None:
        if self._style_map.get(style_name):
            log_warning(
                logger,
                f'{style_name} has already existed. Overriding the style settings.',
            )
        self._style_map[style_name] = self._get_default_style()
        self._style_map[style_name]['Font'] = self._get_font_style(custom_style)
        self._style_map[style_name]['Fill'] = self._get_fill_style(custom_style)
        self._style_map[style_name]['Border'] = self._get_border_style(custom_style)
        self._style_map[style_name]['Alignment'] = self._get_alignment_style(custom_style)
        self._style_map[style_name]['Protection'] = self._get_protection_style(custom_style)
        self._style_map[style_name]['CustomNumFmt'] = custom_style.number_format

    def _get_font_style(self, style: CustomStyle) -> dict[str, str | int | bool | None]:
        return style.font.model_dump(by_alias=True)

    def _get_fill_style(self, style: CustomStyle) -> dict[str, str]:
        return style.fill.model_dump(by_alias=True)

    def _get_border_style(self, style: CustomStyle) -> dict[str, str]:
        return style.border.model_dump(by_alias=True)

    def _get_alignment_style(self, style: CustomStyle) -> dict[str, str]:
        return style.ali.model_dump(by_alias=True)

    def _get_protection_style(self, style: CustomStyle) -> dict[str, str]:
        return style.protection.model_dump(by_alias=True)
