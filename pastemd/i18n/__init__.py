"""Lightweight translation utilities for PasteMD."""

from __future__ import annotations

import json
import locale
import logging
import os
from typing import Dict, Iterator, Optional
from ..config.paths import resource_path


DEFAULT_LANGUAGE = "zh"

# Display names for each supported language (order preserved).
LANGUAGE_DISPLAY_NAMES: Dict[str, str] = {
    "zh": "简体中文",
    "en": "English",
}

_LOCALE_PACKAGE = "pastemd.i18n.locales"
_LOCALE_FILES = {
    "zh": "zh.json",
    "en": "en.json",
}

_loaded_translations: Dict[str, Dict[str, str]] = {}
_current_language = DEFAULT_LANGUAGE
_logger = logging.getLogger(__name__)


def _normalize_language_code(language: Optional[str]) -> Optional[str]:
    """Normalize locale strings like 'en_US' or 'en-US' to 'en'."""
    if not language:
        return None
    normalized = language.replace("_", "-").split("-")[0].lower()
    return normalized or None


def _load_translations(language: str) -> Dict[str, str]:
    """Load translation dictionary for a language (with caching) from file system."""
    normalized = _normalize_language_code(language) or DEFAULT_LANGUAGE
    if normalized in _loaded_translations:
        return _loaded_translations[normalized]

    data: Dict[str, str] = {}
    file_name = _LOCALE_FILES.get(normalized)

    if file_name:
        try:
            json_path = resource_path(os.path.join("i18n", "locales", file_name))
            if not os.path.isfile(json_path):
                json_path = os.path.join(os.path.join("pastemd", "i18n", "locales", file_name))

            with open(json_path, "r", encoding="utf-8") as fp:
                loaded = json.load(fp)

            if isinstance(loaded, dict):
                data = {str(k): str(v) for k, v in loaded.items()}
        
        except FileNotFoundError:
            _logger.warning("Translation file missing for %s at %s", normalized, json_path)
        except Exception as exc:
            _logger.warning("Failed to load translations for %s: %s", normalized, exc)

    _loaded_translations[normalized] = data or {}
    return _loaded_translations[normalized]


# Ensure default translations are available for fallback.
_load_translations(DEFAULT_LANGUAGE)


def is_supported_language(language: Optional[str]) -> bool:
    """Return True if the given language code is supported."""
    normalized = _normalize_language_code(language)
    return bool(normalized and normalized in LANGUAGE_DISPLAY_NAMES)


def set_language(language: str) -> None:
    """Set the active language if supported, otherwise fall back to the default."""
    global _current_language
    normalized = _normalize_language_code(language)
    if normalized in LANGUAGE_DISPLAY_NAMES:
        _current_language = normalized
    else:
        _current_language = DEFAULT_LANGUAGE
    _load_translations(_current_language)


def get_language() -> str:
    """Return the current UI language code."""
    return _current_language


def get_language_label(language: str) -> str:
    """Return the human readable label for a language code."""
    normalized = _normalize_language_code(language) or language
    return LANGUAGE_DISPLAY_NAMES.get(normalized, language)


def iter_languages() -> Iterator[tuple[str, str]]:
    """Yield supported languages (code, label) pairs."""
    for code, label in LANGUAGE_DISPLAY_NAMES.items():
        yield code, label


def detect_system_language() -> Optional[str]:
    """
    Detect the system UI language (Windows preferred) and return a supported code.

    Returns:
        Language code if detection succeeds and the language is supported, otherwise None.
    """
    candidates: list[str | None] = []

    # Windows-specific detection (preferred if available).
    try:
        import ctypes
        from locale import windows_locale

        lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
        candidates.append(windows_locale.get(lang_id))
    except Exception:
        pass

    # Fallback to locale data provided by Python.
    try:
        lang, _encoding = locale.getdefaultlocale()
        candidates.append(lang)
    except (ValueError, TypeError):
        pass

    try:
        lang, _encoding = locale.getlocale()
        candidates.append(lang)
    except (ValueError, TypeError):
        pass

    for candidate in candidates:
        normalized = _normalize_language_code(candidate)
        if normalized and normalized in LANGUAGE_DISPLAY_NAMES:
            return normalized

    return None


def t(key: str, **kwargs) -> str:
    """
    Translate a key into the active language.

    Args:
        key: Translation key.
        **kwargs: Optional format arguments inserted via str.format.
    """
    translations = _load_translations(_current_language)
    text = translations.get(key)

    if text is None and _current_language != DEFAULT_LANGUAGE:
        text = _load_translations(DEFAULT_LANGUAGE).get(key)

    if text is None:
        # As a last resort look through any other loaded languages.
        for data in _loaded_translations.values():
            if key in data:
                text = data[key]
                break

    if text is None:
        text = key

    if kwargs:
        try:
            text = text.format(**kwargs)
        except (KeyError, IndexError, ValueError):
            pass

    return text


def get_no_app_action_map() -> Dict[str, str]:
    """获取动作值到显示文本的映射（用于 UI）"""
    from ..core.types import NoAppAction
    return {
        NoAppAction.OPEN.value: t("action.open"),
        NoAppAction.SAVE.value: t("action.save"),
        NoAppAction.CLIPBOARD.value: t("action.clipboard"),
        NoAppAction.NONE.value: t("action.none"),
    }
