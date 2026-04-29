"""
Lightweight i18n for contract parsing.

Two layers vary by language:
  * regex PATTERNS used to extract parties / price / deadline / cross-refs
  * LABELS that appear next to parsed values in tool output

System-level text (tool descriptions, errors, log messages) stays English-only
and is NOT routed through this module.

Adding a third language: extend both PATTERNS and LABELS, then update Lang.
"""
from __future__ import annotations

from typing import Literal

Lang = Literal["ru", "en", "ru+en"]

DEFAULT_LANG: Lang = "ru"

# Cyrillic ratio thresholds for language classification:
#   ratio >= RU_THRESHOLD → "ru"
#   ratio <= EN_THRESHOLD → "en"
#   in between            → "ru+en" (bilingual, e.g. two-column EN-RU contract)
RU_THRESHOLD = 0.70
EN_THRESHOLD = 0.30


def detect_lang(text: str, *, sample_chars: int = 50_000) -> Lang:
    """Detect contract language by Cyrillic vs Latin character ratio.

    Sample size is large by default (50 KB) because bilingual contracts
    typically have one full-language column followed by another, and a
    short prefix would mis-classify them. For monolingual docs the answer
    converges within the first paragraph, so the larger sample is harmless.

    Returns "ru+en" when both scripts are well represented (each ≥30% of
    the alphabetic characters) — that's the signal for a bilingual contract.
    """
    if not text:
        return DEFAULT_LANG
    sample = text[:sample_chars]
    cyr = 0
    lat = 0
    for c in sample:
        lo = c.lower()
        if "а" <= lo <= "я" or lo == "ё":
            cyr += 1
        elif "a" <= lo <= "z":
            lat += 1
    total = cyr + lat
    if total == 0:
        return DEFAULT_LANG
    ratio = cyr / total
    if ratio >= RU_THRESHOLD:
        return "ru"
    if ratio <= EN_THRESHOLD:
        return "en"
    return "ru+en"


PATTERNS: dict[Lang, dict[str, str]] = {
    "ru": {
        # «Сторона А», именуемое в дальнейшем «Заказчик»
        "parties":
            r'«([^»]+)»[^«]*именуем\w+ в дальнейшем\s*[«"]([^»"]+)[»"]',
        # в размере 100 000 (сто тысяч) рублей
        "price":
            r'в размере\s+([\d\s]+)\s*\(([^)]+)\)\s*рублей',
        # Срок выполнения работ: 30 календарных дней
        "deadline":
            r'Срок выполнения работ?:\s*(\d+\s*(?:календарных|рабочих)\s*дней[^.]*)',
        # п. 4.2 / п. 4.2.1 / п. 4.2.
        "ref_link":
            r'п\.\s*([\d]+\.[\d]+\.?)',
    },
    "en": {
        # "Party A" (hereinafter referred to as "Client")
        "parties":
            r'"([^"]+)"[^"]*hereinafter referred to as\s*"([^"]+)"',
        # in the amount of 1,000 (one thousand) USD
        "price":
            r'in the amount of\s+([\d,]+)\s*\(([^)]+)\)\s*(USD|EUR|GBP|dollars|euros|pounds)',
        # within 30 calendar days / no later than 14 business days
        "deadline":
            r'(?:within|no later than)\s+(\d+\s*(?:calendar|business)\s*days?[^.]*)',
        # clause 4.2 / section 4.2 / clause 4.2.1
        "ref_link":
            r'(?:clause|section)\s*([\d]+\.[\d]+\.?)',
    },
}


LABELS: dict[Lang, dict[str, str]] = {
    "ru": {
        "price": "Цена",
        "currency": "руб.",
        "deadline": "Срок",
        "structure": "Структура",
        "sub_clauses": "подп.",
        "clause_prefix": "п.",
        "ref_marker": "п.",
    },
    "en": {
        "price": "Price",
        # currency comes from the regex match group; label fallback only
        "currency": "",
        "deadline": "Deadline",
        "structure": "Structure",
        "sub_clauses": "subs.",
        "clause_prefix": "Clause",
        "ref_marker": "Clause",
    },
    # Bilingual contracts (e.g. two-column EN-RU): show labels in both
    # languages. Currency comes from the regex match (RU pattern hardcodes
    # "руб.", EN pattern captures USD/EUR/etc.).
    "ru+en": {
        "price": "Цена / Price",
        "currency": "",
        "deadline": "Срок / Deadline",
        "structure": "Структура / Structure",
        "sub_clauses": "subs.",
        "clause_prefix": "п. / Clause",
        "ref_marker": "п. / Clause",
    },
}


def labels_for(lang: Lang | None = None, text: str | None = None) -> dict[str, str]:
    """Resolve a label dict — explicit lang wins, else auto-detect from text."""
    if lang is None:
        lang = detect_lang(text or "")
    return LABELS[lang]


def patterns_for(lang: Lang | None = None, text: str | None = None) -> dict[str, str]:
    """Pattern lookup. Bilingual mode is handled via `pattern_sets_for`.

    Returning a single dict for "ru+en" doesn't make sense (RU and EN
    regexes don't merge cleanly), so this helper raises for that case to
    push callers toward `pattern_sets_for`.
    """
    if lang is None:
        lang = detect_lang(text or "")
    if lang == "ru+en":
        raise ValueError(
            "patterns_for('ru+en') is ambiguous — use pattern_sets_for() "
            "and apply each language's patterns in turn."
        )
    return PATTERNS[lang]


def pattern_sets_for(lang: Lang | None = None, text: str | None = None) -> list[dict[str, str]]:
    """Return the list of pattern sets to apply.

    Monolingual → single-element list. Bilingual ("ru+en") → both.
    Callers iterate and dedupe matches by content.
    """
    if lang is None:
        lang = detect_lang(text or "")
    if lang == "ru+en":
        return [PATTERNS["ru"], PATTERNS["en"]]
    return [PATTERNS[lang]]
