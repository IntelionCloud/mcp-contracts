"""Tests for the i18n module: language detection + label/pattern selection."""
from __future__ import annotations

from pathlib import Path

import pytest

from core.contract_model import (
    contract_summary,
    parse_contract,
    validate_references,
)
from core.i18n import LABELS, PATTERNS, detect_lang, labels_for, patterns_for


FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE_EN = (FIXTURES / "sample_en.md").read_text(encoding="utf-8")

SAMPLE_RU = """
ДОГОВОР ОКАЗАНИЯ УСЛУГ

Настоящий договор заключен между ООО «Альфа» («Заказчик»),
именуемое в дальнейшем «Заказчик», и ООО «Бета» («Исполнитель»),
именуемое в дальнейшем «Исполнитель».

1. ПРЕДМЕТ ДОГОВОРА

1.1. Исполнитель обязуется оказать услуги, указанные в п. 2.1.

1.2. Заказчик обязуется оплатить услуги в размере 100 000
(сто тысяч) рублей в порядке, указанном в п. 9.9 договора.

2. СРОКИ

2.1. Срок выполнения работ: 30 календарных дней.
"""


# ---------------------------------------------------------------------------
# detect_lang
# ---------------------------------------------------------------------------


def test_detect_lang_ru():
    assert detect_lang(SAMPLE_RU) == "ru"


def test_detect_lang_en():
    assert detect_lang(SAMPLE_EN) == "en"


def test_detect_lang_empty_falls_back_to_default():
    assert detect_lang("") == "ru"  # current DEFAULT_LANG


def test_detect_lang_only_digits_falls_back():
    # No alphabetic chars at all → default
    assert detect_lang("1.2.3 — 4 5 6 7 8 9 10") == "ru"


def test_detect_lang_mixed_majority_wins():
    # 3 lines of English, 1 short Russian phrase → English wins
    mixed = (
        "This is a long english introduction with many words. " * 5
        + "Краткая фраза."
    )
    assert detect_lang(mixed) == "en"


def test_detect_lang_bilingual():
    # ~50/50 mix → bilingual detected
    bi = (
        "Limited Liability Company Agrobiotech, hereinafter Contractor. "
        "Общество с ограниченной ответственностью Агробиотех, именуемое Исполнитель. "
    ) * 5
    assert detect_lang(bi) == "ru+en"


# ---------------------------------------------------------------------------
# labels_for / patterns_for
# ---------------------------------------------------------------------------


def test_labels_for_explicit_language():
    assert labels_for("ru")["price"] == "Цена"
    assert labels_for("en")["price"] == "Price"


def test_labels_for_auto_detect_from_text():
    assert labels_for(text=SAMPLE_RU)["structure"] == "Структура"
    assert labels_for(text=SAMPLE_EN)["structure"] == "Structure"


def test_patterns_for_returns_language_specific_regex():
    assert patterns_for("ru")["price"] != patterns_for("en")["price"]
    # Sanity: each language has the same logical keys.
    assert set(PATTERNS["ru"]) == set(PATTERNS["en"])
    assert set(LABELS["ru"]) == set(LABELS["en"])


# ---------------------------------------------------------------------------
# contract_summary integration
# ---------------------------------------------------------------------------


def test_summary_ru_uses_russian_labels():
    clauses = parse_contract(SAMPLE_RU)
    out = contract_summary(clauses, SAMPLE_RU)
    assert "Цена:" in out
    assert "Срок:" in out
    assert "Структура:" in out
    assert "100 000" in out
    assert "(сто тысяч)" in out


def test_summary_en_uses_english_labels():
    clauses = parse_contract(SAMPLE_EN)
    out = contract_summary(clauses, SAMPLE_EN)
    assert "Price:" in out
    assert "Deadline:" in out
    assert "Structure:" in out
    # parties match: hereinafter referred to as
    assert "Provider" in out and "Client" in out


def test_summary_explicit_override_swaps_labels():
    """Forcing language=en on a Russian doc swaps labels but RU patterns
    won't match — that's expected (the parsed values are RU). Output should
    still have EN labels for the *frame* (Structure/Price)."""
    clauses = parse_contract(SAMPLE_RU)
    out = contract_summary(clauses, SAMPLE_RU, language="en")
    assert "Structure:" in out
    # Russian pattern won't match because we asked for EN patterns:
    assert "Цена:" not in out


def test_summary_bilingual_uses_dual_labels_and_dedupes():
    """Bilingual contract → ru+en labels; price/parties not duplicated even
    when both RU and EN patterns hit."""
    bi = SAMPLE_RU + "\n\n" + SAMPLE_EN
    clauses = parse_contract(bi)
    out = contract_summary(clauses, bi)
    # Bilingual labels:
    assert "Структура / Structure:" in out
    assert "Цена / Price:" in out or "Срок / Deadline:" in out
    # Both prices should be present once each (RU 100k rub + EN 12k USD).
    # Dedupe is per (amount, words), so no double-listing of the same value.
    assert out.count("100 000 (сто тысяч)") == 1
    assert out.count("12,000 (twelve thousand)") == 1


# ---------------------------------------------------------------------------
# validate_references integration
# ---------------------------------------------------------------------------


def test_validate_references_ru_finds_dangling_ref():
    # Sample says "в порядке, указанном в п. 9.9 договора" — there's no 9.9.
    clauses = parse_contract(SAMPLE_RU)
    issues = validate_references(clauses, SAMPLE_RU)
    refs = {i["ref"] for i in issues}
    assert "п. 9.9" in refs


def test_validate_references_en_finds_dangling_ref():
    # Fixture references "see clause 6.2" but there is no clause 6.x.
    clauses = parse_contract(SAMPLE_EN)
    issues = validate_references(clauses, SAMPLE_EN)
    refs = {i["ref"] for i in issues}
    assert "Clause 6.2" in refs
