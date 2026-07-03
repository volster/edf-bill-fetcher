"""Phase 2.x — formula-injection guard on Excel text cells.

External text sources (PDF/PST/email bodies) are attacker-
controllable in principle: a user could receive a bill whose
``Details`` field starts with ``=cmd|'/c calc'!A1`` or whose
``Sender`` is ``+MALICIOUS_FORMULA``.  Excel will auto-evaluate
any cell value starting with ``=``, ``+``, ``-`` or ``@`` as a
formula when the workbook is opened, prompting a function-call
side-effect or a credential-leak formula chain.  This is real
for ombudsman submissions because they are passed to a third
party who opens the workbook.

The mitigation:

    1. ``_text`` coerces the cell value to ``str`` first so
       pandas-introduced types don't fall through.
    2. ``_text`` pins ``cell.data_type = 's'`` (text) so Excel
       never tries to auto-format the cell as a formula even
       when the workbook is opened with maximum-fidelity
       parsing.
    3. Belt-and-braces: a leading ``=``, ``+``, ``-`` or ``@``
       triggers an apostrophe-prefix so even LibreOffice
       (which sometimes ignores the data_type pin) renders the
       cell as text.

These tests pin all three guarantees.
"""

from __future__ import annotations

from openpyxl import Workbook

from edf_collector import _text


class TestFormulaInjectionGuard:
    """Phase 2.x — formula-injection guard.

    Excel auto-evaluation of a cell whose textual value starts
    with ``=``, ``+``, ``-`` or ``@`` is the standard
    formula-injection attack surface.  The guard pins:

        * data_type = 's' (text) on every text cell,
        * apostrophe-prefix on leading special chars,
        * benign strings render unchanged.
    """

    def test_data_type_is_text_pin(self) -> None:
        wb = Workbook()
        ws = wb.active
        _text(ws, 1, 1, "Anything here")
        cell = ws.cell(row=1, column=1)
        # The data_type pin is what flips Excel out of
        # auto-evaluate mode on most builds.
        assert cell.data_type == "s", f"Expected data_type='s' (text), got {cell.data_type!r}"

    def test_leading_equals_triggers_apostrophe_prefix(self) -> None:
        wb = Workbook()
        ws = wb.active
        _text(ws, 1, 1, "=cmd|'/c calc'!A1")
        cell = ws.cell(row=1, column=1)
        # Must not equal the raw ``=cmd...`` payload; we want
        # an apostrophe-prefix to defeat Excel's auto-format.
        assert cell.value == "'=cmd|'/c calc'!A1", (
            f"Expected apostrophe-prefix, got value={cell.value!r}"
        )
        assert cell.data_type == "s"

    def test_leading_plus_triggers_apostrophe_prefix(self) -> None:
        wb = Workbook()
        ws = wb.active
        _text(ws, 1, 1, "+MALICIOUS_FORMULA")
        cell = ws.cell(row=1, column=1)
        assert cell.value == "'+MALICIOUS_FORMULA"
        assert cell.data_type == "s"

    def test_leading_minus_triggers_apostrophe_prefix(self) -> None:
        wb = Workbook()
        ws = wb.active
        _text(ws, 1, 1, "-1+1")  # Excel would evaluate to 0
        cell = ws.cell(row=1, column=1)
        assert cell.value == "'-1+1", f"Leading minus must be guarded; got {cell.value!r}"
        assert cell.data_type == "s"

    def test_leading_at_triggers_apostrophe_prefix(self) -> None:
        wb = Workbook()
        ws = wb.active
        _text(ws, 1, 1, "@SUM(1+1)")
        cell = ws.cell(row=1, column=1)
        assert cell.value == "'@SUM(1+1)"
        assert cell.data_type == "s"

    def test_benign_strings_render_unchanged(self) -> None:
        # A non-special-char leading string should be preserved
        # *verbatim* — the guard mustn't add an apostrophe when
        # one isn't needed.
        wb = Workbook()
        ws = wb.active
        _text(ws, 1, 1, "Period To")
        cell = ws.cell(row=1, column=1)
        assert cell.value == "Period To", (
            f"Leading non-special char triggered apostrophe?  got {cell.value!r}"
        )
        assert cell.data_type == "s"

    def test_special_char_not_at_start_unchanged(self) -> None:
        # ``Config=foo`` has ``=`` mid-string, not leading.  The
        # guard should not append an apostrophe here because the
        # leading char isn't special.
        wb = Workbook()
        ws = wb.active
        _text(ws, 1, 1, "Config=foo")
        cell = ws.cell(row=1, column=1)
        assert cell.value == "Config=foo"
        assert cell.data_type == "s"

    def test_none_becomes_empty_string(self) -> None:
        # A None value used to render as the literal ``None``
        # in some legacy Excel saves (``openpyxl`` writes it as
        # an empty cell).  We pin the empty-string behaviour so
        # downstream consumers (the analyst) see ``""`` rather
        # than ``None``.
        wb = Workbook()
        ws = wb.active
        _text(ws, 1, 1, None)
        cell = ws.cell(row=1, column=1)
        assert cell.value == ""
        assert cell.data_type == "s"

    def test_non_string_values_are_coerced(self) -> None:
        # Realistic case: ``Details`` is sometimes a pandas
        # ``float`` (NaN) or ``int``.  The guard must coerce to
        # ``str`` first so ``data_type = 's'`` is set against a
        # textual cell, not a numeric cell.
        wb = Workbook()
        ws = wb.active
        _text(ws, 1, 1, 12345)
        cell = ws.cell(row=1, column=1)
        assert cell.value == "12345"
        assert cell.data_type == "s"

    def test_apostrophe_in_legitimate_value_stays_intact(self) -> None:
        # Edge case: a real Description field that already
        # starts with ``'`` (e.g., a debt collector's
        # bracketed-style address).  We don't want to contaminate
        # it further, just pin the data_type.
        wb = Workbook()
        ws = wb.active
        _text(ws, 1, 1, "Bob's Mobile")
        cell = ws.cell(row=1, column=1)
        assert cell.value == "Bob's Mobile"
        assert cell.data_type == "s"
