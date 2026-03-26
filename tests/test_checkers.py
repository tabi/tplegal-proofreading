"""Tests for checkers.py — pluggable text checker architecture."""

import os
import sys
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from checkers import (
    _dedup_overlapping,
    LanguageToolChecker,
    RegexChecker,
    CompositeChecker,
    polish_legal_rules,
)


# ── _dedup_overlapping ────────────────────────────────────────────────────

class TestDedupOverlapping:
    def test_no_overlap(self):
        matches = [
            {'offset': 0, 'end': 3},
            {'offset': 5, 'end': 8},
        ]
        assert len(_dedup_overlapping(matches)) == 2

    def test_overlap_keeps_earlier(self):
        matches = [
            {'offset': 0, 'end': 5},
            {'offset': 3, 'end': 7},
        ]
        result = _dedup_overlapping(matches)
        assert len(result) == 1
        assert result[0]['offset'] == 0

    def test_empty(self):
        assert _dedup_overlapping([]) == []

    def test_unsorted_input(self):
        matches = [
            {'offset': 10, 'end': 15},
            {'offset': 0, 'end': 3},
        ]
        result = _dedup_overlapping(matches)
        assert result[0]['offset'] == 0


# ── LanguageToolChecker ───────────────────────────────────────────────────

class TestLanguageToolChecker:
    def test_empty_text(self):
        checker = LanguageToolChecker()
        assert checker.check('') == []
        assert checker.check('  ') == []

    @patch('requests.post')
    def test_parses_response(self, mock_post):
        mock_resp = MagicMock()
        mock_resp.json.return_value = {
            'matches': [{
                'offset': 0, 'length': 3,
                'message': 'Typo',
                'replacements': [{'value': 'fix'}],
                'rule': {'id': 'R1', 'category': {'id': 'SPELLING'}},
            }]
        }
        mock_resp.raise_for_status = MagicMock()
        mock_post.return_value = mock_resp

        checker = LanguageToolChecker(url='http://test', language='pl-PL')
        matches = checker.check('foo bar')
        assert len(matches) == 1
        assert matches[0]['replacement'] == 'fix'

    @patch('requests.post')
    def test_api_error(self, mock_post):
        mock_post.side_effect = Exception("fail")
        checker = LanguageToolChecker()
        assert checker.check('test') == []

    def test_config_from_env(self):
        with patch.dict(os.environ, {'LANGUAGETOOL_URL': 'http://custom'}):
            checker = LanguageToolChecker()
            assert checker.url == 'http://custom'


# ── RegexChecker ──────────────────────────────────────────────────────────

class TestRegexChecker:
    def test_simple_rule(self):
        checker = RegexChecker()
        checker.add_rule(r'  +', ' ', 'Double space')
        matches = checker.check('Ala  ma   kota.')
        assert len(matches) == 2
        assert matches[0]['replacement'] == ' '

    def test_no_match(self):
        checker = RegexChecker()
        checker.add_rule(r'xyz', 'abc', 'test')
        assert checker.check('Ala ma kota.') == []

    def test_empty_text(self):
        checker = RegexChecker()
        checker.add_rule(r'a', 'b', 'test')
        assert checker.check('') == []

    def test_skips_noop_replacement(self):
        """If replacement equals matched text, skip it."""
        checker = RegexChecker()
        checker.add_rule(r'abc', 'abc', 'noop')
        assert checker.check('abc') == []

    def test_group_replacement(self):
        checker = RegexChecker()
        checker.add_rule(r'art\.(\d)', r'art. \1', 'Space after art.')
        matches = checker.check('na podstawie art.5')
        assert len(matches) == 1
        assert matches[0]['replacement'] == 'art. 5'

    def test_compiled_pattern(self):
        import re
        checker = RegexChecker()
        checker.add_rule(re.compile(r'\bw/w\b'), 'ww.', 'Niestandardowy skrót')
        matches = checker.check('w/w firma')
        assert len(matches) == 1
        assert matches[0]['replacement'] == 'ww.'


# ── CompositeChecker ──────────────────────────────────────────────────────

class TestCompositeChecker:
    def test_combines_checkers(self):
        c1 = RegexChecker()
        c1.add_rule(r'  +', ' ', 'double space')
        c2 = RegexChecker()
        c2.add_rule(r'\bw/w\b', 'ww.', 'abbreviation')

        composite = CompositeChecker([c1, c2])
        matches = composite.check('w/w  firma')
        assert len(matches) == 2

    def test_deduplicates_overlapping(self):
        """If two checkers match the same range, keep only the first."""
        c1 = RegexChecker()
        c1.add_rule(r'abc', 'x', 'rule1')
        c2 = RegexChecker()
        c2.add_rule(r'abc', 'y', 'rule2')

        composite = CompositeChecker([c1, c2])
        matches = composite.check('abc')
        assert len(matches) == 1

    def test_add_returns_self(self):
        composite = CompositeChecker()
        result = composite.add(RegexChecker())
        assert result is composite

    def test_empty(self):
        assert CompositeChecker().check('text') == []


# ── polish_legal_rules ────────────────────────────────────────────────────

class TestPolishLegalRules:
    def test_double_space(self):
        checker = polish_legal_rules()
        matches = checker.check('Ala  ma kota.')
        assert any(m['category'] == 'WHITESPACE' for m in matches)

    def test_ww_abbreviation(self):
        checker = polish_legal_rules()
        matches = checker.check('w/w przedmiot')
        assert any(m['replacement'] == 'ww.' for m in matches)

    def test_art_space(self):
        checker = polish_legal_rules()
        matches = checker.check('na podstawie art.5 ustawy')
        assert any('art. 5' in m['replacement'] for m in matches)
