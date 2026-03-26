"""Text checkers for proofreading — pluggable architecture.

Each checker implements `check(text) -> list[Match]` where Match is a dict with:
    offset, length, end, replacement, message, rule_id, category

Checkers can be combined via CompositeChecker.
"""

import logging
import os
import re
from abc import ABC, abstractmethod

log = logging.getLogger(__name__)

# Match dict type (for documentation — we keep it as plain dicts for simplicity)
# Match = {"offset": int, "length": int, "end": int, "replacement": str,
#          "message": str, "rule_id": str, "category": str}


def _dedup_overlapping(matches):
    """Remove overlapping matches, keeping the earlier one."""
    matches.sort(key=lambda m: m['offset'])
    filtered = []
    last_end = -1
    for m in matches:
        if m['offset'] >= last_end:
            filtered.append(m)
            last_end = m['end']
    return filtered


class BaseChecker(ABC):
    """Abstract base for text checkers."""

    @abstractmethod
    def check(self, text):
        """Check text and return a list of match dicts."""


class LanguageToolChecker(BaseChecker):
    """LanguageTool API checker."""

    def __init__(self, url=None, language=None, request_delay=None):
        self.url = url or os.environ.get('LANGUAGETOOL_URL', 'https://api.languagetool.org/v2/check')
        self.language = language or os.environ.get('PROOFREAD_LANGUAGE', 'pl-PL')
        self.request_delay = request_delay or float(os.environ.get('PROOFREAD_REQUEST_DELAY', '0.3'))

    def check(self, text):
        if not text.strip():
            return []

        import requests

        data = {
            'text': text,
            'language': self.language,
            'enabledOnly': 'false',
        }

        try:
            resp = requests.post(self.url, data=data, timeout=30)
            resp.raise_for_status()
            result = resp.json()
        except Exception as e:
            log.warning("LanguageTool API error: %s", e)
            return []

        matches = []
        for m in result.get('matches', []):
            replacements = [r['value'] for r in m.get('replacements', []) if r.get('value')]
            if not replacements:
                continue

            matches.append({
                'offset': m['offset'],
                'length': m['length'],
                'end': m['offset'] + m['length'],
                'message': m.get('message', ''),
                'replacement': replacements[0],
                'rule_id': m.get('rule', {}).get('id', ''),
                'category': m.get('rule', {}).get('category', {}).get('id', ''),
            })

        return _dedup_overlapping(matches)


class RegexChecker(BaseChecker):
    """Rule-based checker using regex patterns.

    Rules are (pattern, replacement, message, category) tuples.
    The pattern should match the text to be replaced.
    """

    def __init__(self, rules=None):
        self.rules = rules or []

    def add_rule(self, pattern, replacement, message='', category='CUSTOM'):
        """Add a regex-based rule.

        Args:
            pattern: Regex pattern (str or compiled) matching text to replace.
            replacement: Replacement string (can use \\1 etc. for groups).
            message: Human-readable description of the issue.
            category: Category label for the match.
        """
        if isinstance(pattern, str):
            pattern = re.compile(pattern)
        self.rules.append((pattern, replacement, message, category))

    def check(self, text):
        if not text.strip():
            return []

        matches = []
        for pattern, replacement, message, category in self.rules:
            for m in pattern.finditer(text):
                expanded = m.expand(replacement)
                if expanded == m.group(0):
                    continue
                matches.append({
                    'offset': m.start(),
                    'length': m.end() - m.start(),
                    'end': m.end(),
                    'replacement': expanded,
                    'message': message,
                    'rule_id': pattern.pattern[:40],
                    'category': category,
                })

        return _dedup_overlapping(matches)


class CompositeChecker(BaseChecker):
    """Combines multiple checkers, deduplicating overlapping matches."""

    def __init__(self, checkers=None):
        self.checkers = checkers or []

    def add(self, checker):
        self.checkers.append(checker)
        return self

    def check(self, text):
        all_matches = []
        for checker in self.checkers:
            all_matches.extend(checker.check(text))
        return _dedup_overlapping(all_matches)


# ── Pre-built rule sets ───────────────────────────────────────────────────

def polish_legal_rules():
    """Common Polish legal writing rules."""
    checker = RegexChecker()

    # Double spaces
    checker.add_rule(r'  +', ' ', 'Podwójna spacja', 'WHITESPACE')

    # Comma before "który/która/które"
    checker.add_rule(
        r'(?<=[a-ząćęłńóśźż]) (któr[yaeąęiou]\w*)',
        r', \1',
        'Brak przecinka przed zaimkiem względnym',
        'PUNCTUATION',
    )

    # "w/w" → "wyżej wymieniony" or similar (informal abbreviation)
    checker.add_rule(
        r'\bw/w\b',
        'ww.',
        'Niestandardowy skrót — użyj "ww."',
        'STYLE',
    )

    # "na podstawie art." with missing space
    checker.add_rule(
        r'\bart\.(\d)',
        r'art. \1',
        'Brak spacji po "art."',
        'FORMATTING',
    )

    return checker
