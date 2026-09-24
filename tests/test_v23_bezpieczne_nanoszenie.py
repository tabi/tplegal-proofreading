"""Regresje v2.3 — błędy znalezione w audycie 24.09.2026.

Każdy test odtwarza konkretny defekt v2.0: cichą zmianę poza tracked changes,
złe miejsce naniesienia albo rozjazd tekstu między extract-text a apply-corrections.
"""

import json
import os
import subprocess
import sys
import zipfile

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from apply_corrections import apply_docx  # noqa: E402
from extract_text import extract_text  # noqa: E402
from verify_docx import compare, verify  # noqa: E402

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
REPO = os.path.join(os.path.dirname(__file__), '..')


def _docx(path, body_xml, extra_members=None):
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body>{body_xml}</w:body></w:document>'
    )
    with zipfile.ZipFile(path, 'w') as zf:
        zf.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
        zf.writestr('word/document.xml', doc_xml)
        for name, data in (extra_members or {}).items():
            zf.writestr(name, data)
    return str(path)


def _doc_xml(path):
    with zipfile.ZipFile(path) as zf:
        return zf.read('word/document.xml').decode('utf-8')


def _apply(tmp_path, body, corrections, extra_members=None):
    src = _docx(tmp_path / 'in.docx', body, extra_members)
    out = str(tmp_path / 'out.docx')
    results = apply_docx(src, corrections, out, author='Test', date_str='2026-09-24T00:00:00Z')
    return src, out, results


def _statuses(results):
    return [r['status'] for r in results]


# ── 1. Elementy nietekstowe nie znikają po cichu ─────────────────────────

class TestNietekstoweElementy:
    def test_tabulator_w_tym_samym_runie_zostaje(self, tmp_path):
        body = '<w:p><w:r><w:t>Sad</w:t><w:tab/><w:t>umwoa jest</w:t></w:r></w:p>'
        src, out, res = _apply(tmp_path, body, [{'original': 'umwoa', 'corrected': 'umowa'}])
        assert _statuses(res) == ['applied']
        assert _doc_xml(out).count('<w:tab/>') == 1
        assert verify(src, out)['ok'] is True

    def test_zlamanie_wiersza_zostaje(self, tmp_path):
        body = '<w:p><w:r><w:t>linia jeden</w:t><w:br/><w:t>umwoa jest</w:t></w:r></w:p>'
        src, out, res = _apply(tmp_path, body, [{'original': 'umwoa', 'corrected': 'umowa'}])
        assert _statuses(res) == ['applied']
        assert _doc_xml(out).count('<w:br/>') == 1
        assert verify(src, out)['ok'] is True

    def test_przypis_wewnatrz_poprawianego_fragmentu_blokuje_korekte(self, tmp_path):
        body = (
            '<w:p><w:r><w:t>ta umwoa</w:t></w:r>'
            '<w:r><w:footnoteReference w:id="1"/></w:r>'
            '<w:r><w:t xml:space="preserve"> jst wazna</w:t></w:r></w:p>'
        )
        src, out, res = _apply(
            tmp_path, body, [{'original': 'umwoa[^1] jst', 'corrected': 'umowa jest'}],
        )
        assert _statuses(res) == ['rejected']
        assert 'nietekstowy' in res[0]['reason']
        assert 'footnoteReference' in _doc_xml(out)
        assert verify(src, out)['ok'] is True

    def test_korekta_obok_przypisu_zachowuje_przypis(self, tmp_path):
        body = (
            '<w:p><w:r><w:t>ta umwoa</w:t></w:r>'
            '<w:r><w:footnoteReference w:id="1"/></w:r>'
            '<w:r><w:t xml:space="preserve"> jest wazna</w:t></w:r></w:p>'
        )
        src, out, res = _apply(tmp_path, body, [{'original': 'ta umwoa', 'corrected': 'ta umowa'}])
        assert _statuses(res) == ['applied']
        assert _doc_xml(out).count('footnoteReference') == 1
        assert verify(src, out)['ok'] is True


# ── 2. Formatowanie nie przelewa się na sąsiedni tekst ───────────────────

class TestFormatowanie:
    def test_poprawka_w_pogrubionym_slowie_nie_pogrubia_reszty(self, tmp_path):
        body = (
            '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Powod</w:t></w:r>'
            '<w:r><w:t xml:space="preserve"> wnosi o umwoa</w:t></w:r></w:p>'
        )
        src, out, res = _apply(tmp_path, body, [
            {'original': 'Powod wnosi', 'corrected': 'Powód wnosi'},
            {'original': 'o umwoa', 'corrected': 'o umowa'},
        ])
        assert _statuses(res) == ['applied', 'applied']
        xml = _doc_xml(out)
        # bold tylko na 'Powod' (del) i 'Powód' (ins) — nigdzie więcej
        assert xml.count('<w:b/>') == 2
        assert verify(src, out)['ok'] is True

    def test_fragment_o_mieszanym_formatowaniu_odrzucony(self, tmp_path):
        body = (
            '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Pow</w:t></w:r>'
            '<w:r><w:t xml:space="preserve">od wnosi</w:t></w:r></w:p>'
        )
        src, out, res = _apply(tmp_path, body, [{'original': 'Powod wnosi', 'corrected': 'Powód wnosi'}])
        assert _statuses(res) == ['rejected']
        assert 'formatowanie' in res[0]['reason']
        assert verify(src, out)['ok'] is True


# ── 3. Miejsce naniesienia jest jednoznaczne ─────────────────────────────

class TestJednoznacznosc:
    BODY = (
        '<w:p><w:r><w:t>Pozwany ktory zalega.</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>Powod ktory wnosi.</w:t></w:r></w:p>'
    )

    def test_wiele_wystapien_bez_akapitu_odrzucone(self, tmp_path):
        src, out, res = _apply(tmp_path, self.BODY, [{'original': 'ktory', 'corrected': 'który'}])
        assert _statuses(res) == ['rejected']
        assert '¶001' in res[0]['reason'] and '¶002' in res[0]['reason']

    def test_wskazany_akapit_trafia_we_wlasciwe_miejsce(self, tmp_path):
        src, out, res = _apply(
            tmp_path, self.BODY, [{'original': 'ktory', 'corrected': 'który', 'paragraph': 2}],
        )
        assert _statuses(res) == ['applied']
        assert res[0]['paragraph'] == 2
        changes = verify(src, out)['changes']
        assert [c['paragraph'] for c in changes] == [2]

    def test_akapit_jako_napis_z_pilcrowem(self, tmp_path):
        src, out, res = _apply(
            tmp_path, self.BODY, [{'original': 'ktory', 'corrected': 'który', 'paragraph': '¶002'}],
        )
        assert _statuses(res) == ['applied']

    def test_brak_fallbacku_na_wielkosc_liter(self, tmp_path):
        body = '<w:p><w:r><w:t>Wielka Litera</w:t></w:r></w:p>'
        src, out, res = _apply(tmp_path, body, [{'original': 'wielka litera', 'corrected': 'mała litera'}])
        assert _statuses(res) == ['not_found']


# ── 4. extract-text i apply-corrections widzą ten sam tekst ──────────────

class TestSpojnoscTekstu:
    def test_drugi_przebieg_na_pliku_z_tracked_changes(self, tmp_path):
        body = '<w:p><w:r><w:t>Powod wnosi o zasadzenie kwoty od pozwanego ktory zalega</w:t></w:r></w:p>'
        src, out1, res1 = _apply(tmp_path, body, [{'original': 'o zasadzenie', 'corrected': 'o zasądzenie'}])
        assert _statuses(res1) == ['applied']
        line = extract_text(out1)[0]
        assert 'zasądzenie kwoty od pozwanego ktory' in line
        out2 = str(tmp_path / 'out2.docx')
        res2 = apply_docx(
            out1, [{'original': 'zasądzenie kwoty od pozwanego ktory', 'corrected': 'zasądzenie kwoty od pozwanego, który'}],
            out2, author='Test', date_str='2026-09-24T00:00:00Z',
        )
        assert _statuses(res2) == ['applied']
        assert 'pozwanego, który zalega' in extract_text(out2)[0]
        report = verify(src, out2)
        assert report['ok'] is True
        assert len(report['changes']) == 3

    def test_poprawka_wewnatrz_cudzej_zmiany_odrzucona(self, tmp_path):
        body = (
            '<w:p><w:r><w:t xml:space="preserve">Powod wnosi o </w:t></w:r>'
            '<w:ins w:id="7" w:author="Bartek" w:date="2026-01-01T00:00:00Z">'
            '<w:r><w:t>zasadzenie</w:t></w:r></w:ins>'
            '<w:r><w:t xml:space="preserve"> kwoty</w:t></w:r></w:p>'
        )
        src, out, res = _apply(tmp_path, body, [{'original': 'o zasadzenie', 'corrected': 'o zasądzenie'}])
        assert _statuses(res) == ['rejected']
        assert 'śledzon' in res[0]['reason']

    def test_hiperlink_widoczny_dla_obu_narzedzi(self, tmp_path):
        body = (
            '<w:p><w:r><w:t xml:space="preserve">Kontakt: </w:t></w:r>'
            '<w:hyperlink><w:r><w:t>biuro@tabert.pl</w:t></w:r></w:hyperlink>'
            '<w:r><w:t xml:space="preserve"> lub telefonicznie ktory</w:t></w:r></w:p>'
        )
        assert extract_text(_docx(tmp_path / 'x.docx', body))[0] == '¶001: Kontakt: biuro@tabert.pl lub telefonicznie ktory'
        src, out, res = _apply(tmp_path, body, [
            {'original': 'biuro@tabert.pl lub telefonicznie ktory', 'corrected': 'biuro@tabert.pl lub telefonicznie, który'},
        ])
        assert _statuses(res) == ['applied']
        assert verify(src, out)['ok'] is True

    def test_tabulator_widoczny_w_extract_text(self, tmp_path):
        body = '<w:p><w:r><w:t>Sad</w:t><w:tab/><w:t>Rejonowy</w:t></w:r></w:p>'
        assert extract_text(_docx(tmp_path / 'x.docx', body))[0] == '¶001: Sad\tRejonowy'


# ── 5. Minimalna zmiana i blokady treści chronionych ─────────────────────

class TestZakresZmiany:
    def test_sam_przecinek_to_samo_wstawienie(self, tmp_path):
        body = '<w:p><w:r><w:t>zalezy od tego czy przyjdzie</w:t></w:r></w:p>'
        src, out, res = _apply(tmp_path, body, [{'original': 'od tego czy', 'corrected': 'od tego, czy'}])
        assert _statuses(res) == ['applied']
        xml = _doc_xml(out)
        assert '<w:del ' not in xml
        assert xml.count('<w:ins ') == 1
        assert res[0]['deleted'] == '' and res[0]['inserted'] == ','

    def test_zmiana_cyfr_odrzucona(self, tmp_path):
        body = '<w:p><w:r><w:t>sygn. akt V GNc 808/26 z dnia</w:t></w:r></w:p>'
        src, out, res = _apply(tmp_path, body, [{'original': 'GNc 808/26', 'corrected': 'GNc 808/25'}])
        assert _statuses(res) == ['rejected']
        assert 'cyfr' in res[0]['reason']

    def test_ampersand_w_nazwie_odrzucony(self, tmp_path):
        body = '<w:p><w:r><w:t>spółka Janus&amp;Janus EU</w:t></w:r></w:p>'
        src, out, res = _apply(tmp_path, body, [{'original': 'Janus&Janus', 'corrected': 'Janus & Janus'}])
        assert _statuses(res) == ['rejected']

    def test_pole_formularza_zablokowane(self, tmp_path):
        body = (
            '<w:p><w:r><w:t xml:space="preserve">Strona </w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText> DOCPROPERTY Tytul </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>umwoa</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        )
        src, out, res = _apply(tmp_path, body, [{'original': 'umwoa', 'corrected': 'umowa'}])
        assert _statuses(res) == ['rejected']
        assert 'pol' in res[0]['reason']


# ── 6. verify-docx łapie każdą cichą zmianę i listuje wszystkie zmiany ───

class TestVerify:
    def _pair(self, tmp_path, a, b, extra_a=None, extra_b=None):
        return (
            _docx(tmp_path / 'a.docx', a, extra_a),
            _docx(tmp_path / 'b.docx', b, extra_b),
        )

    def test_cicha_zmiana_tekstu(self, tmp_path):
        a, b = self._pair(tmp_path, '<w:p><w:r><w:t>Janus&amp;Janus</w:t></w:r></w:p>',
                          '<w:p><w:r><w:t>Janus &amp; Janus</w:t></w:r></w:p>')
        assert verify(a, b)['ok'] is False
        assert compare(a, b, quiet=True) == 2

    def test_zgubiony_tabulator(self, tmp_path):
        a, b = self._pair(tmp_path, '<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r></w:p>',
                          '<w:p><w:r><w:t>A</w:t><w:t>B</w:t></w:r></w:p>')
        assert verify(a, b)['ok'] is False

    def test_zmienione_pogrubienie(self, tmp_path):
        a, b = self._pair(tmp_path, '<w:p><w:r><w:t>Tekst</w:t></w:r></w:p>',
                          '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Tekst</w:t></w:r></w:p>')
        report = verify(a, b)
        assert report['ok'] is False
        assert any('formatowan' in e for e in report['errors'])

    def test_same_rsid_to_nie_blad(self, tmp_path):
        a, b = self._pair(tmp_path, '<w:p w:rsidR="00AA"><w:r w:rsidR="00AB"><w:t>Tekst</w:t></w:r></w:p>',
                          '<w:p w:rsidR="00CC"><w:r><w:t>Tekst</w:t></w:r></w:p>')
        assert verify(a, b)['ok'] is True

    def test_zmieniony_inny_plik_w_zip(self, tmp_path):
        body = '<w:p><w:r><w:t>Tekst</w:t></w:r></w:p>'
        a, b = self._pair(tmp_path, body, body,
                          {'word/footnotes.xml': '<a>1</a>'}, {'word/footnotes.xml': '<a>2</a>'})
        report = verify(a, b)
        assert report['ok'] is False
        assert any('word/footnotes.xml' in e for e in report['errors'])

    def test_lista_zmian_kompletna(self, tmp_path):
        body = '<w:p><w:r><w:t>Ala ma kta i psa ktory szczeka.</w:t></w:r></w:p>'
        src, out, res = _apply(tmp_path, body, [
            {'original': 'kta', 'corrected': 'kota'},
            {'original': 'psa ktory', 'corrected': 'psa, który'},
        ])
        report = verify(src, out)
        assert report['ok'] is True
        assert [(c['deleted'], c['inserted']) for c in report['changes']] == [
            ('kta', 'kota'), ('', ','), ('ktory', 'który'),
        ]

    def test_cli_przyjmuje_strict_text_i_wypisuje_liste(self, tmp_path):
        src, out, _ = _apply(tmp_path, '<w:p><w:r><w:t>Ala ma kta.</w:t></w:r></w:p>',
                             [{'original': 'kta', 'corrected': 'kota'}])
        r = subprocess.run([sys.executable, '-m', 'verify_docx', src, out, '--strict-text'],
                           capture_output=True, text=True, cwd=REPO)
        assert r.returncode == 0, r.stdout + r.stderr
        assert 'kta' in r.stdout and 'kota' in r.stdout
        assert 'ZMIANY W DOKUMENCIE: 1' in r.stdout


# ── 7. CLI apply-corrections — raport i kody wyjścia ─────────────────────

class TestApplyCLI:
    def test_raport_wymienia_odrzucone_i_kod_1(self, tmp_path):
        src = _docx(tmp_path / 'in.docx', '<w:p><w:r><w:t>Ala ma kta. Nr 12.</w:t></w:r></w:p>')
        corr = tmp_path / 'c.json'
        corr.write_text(json.dumps([
            {'original': 'kta', 'corrected': 'kota', 'note': 'literówka'},
            {'original': 'Nr 12', 'corrected': 'Nr 13', 'note': 'zła liczba'},
            {'original': 'nie ma', 'corrected': 'jest', 'note': 'x'},
        ]), encoding='utf-8')
        out = str(tmp_path / 'out.docx')
        r = subprocess.run([sys.executable, '-m', 'apply_corrections', src, str(corr), '-o', out],
                           capture_output=True, text=True, cwd=REPO)
        assert r.returncode == 1
        assert 'NANIESIONO: 1 z 3' in r.stdout
        assert 'ODRZUCONO' in r.stdout and 'NIE ZNALEZIONO' in r.stdout

    def test_niepoprawny_json_kod_2(self, tmp_path):
        src = _docx(tmp_path / 'in.docx', '<w:p><w:r><w:t>Ala.</w:t></w:r></w:p>')
        corr = tmp_path / 'c.json'
        corr.write_text(json.dumps([{'corrected': 'x'}]), encoding='utf-8')
        r = subprocess.run([sys.executable, '-m', 'apply_corrections', src, str(corr), '-o', str(tmp_path / 'o.docx')],
                           capture_output=True, text=True, cwd=REPO)
        assert r.returncode == 2
        assert not os.path.exists(tmp_path / 'o.docx')


class TestDodatkoweBlokady:
    def test_sledzona_zmiana_formatowania_zablokowana(self, tmp_path):
        body = (
            '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="5" w:author="X" w:date="2026-01-01T00:00:00Z">'
            '<w:rPr/></w:rPrChange></w:rPr><w:t>umwoa</w:t></w:r></w:p>'
        )
        src, out, res = _apply(tmp_path, body, [{'original': 'umwoa', 'corrected': 'umowa'}])
        assert _statuses(res) == ['rejected']
        assert 'rPrChange' in res[0]['reason']

    def test_ostrzezenie_przy_wyrazie_z_wielkiej_litery(self, tmp_path):
        body = '<w:p><w:r><w:t>przed Sadem Rejonowym</w:t></w:r></w:p>'
        src, out, res = _apply(tmp_path, body, [{'original': 'Sadem', 'corrected': 'Sądem'}])
        assert _statuses(res) == ['applied']
        assert res[0]['warnings']

    def test_verify_lapie_zdublowane_id(self, tmp_path):
        a = _docx(tmp_path / 'a.docx', '<w:p><w:r><w:t>Ala ma kota.</w:t></w:r></w:p>')
        b = _docx(tmp_path / 'b.docx', (
            '<w:p><w:r><w:t xml:space="preserve">Ala </w:t></w:r>'
            '<w:ins w:id="9" w:author="K" w:date="D"><w:r><w:t>x</w:t></w:r></w:ins>'
            '<w:ins w:id="9" w:author="K" w:date="D"><w:r><w:t>y</w:t></w:r></w:ins>'
            '<w:r><w:t>ma kota.</w:t></w:r></w:p>'
        ))
        report = verify(a, b)
        assert report['ok'] is False
        assert any('zdublowane' in e for e in report['errors'])

    def test_verify_nie_pomija_zmian_z_poprzednich_przebiegow_wejscia(self, tmp_path):
        body = (
            '<w:p><w:r><w:t xml:space="preserve">Ala </w:t></w:r>'
            '<w:ins w:id="3" w:author="Bartek" w:date="D"><w:r><w:t>ma</w:t></w:r></w:ins>'
            '<w:r><w:t xml:space="preserve"> kta.</w:t></w:r></w:p>'
        )
        src, out, res = _apply(tmp_path, body, [{'original': 'kta', 'corrected': 'kota'}])
        report = verify(src, out)
        assert report['ok'] is True
        assert report['preexisting'] == 1
        assert [(c['deleted'], c['inserted']) for c in report['changes']] == [('kta', 'kota')]


@pytest.mark.parametrize('body', [
    '<w:p><w:r><w:t>Ala ma kta.</w:t></w:r></w:p>',
    '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Komorka umwoa.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
])
def test_pusta_lista_korekt_daje_identyczny_dokument(tmp_path, body):
    src, out, res = _apply(tmp_path, body, [])
    assert res == []
    report = verify(src, out)
    assert report['ok'] is True and report['changes'] == []
