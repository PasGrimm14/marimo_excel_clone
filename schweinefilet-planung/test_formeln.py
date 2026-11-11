"""Tests für formeln.py: Formel-Auswerter, Validierung und Berechnungslogik."""

from __future__ import annotations

from datetime import date

import pytest

import formeln as f


# ---------------------------------------------------------------------------
# Normalisierung deutscher Excel-Schreibweise
# ---------------------------------------------------------------------------


def test_normalisiere_semikolon_wird_zu_komma():
    assert f.normalisiere_formel("MAX(1; 2)") == "MAX(1, 2)"


def test_normalisiere_dezimalkomma_zwischen_ziffern():
    assert f.normalisiere_formel("0,9") == "0.9"
    assert f.normalisiere_formel("12,345") == "12.345"


def test_normalisiere_komma_als_trenner_bleibt_komma():
    # Komma zwischen "VJ" (kein Ziffer) und "0" (Ziffer) ist ein Trenner, kein Dezimalkomma.
    ergebnis = f.normalisiere_formel("MAX(Faktor_VJ, 0,9)")
    assert ergebnis == "MAX(Faktor_VJ, 0.9)"


def test_normalisiere_gemischte_schreibweise_begrenzter_faktor():
    ergebnis = f.normalisiere_formel("Bestellmenge * MIN(MAX(Faktor_VJ; 0,9); 1,3)")
    assert ergebnis == "Bestellmenge * MIN(MAX(Faktor_VJ, 0.9), 1.3)"


def test_normalisiere_gleichheitszeichen():
    ergebnis = f.normalisiere_formel('Wochentag = "Sa"')
    assert ergebnis == 'Wochentag == "Sa"'


def test_normalisiere_ungleichheitszeichen():
    assert f.normalisiere_formel("Wochentag <> \"Sa\"") == 'Wochentag != "Sa"'


def test_normalisiere_laesst_doppelte_vergleichsoperatoren_unveraendert():
    assert f.normalisiere_formel("Faktor_VJ >= 1") == "Faktor_VJ >= 1"
    assert f.normalisiere_formel("Faktor_VJ <= 1") == "Faktor_VJ <= 1"


# ---------------------------------------------------------------------------
# Validierung
# ---------------------------------------------------------------------------


def test_validiere_leere_formel():
    with pytest.raises(f.FormelFehler, match="leer"):
        f.validiere_formel("")


def test_validiere_offene_klammer():
    with pytest.raises(f.FormelFehler, match="Klammer"):
        f.validiere_formel("Bestellmenge * (1 + Zuschlag")


def test_validiere_unbekannte_variable_mit_vorschlag():
    with pytest.raises(f.FormelFehler, match="Meinten Sie 'Bestellmenge'"):
        f.validiere_formel("Bestelmenge * 2")


def test_validiere_unbekannte_funktion_mit_vorschlag():
    with pytest.raises(f.FormelFehler, match="Unbekannte Funktion"):
        f.validiere_formel("MXX(Bestellmenge; 1)")


@pytest.mark.parametrize(
    "ausdruck",
    [
        "__import__('os').system('echo hi')",
        "(1).__class__",
        "[x for x in range(10)]",
        "lambda x: x",
        "Bestellmenge.bit_length()",
        "Bestellmenge[0]",
        "open('x.txt')",
        "exec('1')",
        "eval('1')",
    ],
)
def test_validiere_lehnt_unerlaubte_ausdruecke_ab(ausdruck):
    with pytest.raises(f.FormelFehler):
        f.validiere_formel(ausdruck)


def test_werte_formel_aus_lehnt_unerlaubte_ausdruecke_auch_bei_auswertung_ab():
    with pytest.raises(f.FormelFehler):
        f.werte_formel_aus("__import__('os')", f.beispielwerte())


# ---------------------------------------------------------------------------
# Auswertung der Standardformeln
# ---------------------------------------------------------------------------


def test_standardformel_bei_voller_gewichtung_entspricht_vorjahresfaktor():
    werte = {
        "Bestellmenge": 100.0,
        "Bestellmenge_VJ": 100.0,
        "Ist_VJ": 110.0,
        "Faktor_VJ": 1.1,
        "Gewichtung": 1.0,
        "Zuschlag": 0.0,
        "Wochentag": "Mo",
        "Tage_bis_Heiligabend": 5,
    }
    formel = f.formel_by_name(f.STANDARD_FORMEL_NAME)
    ergebnis = f.werte_formel_aus(formel.formel, werte)
    assert ergebnis == pytest.approx(110.0)


def test_standardformel_bei_gewichtung_null_entspricht_bestellmenge():
    werte = {
        "Bestellmenge": 100.0,
        "Bestellmenge_VJ": 100.0,
        "Ist_VJ": 200.0,
        "Faktor_VJ": 2.0,
        "Gewichtung": 0.0,
        "Zuschlag": 0.0,
        "Wochentag": "Mo",
        "Tage_bis_Heiligabend": 5,
    }
    formel = f.formel_by_name(f.STANDARD_FORMEL_NAME)
    ergebnis = f.werte_formel_aus(formel.formel, werte)
    assert ergebnis == pytest.approx(100.0)


def test_standardformel_mit_zuschlag():
    werte = {
        "Bestellmenge": 100.0,
        "Bestellmenge_VJ": 100.0,
        "Ist_VJ": 100.0,
        "Faktor_VJ": 1.0,
        "Gewichtung": 0.5,
        "Zuschlag": 0.1,
        "Wochentag": "Mo",
        "Tage_bis_Heiligabend": 5,
    }
    formel = f.formel_by_name(f.STANDARD_FORMEL_NAME)
    ergebnis = f.werte_formel_aus(formel.formel, werte)
    assert ergebnis == pytest.approx(110.0)


def test_maximum_formel():
    formel = f.formel_by_name("Maximum aus Bestellung und Vorjahr")
    werte = dict(f.beispielwerte())
    werte["Bestellmenge"] = 100.0
    werte["Faktor_VJ"] = 0.8
    ergebnis = f.werte_formel_aus(formel.formel, werte)
    assert ergebnis == pytest.approx(100.0)  # MAX(100, 80) = 100

    werte["Faktor_VJ"] = 1.5
    ergebnis = f.werte_formel_aus(formel.formel, werte)
    assert ergebnis == pytest.approx(150.0)


def test_begrenzter_faktor_formel_kappt_nach_oben_und_unten():
    formel = f.formel_by_name("Begrenzter Faktor")
    werte = dict(f.beispielwerte())
    werte["Bestellmenge"] = 100.0

    werte["Faktor_VJ"] = 5.0  # weit über der Obergrenze 1,3
    assert f.werte_formel_aus(formel.formel, werte) == pytest.approx(130.0)

    werte["Faktor_VJ"] = 0.1  # weit unter der Untergrenze 0,9
    assert f.werte_formel_aus(formel.formel, werte) == pytest.approx(90.0)


def test_wochentagsabhaengige_formel():
    formel = f.formel_by_name("Wochentagsabhängig")
    werte = dict(f.beispielwerte())
    werte["Bestellmenge"] = 100.0
    werte["Faktor_VJ"] = 1.2

    werte["Wochentag"] = "Sa"
    assert f.werte_formel_aus(formel.formel, werte) == pytest.approx(115.0)

    werte["Wochentag"] = "Mo"
    assert f.werte_formel_aus(formel.formel, werte) == pytest.approx(120.0)


# ---------------------------------------------------------------------------
# Funktionen: RUNDEN / AUFRUNDEN / ABRUNDEN / WENN / MITTELWERT
# ---------------------------------------------------------------------------


def test_runden_kaufmaennisch():
    assert f.runden(12.345, 2) == pytest.approx(12.35)
    assert f.runden(12.344, 2) == pytest.approx(12.34)
    assert f.runden(-2.5, 0) == pytest.approx(-3.0)


def test_aufrunden_immer_vom_nullpunkt_weg():
    assert f.aufrunden(12.01, 0) == pytest.approx(13.0)
    assert f.aufrunden(-12.01, 0) == pytest.approx(-13.0)


def test_abrunden_immer_zum_nullpunkt_hin():
    assert f.abrunden(12.99, 0) == pytest.approx(12.0)
    assert f.abrunden(-12.99, 0) == pytest.approx(-12.0)


def test_mittelwert():
    assert f.mittelwert(1, 2, 3) == pytest.approx(2.0)


def test_wenn_funktion():
    assert f.wenn(True, 1, 2) == 1
    assert f.wenn(False, 1, 2) == 2


# ---------------------------------------------------------------------------
# Rundung auf Gebindegröße
# ---------------------------------------------------------------------------


def test_runde_auf_gebinde_10kg():
    assert f.runde_auf_gebinde(104.0, 10) == 100.0
    assert f.runde_auf_gebinde(106.0, 10) == 110.0


def test_runde_auf_gebinde_keine_rundung_bei_0():
    assert f.runde_auf_gebinde(104.3, 0) == 104.3


# ---------------------------------------------------------------------------
# parse_zahl
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "eingabe,erwartet",
    [
        ("123", 123.0),
        ("123,45", 123.45),
        ("123.45", 123.45),
        ("1.234,56", 1234.56),
        ("1,234.56", 1234.56),
        (123, 123.0),
        (123.45, 123.45),
    ],
)
def test_parse_zahl_formate(eingabe, erwartet):
    assert f.parse_zahl(eingabe) == pytest.approx(erwartet)


@pytest.mark.parametrize("eingabe", [None, "", "abc", float("nan")])
def test_parse_zahl_ungueltig(eingabe):
    with pytest.raises(ValueError):
        f.parse_zahl(eingabe)


# ---------------------------------------------------------------------------
# Vorjahresvergleich: Wochentag + Abstand zu Heiligabend
# ---------------------------------------------------------------------------


def test_tage_bis_heiligabend():
    assert f.tage_bis_heiligabend(date(2025, 12, 24)) == 0
    assert f.tage_bis_heiligabend(date(2025, 12, 20)) == 4
    assert f.tage_bis_heiligabend(date(2025, 12, 26)) == -2


def test_finde_vorjahresvergleichstag_waehlt_gleichen_wochentag_und_naehesten_abstand():
    historie = [
        f.HistorienTag(date(2025, 12, 12), "Fr", 100, 110),  # 12 Tage vor Heiligabend
        f.HistorienTag(date(2025, 12, 19), "Fr", 200, 220),  # 5 Tage vor Heiligabend
        f.HistorienTag(date(2025, 12, 20), "Sa", 300, 330),  # 4 Tage vor Heiligabend
    ]
    treffer = f.finde_vorjahresvergleichstag("Fr", 6, historie)
    assert treffer is not None
    assert treffer.bestellmenge == 200  # näher an 6 Tagen als der Fr mit 12 Tagen Abstand


def test_finde_vorjahresvergleichstag_ohne_treffer():
    historie = [f.HistorienTag(date(2025, 12, 20), "Sa", 300, 330)]
    assert f.finde_vorjahresvergleichstag("Mo", 5, historie) is None


# ---------------------------------------------------------------------------
# berechne_tag / berechne_woche / berechne_wochensumme
# ---------------------------------------------------------------------------


def test_berechne_tag_grundfall():
    ergebnis = f.berechne_tag(
        datum=date(2026, 12, 18),
        wochentag="Fr",
        bestellmenge=100.0,
        bestellmenge_vj=100.0,
        ist_vj=110.0,
        gewichtung=1.0,
        zuschlag=0.0,
        formel_text=f.formel_by_name(f.STANDARD_FORMEL_NAME).formel,
        gebindegroesse=0,
        tage_bis_heiligabend_wert=6,
    )
    assert ergebnis.faktor_vj == pytest.approx(1.1)
    assert ergebnis.produktionsmenge == pytest.approx(110.0)
    assert ergebnis.differenz == pytest.approx(10.0)
    assert ergebnis.differenz_prozent == pytest.approx(10.0)
    assert ergebnis.hinweise == []


def test_berechne_tag_ohne_vorjahreswert_setzt_hinweis():
    ergebnis = f.berechne_tag(
        datum=date(2026, 12, 18),
        wochentag="Fr",
        bestellmenge=100.0,
        bestellmenge_vj=None,
        ist_vj=None,
        gewichtung=0.5,
        zuschlag=0.0,
        formel_text=f.formel_by_name(f.STANDARD_FORMEL_NAME).formel,
        gebindegroesse=0,
        tage_bis_heiligabend_wert=6,
    )
    assert ergebnis.faktor_vj is None
    assert any("Kein passender Vorjahreswert" in h for h in ergebnis.hinweise)
    assert ergebnis.produktionsmenge == pytest.approx(100.0)


def test_berechne_tag_faktor_ausserhalb_plausibler_grenzen_erzeugt_hinweis():
    ergebnis = f.berechne_tag(
        datum=date(2026, 12, 18),
        wochentag="Fr",
        bestellmenge=100.0,
        bestellmenge_vj=100.0,
        ist_vj=300.0,  # Faktor 3.0 -> deutlich außerhalb 0,5-2,0
        gewichtung=1.0,
        zuschlag=0.0,
        formel_text=f.formel_by_name(f.STANDARD_FORMEL_NAME).formel,
        gebindegroesse=0,
        tage_bis_heiligabend_wert=6,
    )
    assert any("außerhalb des üblichen Bereichs" in h for h in ergebnis.hinweise)


def test_berechne_tag_mit_gebinderundung():
    ergebnis = f.berechne_tag(
        datum=date(2026, 12, 18),
        wochentag="Fr",
        bestellmenge=103.0,
        bestellmenge_vj=100.0,
        ist_vj=100.0,
        gewichtung=1.0,
        zuschlag=0.0,
        formel_text="Bestellmenge",
        gebindegroesse=10,
        tage_bis_heiligabend_wert=6,
    )
    assert ergebnis.produktionsmenge_gerundet == pytest.approx(100.0)


def test_berechne_tag_defekte_formel_faellt_auf_bestellmenge_zurueck():
    ergebnis = f.berechne_tag(
        datum=date(2026, 12, 18),
        wochentag="Fr",
        bestellmenge=100.0,
        bestellmenge_vj=100.0,
        ist_vj=110.0,
        gewichtung=1.0,
        zuschlag=0.0,
        formel_text="Unbekannte_Variable * 2",
        gebindegroesse=0,
        tage_bis_heiligabend_wert=6,
    )
    assert ergebnis.produktionsmenge == pytest.approx(100.0)
    assert any("Formel-Fehler" in h for h in ergebnis.hinweise)


def test_berechne_woche_und_wochensumme():
    historie = f.erzeuge_beispielhistorie(2025)
    aktuelle_woche = f.erzeuge_wochen_vorlage(2026, 51, {"Mo": 300, "Di": 300, "Mi": 300, "Do": 300, "Fr": 400, "Sa": 500})
    ergebnisse = f.berechne_woche(
        aktuelle_woche,
        historie,
        f.formel_by_name(f.STANDARD_FORMEL_NAME).formel,
        gewichtung=0.5,
        zuschlag=0.05,
        gebindegroesse=10,
    )
    assert len(ergebnisse) == 6
    summe = f.berechne_wochensumme(ergebnisse)
    assert summe["summe_bestellung"] == pytest.approx(2100.0)
    assert summe["summe_produktion"] == pytest.approx(sum(e.produktionsmenge_gerundet for e in ergebnisse))


def test_berechne_woche_akzeptiert_datum_und_bestellmenge_als_text():
    # mo.ui.data_editor kann Datum/Zahl je nach Frontend als Text zurückgeben.
    historie = f.erzeuge_beispielhistorie(2025)
    aktuelle_woche = [
        {"Datum": "2026-12-14", "Wochentag": "Mo", "Bestellmenge (kg)": "300,5"},
    ]
    ergebnisse = f.berechne_woche(
        aktuelle_woche, historie, "Bestellmenge", gewichtung=0.5, zuschlag=0.0, gebindegroesse=0
    )
    assert ergebnisse[0].datum == date(2026, 12, 14)
    assert ergebnisse[0].bestellmenge == pytest.approx(300.5)


def test_vergleiche_formeln():
    historie = f.erzeuge_beispielhistorie(2025)
    aktuelle_woche = f.erzeuge_wochen_vorlage(2026, 51, {"Mo": 300, "Di": 300, "Mi": 300, "Do": 300, "Fr": 400, "Sa": 500})
    ergebnis = f.vergleiche_formeln(
        [f.STANDARD_FORMEL_NAME, "Nur Bestellmenge"],
        aktuelle_woche,
        historie,
        gewichtung=0.5,
        zuschlag=0.05,
        gebindegroesse=0,
    )
    assert set(ergebnis.keys()) == {f.STANDARD_FORMEL_NAME, "Nur Bestellmenge"}
    assert len(ergebnis[f.STANDARD_FORMEL_NAME]) == 6


# ---------------------------------------------------------------------------
# Eingabetabelle validieren
# ---------------------------------------------------------------------------


def test_validiere_eingabetabelle_erkennt_fehler():
    zeilen = [
        {"Datum": date(2026, 12, 14), "Wochentag": "Mo", "Bestellmenge (kg)": "abc"},
        {"Datum": date(2026, 12, 15), "Wochentag": "Di", "Bestellmenge (kg)": -5},
        {"Datum": date(2026, 12, 16), "Wochentag": "Mi", "Bestellmenge (kg)": "300,5"},
    ]
    bereinigt, fehler = f.validiere_eingabetabelle(zeilen)
    assert len(bereinigt) == 1
    assert bereinigt[0]["Bestellmenge (kg)"] == pytest.approx(300.5)
    assert len(fehler) == 2


# ---------------------------------------------------------------------------
# Eigene Formeln speichern/laden/löschen
# ---------------------------------------------------------------------------


def test_speichere_lade_loesche_eigene_formel(tmp_path):
    pfad = tmp_path / "formeln.json"
    formel = f.Formel(name="Testformel", formel="Bestellmenge * 1,2", beschreibung="Testbeschreibung")

    f.speichere_eigene_formel(formel, pfad)
    geladen = f.lade_eigene_formeln(pfad)
    assert len(geladen) == 1
    assert geladen[0].name == "Testformel"
    assert geladen[0].eingebaut is False

    f.loesche_eigene_formel("Testformel", pfad)
    assert f.lade_eigene_formeln(pfad) == []


def test_speichere_eigene_formel_lehnt_namen_einer_standardformel_ab(tmp_path):
    pfad = tmp_path / "formeln.json"
    formel = f.Formel(name=f.STANDARD_FORMEL_NAME, formel="Bestellmenge", beschreibung="x")
    with pytest.raises(f.FormelFehler):
        f.speichere_eigene_formel(formel, pfad)


def test_speichere_eigene_formel_lehnt_ungueltige_formel_ab(tmp_path):
    pfad = tmp_path / "formeln.json"
    formel = f.Formel(name="Kaputt", formel="Bestellmenge *", beschreibung="x")
    with pytest.raises(f.FormelFehler):
        f.speichere_eigene_formel(formel, pfad)


def test_speichere_eigene_formel_ueberschreibt_gleichnamige_eigene_formel(tmp_path):
    pfad = tmp_path / "formeln.json"
    f.speichere_eigene_formel(f.Formel("Meine Formel", "Bestellmenge", "v1"), pfad)
    f.speichere_eigene_formel(f.Formel("Meine Formel", "Bestellmenge * 2", "v2"), pfad)
    geladen = f.lade_eigene_formeln(pfad)
    assert len(geladen) == 1
    assert geladen[0].formel == "Bestellmenge * 2"


# ---------------------------------------------------------------------------
# Beispielhistorie
# ---------------------------------------------------------------------------


def test_erzeuge_beispielhistorie_umfasst_kw_49_bis_52():
    historie = f.erzeuge_beispielhistorie(2025)
    kalenderwochen = {h.datum.isocalendar()[1] for h in historie}
    assert kalenderwochen == {49, 50, 51, 52}
    assert len(historie) == 4 * 6  # 4 Wochen x 6 Arbeitstage
    assert all(h.bestellmenge > 0 and h.ist > 0 for h in historie)


# ---------------------------------------------------------------------------
# CSV-Import
# ---------------------------------------------------------------------------


def test_lese_csv_bytes_mit_semikolon():
    daten = "Datum;Wochentag;Bestellmenge (kg)\n2026-12-14;Mo;300\n".encode("utf-8")
    zeilen = f.lese_csv_bytes(daten)
    assert zeilen == [{"Datum": "2026-12-14", "Wochentag": "Mo", "Bestellmenge (kg)": "300"}]


def test_zeilen_zu_aktueller_woche():
    zeilen = [{"Datum": "2026-12-14", "Wochentag": "Mo", "Bestellmenge (kg)": "300,5"}]
    ergebnis = f.zeilen_zu_aktueller_woche(zeilen)
    assert ergebnis[0]["Datum"] == date(2026, 12, 14)
    assert ergebnis[0]["Bestellmenge (kg)"] == pytest.approx(300.5)


def test_zeilen_zu_historie():
    zeilen = [{"Datum": "2025-12-15", "Wochentag": "Mo", "Bestellmenge_VJ": "300", "Ist_VJ": "330"}]
    historie = f.zeilen_zu_historie(zeilen)
    assert len(historie) == 1
    assert historie[0].faktor_vj == pytest.approx(1.1)


# ---------------------------------------------------------------------------
# Excel-Export
# ---------------------------------------------------------------------------


def test_exportiere_ergebnis_als_excel_erzeugt_gueltige_datei(tmp_path):
    historie = f.erzeuge_beispielhistorie(2025)
    aktuelle_woche = f.erzeuge_wochen_vorlage(2026, 51, {"Mo": 300, "Di": 300, "Mi": 300, "Do": 300, "Fr": 400, "Sa": 500})
    formel = f.formel_by_name(f.STANDARD_FORMEL_NAME)
    ergebnisse = f.berechne_woche(aktuelle_woche, historie, formel.formel, 0.5, 0.05, 10)
    wochensumme = f.berechne_wochensumme(ergebnisse)

    daten = f.exportiere_ergebnis_als_excel(
        ergebnisse, wochensumme, formel.name, formel.formel, {"Gewichtung": 0.5, "Zuschlag": 0.05}
    )
    pfad = tmp_path / "export.xlsx"
    pfad.write_bytes(daten)

    import openpyxl

    arbeitsbuch = openpyxl.load_workbook(pfad)
    assert arbeitsbuch.sheetnames == ["Ergebnis", "Formel & Parameter"]
    assert arbeitsbuch["Ergebnis"]["A1"].value == "Datum"
    assert arbeitsbuch["Formel & Parameter"]["A1"].value == "Verwendete Formel"
