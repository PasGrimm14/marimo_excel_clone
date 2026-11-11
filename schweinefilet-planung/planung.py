# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "marimo",
#     "polars==2.1.0",
#     "altair==6.3.0",
#     "openpyxl==3.1.5",
#     "simpleeval==1.0.8",
# ]
# ///

import marimo

__generated_with = "0.25.2"
app = marimo.App(width="medium", app_title="Schweinefilet-Produktionsplanung")


@app.cell
def _():
    import marimo as mo

    return (mo,)


@app.cell
def _():
    import altair as alt
    import polars as pl

    import formeln as f

    return alt, f, pl


@app.cell
def _(mo):
    mo.md("""
    # Produktionsplanung Schweinefilet - Weihnachtswoche

    Dieses Notebook berechnet die geplante Produktionsmenge für eine
    Woche (Montag bis Samstag) auf Basis der Bestellmengen und der
    Erfahrungswerte aus dem Vorjahr. Es funktioniert wie ein
    Excel-Tabellenblatt: Zahlen eintragen, Formel wählen, auf
    **Berechnen** klicken.

    Die Reiter unten führen durch die Bedienung:

    - **Berechnung**: Bestellmengen eintragen, Formel wählen, Ergebnis ansehen und exportieren.
    - **Formeln verwalten**: eigene Formeln erstellen, speichern und löschen.
    - **Vergleich**: mehrere Formeln nebeneinander vergleichen.
    """)
    return


@app.cell
def _(mo):
    # Zähler, die per Knopfdruck hochgezählt werden, um UI-Elemente gezielt
    # neu zu erzeugen (z. B. die Eingabetabelle zurückzusetzen oder die
    # Formelliste nach dem Speichern/Löschen zu aktualisieren).
    get_tabellen_reset, set_tabellen_reset = mo.state(0)
    get_formel_version, set_formel_version = mo.state(0)
    get_formel_vorlage_text, set_formel_vorlage_text = mo.state("")
    return (
        get_formel_version,
        get_formel_vorlage_text,
        get_tabellen_reset,
        set_formel_version,
        set_formel_vorlage_text,
        set_tabellen_reset,
    )


@app.cell
def _(mo):
    ueberschrift_historie = mo.md("## 1. Historiendaten des Vorjahres")
    return (ueberschrift_historie,)


@app.cell
def _(mo):
    historie_upload = mo.ui.file(
        filetypes=[".csv", ".xlsx"],
        kind="area",
        label="Historiendaten hochladen (CSV oder Excel, Spalten: Datum, Wochentag, Bestellmenge_VJ, Ist_VJ). Ohne Upload werden Beispieldaten verwendet.",
    )
    return (historie_upload,)


@app.cell
def _(f, historie_upload):
    if historie_upload.value:
        _datei = historie_upload.value[0]
        if _datei.name.lower().endswith(".csv"):
            _zeilen = f.lese_csv_bytes(_datei.contents)
        else:
            _zeilen = f.lese_excel_bytes(_datei.contents)
        historie = f.zeilen_zu_historie(_zeilen)
        historie_quelle = f"Hochgeladene Datei **{_datei.name}** mit {len(historie)} Tagen."
    else:
        historie = f.erzeuge_beispielhistorie(2025)
        historie_quelle = (
            "Keine Datei hochgeladen - es werden automatisch erzeugte "
            "**Beispieldaten für KW 49-52 des Vorjahres** verwendet."
        )
    return historie, historie_quelle


@app.cell
def _(historie, historie_quelle, mo):
    historie_hinweis = mo.callout(
        mo.md(f"{historie_quelle} Verfügbare Vorjahrestage: **{len(historie)}**."),
        kind="info",
    )
    return (historie_hinweis,)


@app.cell
def _(mo):
    ueberschrift_bestellung = mo.md("## 2. Bestellmengen der aktuellen Woche")
    return (ueberschrift_bestellung,)


@app.cell
def _(mo):
    kalenderwoche_dropdown = mo.ui.dropdown(
        options=["49", "50", "51", "52"],
        value="51",
        label="Kalenderwoche der Planungswoche",
    )
    jahr_eingabe = mo.ui.number(start=2020, stop=2100, step=1, value=2026, label="Jahr")
    return jahr_eingabe, kalenderwoche_dropdown


@app.cell
def _(mo):
    bestellung_upload = mo.ui.file(
        filetypes=[".csv", ".xlsx"],
        kind="button",
        label="Bestellmengen aus Datei übernehmen (optional, Spalten: Datum, Wochentag, Bestellmenge (kg))",
    )
    return (bestellung_upload,)


@app.cell
def _(bestellung_upload, f, jahr_eingabe, kalenderwoche_dropdown):
    if bestellung_upload.value:
        _datei = bestellung_upload.value[0]
        if _datei.name.lower().endswith(".csv"):
            _zeilen = f.lese_csv_bytes(_datei.contents)
        else:
            _zeilen = f.lese_excel_bytes(_datei.contents)
        wochen_vorlage = f.zeilen_zu_aktueller_woche(_zeilen)
    else:
        wochen_vorlage = f.erzeuge_wochen_vorlage(
            int(jahr_eingabe.value), int(kalenderwoche_dropdown.value)
        )
    return (wochen_vorlage,)


@app.cell
def _(get_tabellen_reset, mo, wochen_vorlage):
    # get_tabellen_reset() wird referenziert, damit diese Zelle beim Klick auf
    # "Zurücksetzen" erneut läuft und die Tabelle neu (mit der Vorbelegung) erzeugt.
    _ = get_tabellen_reset()
    tabelle = mo.ui.data_editor(
        data=wochen_vorlage,
        label="Bestellmengen (kg) - Datum und Wochentag sind vorbefüllt, bitte nur die Mengen eintragen",
        editable_columns=["Bestellmenge (kg)"],
    )
    return (tabelle,)


@app.cell
def _(mo):
    berechnen_knopf = mo.ui.run_button(label="Berechnen", kind="success")
    zuruecksetzen_knopf = mo.ui.run_button(label="Zurücksetzen")
    return berechnen_knopf, zuruecksetzen_knopf


@app.cell
def _(get_tabellen_reset, set_tabellen_reset, zuruecksetzen_knopf):
    if zuruecksetzen_knopf.value:
        set_tabellen_reset(get_tabellen_reset() + 1)
    return


@app.cell
def _(mo):
    ueberschrift_parameter = mo.md("## 3. Parameter")
    return (ueberschrift_parameter,)


@app.cell
def _(mo):
    gewichtung_slider = mo.ui.slider(
        start=0,
        stop=100,
        step=5,
        value=50,
        show_value=True,
        label="Gewichtung der Historie (0 % = nur Bestellmenge, 100 % = volle Vorjahres-Abweichung)",
    )
    gewichtung_slider
    zuschlag_eingabe = mo.ui.number(
        start=-20, stop=50, step=1, value=5, label="Sicherheitszuschlag (%)"
    )
    zuschlag_eingabe
    rundung_dropdown = mo.ui.dropdown(
        options={
            "Keine Rundung": 0,
            "Auf volle 1 kg": 1,
            "Auf volle 5 kg": 5,
            "Auf volle 10 kg": 10,
            "Auf Kisten à 25 kg": 25,
        },
        value="Keine Rundung",
        label="Rundung auf Gebinde-/Chargengröße",
    )
    return gewichtung_slider, rundung_dropdown, zuschlag_eingabe


@app.cell
def _(mo):
    ueberschrift_formel = mo.md("## 4. Formel auswählen")
    return (ueberschrift_formel,)


@app.cell
def _(f, get_formel_version):
    # get_formel_version() wird referenziert, damit diese Zelle nach dem
    # Speichern/Löschen einer eigenen Formel erneut läuft.
    _ = get_formel_version()
    eigene_formeln = f.lade_eigene_formeln()
    return (eigene_formeln,)


@app.cell
def _(eigene_formeln, f, mo):
    formel_quelle_radio = mo.ui.radio(
        options=["Aus Formelbibliothek wählen", "Eigene Formel direkt eingeben (einmalig)"],
        value="Aus Formelbibliothek wählen",
        label="Formelquelle",
    )
    formel_dropdown = mo.ui.dropdown(
        options={formel.name: formel.name for formel in f.STANDARDFORMELN + eigene_formeln},
        value=f.STANDARD_FORMEL_NAME,
        label="Formel auswählen",
    )
    ad_hoc_formel_text = mo.ui.text_area(
        value="Bestellmenge * Faktor_VJ",
        label="Eigene Formel (wird nur für diese Berechnung verwendet, nicht gespeichert)",
        rows=3,
        full_width=True,
    )
    return ad_hoc_formel_text, formel_dropdown, formel_quelle_radio


@app.cell
def _(
    ad_hoc_formel_text,
    eigene_formeln,
    f,
    formel_dropdown,
    formel_quelle_radio,
    mo,
):
    if formel_quelle_radio.value == "Aus Formelbibliothek wählen":
        _formel = f.formel_by_name(formel_dropdown.value, eigene_formeln)
        aktive_formel_name = _formel.name
        aktive_formel_text = _formel.formel
        formel_hinweis = mo.callout(
            mo.md(f"**{_formel.name}**\n\n`{_formel.formel}`\n\n{_formel.beschreibung}"),
            kind="info",
        )
    else:
        aktive_formel_name = "Eigene Formel (einmalig)"
        aktive_formel_text = ad_hoc_formel_text.value
        try:
            _normalisiert = f.validiere_formel(aktive_formel_text)
            formel_hinweis = mo.callout(
                mo.md(f"Formel ist gültig und wird so berechnet: `{_normalisiert}`"), kind="success"
            )
        except f.FormelFehler as _exc:
            formel_hinweis = mo.callout(mo.md(f"**Formel-Fehler:** {_exc}"), kind="danger")
    return aktive_formel_name, aktive_formel_text, formel_hinweis


@app.cell
def _(
    ad_hoc_formel_text,
    formel_dropdown,
    formel_hinweis,
    formel_quelle_radio,
    mo,
):
    if formel_quelle_radio.value == "Aus Formelbibliothek wählen":
        bereich_formelauswahl = mo.vstack([formel_quelle_radio, formel_dropdown, formel_hinweis])
    else:
        bereich_formelauswahl = mo.vstack([formel_quelle_radio, ad_hoc_formel_text, formel_hinweis])
    return (bereich_formelauswahl,)


@app.cell
def _(
    berechnen_knopf,
    bereich_formelauswahl,
    bestellung_upload,
    gewichtung_slider,
    historie_hinweis,
    historie_upload,
    jahr_eingabe,
    kalenderwoche_dropdown,
    mo,
    rundung_dropdown,
    tabelle,
    ueberschrift_bestellung,
    ueberschrift_formel,
    ueberschrift_historie,
    ueberschrift_parameter,
    zuruecksetzen_knopf,
    zuschlag_eingabe,
):
    bereich_eingabe = mo.vstack(
        [
            ueberschrift_historie,
            historie_upload,
            historie_hinweis,
            mo.md("---"),
            ueberschrift_bestellung,
            mo.hstack([jahr_eingabe, kalenderwoche_dropdown], justify="start"),
            bestellung_upload,
            tabelle,
            mo.hstack([berechnen_knopf, zuruecksetzen_knopf], justify="start"),
            mo.md("---"),
            ueberschrift_parameter,
            gewichtung_slider,
            zuschlag_eingabe,
            rundung_dropdown,
            mo.md("---"),
            ueberschrift_formel,
            bereich_formelauswahl,
        ]
    )
    return (bereich_eingabe,)


@app.cell
def _(
    aktive_formel_name,
    aktive_formel_text,
    alt,
    berechnen_knopf,
    f,
    gewichtung_slider,
    historie,
    mo,
    pl,
    rundung_dropdown,
    tabelle,
    zuschlag_eingabe,
):
    if not berechnen_knopf.value:
        berechnung_ausgabe = mo.callout(
            mo.md("Bitte Bestellmengen eintragen und auf **Berechnen** klicken."), kind="neutral"
        )
    else:
        _zeilen, _fehler = f.validiere_eingabetabelle(tabelle.value)

        if not _zeilen:
            berechnung_ausgabe = mo.callout(
                mo.md("Keine gültigen Zeilen in der Eingabetabelle. Bitte Bestellmengen prüfen:\n\n"
                      + "\n".join(f"- {m}" for m in _fehler)),
                kind="danger",
            )
        else:
            _gewichtung = gewichtung_slider.value / 100
            _zuschlag = zuschlag_eingabe.value / 100
            _gebinde = rundung_dropdown.value

            ergebnisse = f.berechne_woche(_zeilen, historie, aktive_formel_text, _gewichtung, _zuschlag, _gebinde)
            wochensumme = f.berechne_wochensumme(ergebnisse)

            _tabellen_daten = pl.DataFrame([e.as_dict() for e in ergebnisse])
            _ergebnistabelle = mo.ui.table(_tabellen_daten, show_download=False, label="Ergebnis pro Tag")

            _kennzahlen = mo.hstack(
                [
                    mo.stat(f"{wochensumme['summe_bestellung']:.0f} kg", label="Bestellmenge Woche"),
                    mo.stat(f"{wochensumme['summe_produktion']:.0f} kg", label="Geplante Produktion Woche", bordered=True),
                    mo.stat(
                        f"{wochensumme['differenz']:+.0f} kg",
                        label="Differenz zur Bestellmenge",
                        caption=f"{wochensumme['differenz_prozent']:+.1f} %" if wochensumme["differenz_prozent"] is not None else "",
                    ),
                    mo.stat(
                        f"{wochensumme['summe_ist_vj']:.0f} kg" if wochensumme["summe_ist_vj"] is not None else "n/v",
                        label="Vorjahres-Ist Woche",
                    ),
                ],
                justify="start",
            )

            _alle_hinweise = [f"{e.wochentag} {e.datum}: {h}" for e in ergebnisse for h in e.hinweise]
            if _fehler:
                _alle_hinweise = [f"Eingabe: {m}" for m in _fehler] + _alle_hinweise
            if _alle_hinweise:
                _hinweise_anzeige = mo.callout(
                    mo.md("**Hinweise zur Plausibilität:**\n\n" + "\n".join(f"- {h}" for h in _alle_hinweise)),
                    kind="warn",
                )
            else:
                _hinweise_anzeige = mo.md("")

            _diagramm_daten = pl.concat(
                [
                    _tabellen_daten.select(
                        pl.col("datum"), pl.col("bestellmenge").alias("Menge"), pl.lit("Bestellmenge").alias("Kategorie")
                    ),
                    _tabellen_daten.select(
                        pl.col("datum"),
                        pl.col("produktionsmenge_gerundet").alias("Menge"),
                        pl.lit("Produktionsmenge (geplant)").alias("Kategorie"),
                    ),
                    _tabellen_daten.filter(pl.col("ist_vj").is_not_null()).select(
                        pl.col("datum"), pl.col("ist_vj").alias("Menge"), pl.lit("Vorjahr (Ist)").alias("Kategorie")
                    ),
                ]
            )
            _diagramm = (
                alt.Chart(_diagramm_daten)
                .mark_line(point=True)
                .encode(
                    x=alt.X("datum:T", title="Tag"),
                    y=alt.Y("Menge:Q", title="Menge (kg)"),
                    color=alt.Color("Kategorie:N", title=""),
                    tooltip=["datum", "Kategorie", "Menge"],
                )
                .properties(title="Bestellmenge vs. geplante Produktion vs. Vorjahr", width=600, height=320)
            )

            _export_bytes = f.exportiere_ergebnis_als_excel(
                ergebnisse,
                wochensumme,
                aktive_formel_name,
                aktive_formel_text,
                {
                    "Gewichtung Historie (%)": gewichtung_slider.value,
                    "Sicherheitszuschlag (%)": zuschlag_eingabe.value,
                    "Rundung auf (kg, 0 = keine)": rundung_dropdown.value,
                },
            )
            _download = mo.download(
                data=_export_bytes,
                filename="produktionsplanung_woche.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                label="Ergebnis als Excel-Datei herunterladen",
            )

            berechnung_ausgabe = mo.vstack(
                [
                    mo.md("### Ergebnis"),
                    _kennzahlen,
                    _hinweise_anzeige,
                    _ergebnistabelle,
                    _diagramm,
                    _download,
                ]
            )
    return (berechnung_ausgabe,)


@app.cell
def _(berechnung_ausgabe, bereich_eingabe, mo):
    tab_berechnung = mo.vstack([bereich_eingabe, mo.md("---"), berechnung_ausgabe])
    return (tab_berechnung,)


@app.cell
def _(f, mo):
    _zeilen = "\n".join(
        f"| `{name}` | {bedeutung} | {beispiel} |" for name, (bedeutung, beispiel) in f.VARIABLEN_INFO.items()
    )
    bereich_variablenuebersicht = mo.md(
        "### Verfügbare Variablen für eigene Formeln\n\n"
        "| Variable | Bedeutung | Beispielwert |\n"
        "|---|---|---|\n"
        f"{_zeilen}\n\n"
        "**Erlaubte Funktionen:** `MIN`, `MAX`, `RUNDEN`, `AUFRUNDEN`, `ABRUNDEN`, `WENN`, `ABS`, `MITTELWERT`. "
        "**Erlaubte Operatoren:** `+ - * / ( )` sowie Vergleiche (`=`, `<>`, `<`, `<=`, `>`, `>=`). "
        "Dezimalkomma und Semikolon (wie im deutschen Excel) werden ebenso akzeptiert wie Dezimalpunkt und Komma."
    )
    return (bereich_variablenuebersicht,)


@app.cell
def _(eigene_formeln, f, mo):
    _vorgefertigt = "\n\n".join(
        f"**{formel.name}**  \n`{formel.formel}`  \n{formel.beschreibung}" for formel in f.STANDARDFORMELN
    )
    if eigene_formeln:
        _eigene = "\n\n".join(
            f"**{formel.name}**  \n`{formel.formel}`  \n{formel.beschreibung}" for formel in eigene_formeln
        )
    else:
        _eigene = "_Noch keine eigenen Formeln gespeichert._"
    bereich_formelreferenz = mo.accordion(
        {
            "Vorgefertigte Formeln (Formelbibliothek)": mo.md(_vorgefertigt),
            "Eigene gespeicherte Formeln": mo.md(_eigene),
        }
    )
    return (bereich_formelreferenz,)


@app.cell
def _(eigene_formeln, f, mo):
    vorlage_dropdown = mo.ui.dropdown(
        options={formel.name: formel.formel for formel in f.STANDARDFORMELN + eigene_formeln},
        label="Formel als Vorlage wählen (Original bleibt beim Speichern unter neuem Namen unverändert)",
    )
    vorlage_laden_knopf = mo.ui.run_button(label="Als Vorlage in das Textfeld laden")
    return vorlage_dropdown, vorlage_laden_knopf


@app.cell
def _(set_formel_vorlage_text, vorlage_dropdown, vorlage_laden_knopf):
    if vorlage_laden_knopf.value and vorlage_dropdown.value:
        set_formel_vorlage_text(vorlage_dropdown.value)
    return


@app.cell
def _(get_formel_vorlage_text, mo):
    eigene_formel_text = mo.ui.text_area(
        value=get_formel_vorlage_text(),
        placeholder="z. B. Bestellmenge * Faktor_VJ",
        label="Formel bearbeiten (Excel-Schreibweise)",
        rows=4,
        full_width=True,
    )
    return (eigene_formel_text,)


@app.cell
def _(eigene_formel_text, f, mo):
    _beispielwerte = f.beispielwerte()
    try:
        _ergebnis = f.werte_formel_aus(eigene_formel_text.value, _beispielwerte)
        _beschreibung_werte = ", ".join(f"{k} = {v}" for k, v in _beispielwerte.items())
        vorschau_anzeige = mo.callout(
            mo.md(
                f"**Live-Vorschau für einen Beispieltag** ({_beschreibung_werte}):\n\n"
                f"**Ergebnis: {_ergebnis:.2f} kg**"
            ),
            kind="success",
        )
    except f.FormelFehler as _exc:
        vorschau_anzeige = mo.callout(mo.md(f"**Formel-Fehler:** {_exc}"), kind="danger")
    return (vorschau_anzeige,)


@app.cell
def _(mo):
    neuer_formel_name = mo.ui.text(label="Name der neuen Formel", placeholder="z. B. Montag ohne Zuschlag")
    neue_formel_beschreibung = mo.ui.text_area(
        label="Beschreibung (wann ist diese Formel sinnvoll?)", rows=2, full_width=True
    )
    speichern_knopf = mo.ui.run_button(label="Formel speichern", kind="success")
    return neue_formel_beschreibung, neuer_formel_name, speichern_knopf


@app.cell
def _(
    eigene_formel_text,
    f,
    get_formel_version,
    mo,
    neue_formel_beschreibung,
    neuer_formel_name,
    set_formel_version,
    speichern_knopf,
):
    if not speichern_knopf.value:
        speicher_meldung = mo.md("")
    elif not neuer_formel_name.value.strip():
        speicher_meldung = mo.callout(mo.md("Bitte einen Namen für die Formel eingeben."), kind="danger")
    else:
        try:
            f.speichere_eigene_formel(
                f.Formel(
                    name=neuer_formel_name.value.strip(),
                    formel=eigene_formel_text.value,
                    beschreibung=neue_formel_beschreibung.value.strip(),
                )
            )
            set_formel_version(get_formel_version() + 1)
            speicher_meldung = mo.callout(
                mo.md(
                    f"Formel **'{neuer_formel_name.value.strip()}'** wurde gespeichert und erscheint jetzt "
                    "in der Formelauswahl im Reiter 'Berechnung'."
                ),
                kind="success",
            )
        except f.FormelFehler as _exc:
            speicher_meldung = mo.callout(mo.md(f"Formel konnte nicht gespeichert werden: {_exc}"), kind="danger")
    return (speicher_meldung,)


@app.cell
def _(eigene_formeln, mo):
    loeschen_dropdown = mo.ui.dropdown(
        options=[formel.name for formel in eigene_formeln],
        label="Eigene Formel löschen (vorgefertigte Formeln können nicht gelöscht werden)",
    )
    loeschen_knopf = mo.ui.run_button(label="Formel löschen", kind="danger")
    return loeschen_dropdown, loeschen_knopf


@app.cell
def _(
    f,
    get_formel_version,
    loeschen_dropdown,
    loeschen_knopf,
    mo,
    set_formel_version,
):
    if not loeschen_knopf.value:
        loeschen_meldung = mo.md("")
    elif not loeschen_dropdown.value:
        loeschen_meldung = mo.callout(mo.md("Bitte eine eigene Formel zum Löschen auswählen."), kind="danger")
    else:
        f.loesche_eigene_formel(loeschen_dropdown.value)
        set_formel_version(get_formel_version() + 1)
        loeschen_meldung = mo.callout(
            mo.md(f"Formel **'{loeschen_dropdown.value}'** wurde gelöscht."), kind="success"
        )
    return (loeschen_meldung,)


@app.cell
def _(
    bereich_formelreferenz,
    bereich_variablenuebersicht,
    eigene_formel_text,
    loeschen_dropdown,
    loeschen_knopf,
    loeschen_meldung,
    mo,
    neue_formel_beschreibung,
    neuer_formel_name,
    speicher_meldung,
    speichern_knopf,
    vorlage_dropdown,
    vorlage_laden_knopf,
    vorschau_anzeige,
):
    tab_formeln_verwalten = mo.vstack(
        [
            mo.md("## Formeln verwalten"),
            bereich_variablenuebersicht,
            mo.md("---"),
            bereich_formelreferenz,
            mo.md("---"),
            mo.md("### Eigene Formel erstellen"),
            mo.hstack([vorlage_dropdown, vorlage_laden_knopf], justify="start"),
            eigene_formel_text,
            vorschau_anzeige,
            mo.md("---"),
            neuer_formel_name,
            neue_formel_beschreibung,
            speichern_knopf,
            speicher_meldung,
            mo.md("---"),
            mo.md("### Eigene Formel löschen"),
            mo.hstack([loeschen_dropdown, loeschen_knopf], justify="start"),
            loeschen_meldung,
        ]
    )
    return (tab_formeln_verwalten,)


@app.cell
def _(eigene_formeln, f, mo):
    multiselect_formeln = mo.ui.multiselect(
        options=f.alle_formelnamen(eigene_formeln),
        value=[f.STANDARD_FORMEL_NAME, "Nur Bestellmenge"],
        label="Formeln für den Vergleich auswählen",
    )
    vergleich_knopf = mo.ui.run_button(label="Vergleich berechnen", kind="success")
    return multiselect_formeln, vergleich_knopf


@app.cell
def _(
    alt,
    eigene_formeln,
    f,
    gewichtung_slider,
    historie,
    mo,
    multiselect_formeln,
    pl,
    rundung_dropdown,
    tabelle,
    vergleich_knopf,
    zuschlag_eingabe,
):
    if not vergleich_knopf.value:
        vergleich_ausgabe = mo.callout(
            mo.md("Formeln auswählen und auf **Vergleich berechnen** klicken."), kind="neutral"
        )
    else:
        _zeilen, _fehler = f.validiere_eingabetabelle(tabelle.value)

        if not _zeilen or not multiselect_formeln.value:
            vergleich_ausgabe = mo.callout(
                mo.md(
                    "Keine gültigen Bestellmengen oder keine Formel ausgewählt. Bitte im Reiter "
                    "'Berechnung' die Bestellmengen prüfen und hier mindestens eine Formel auswählen.\n\n"
                    + "\n".join(f"- {m}" for m in _fehler)
                ),
                kind="danger",
            )
        else:
            _gewichtung = gewichtung_slider.value / 100
            _zuschlag = zuschlag_eingabe.value / 100
            _gebinde = rundung_dropdown.value

            _ergebnisse_je_formel = f.vergleiche_formeln(
                multiselect_formeln.value, _zeilen, historie, _gewichtung, _zuschlag, _gebinde, eigene_formeln
            )

            _basis = pl.DataFrame(
                [
                    {
                        "Datum": e.datum,
                        "Wochentag": e.wochentag,
                        "Bestellmenge": e.bestellmenge,
                        "Vorjahr (Ist)": e.ist_vj,
                    }
                    for e in next(iter(_ergebnisse_je_formel.values()))
                ]
            )
            for _name, _ergebnisse in _ergebnisse_je_formel.items():
                _basis = _basis.with_columns(
                    pl.Series(_name, [e.produktionsmenge_gerundet for e in _ergebnisse])
                )

            _vergleichstabelle = mo.ui.table(
                _basis, show_download=False, label="Produktionsmenge (kg) je Formel und Tag"
            )

            _zusammenfassung = []
            for _name, _ergebnisse in _ergebnisse_je_formel.items():
                _summe = f.berechne_wochensumme(_ergebnisse)
                if _summe["summe_ist_vj"] is not None:
                    _abweichung_vj = _summe["summe_produktion"] - _summe["summe_ist_vj"]
                    _abweichung_vj_prozent = (
                        _abweichung_vj / _summe["summe_ist_vj"] * 100 if _summe["summe_ist_vj"] else None
                    )
                else:
                    _abweichung_vj = None
                    _abweichung_vj_prozent = None
                _zusammenfassung.append(
                    {
                        "Formel": _name,
                        "Wochensumme Produktion (kg)": round(_summe["summe_produktion"], 1),
                        "Differenz zur Bestellmenge (kg)": round(_summe["differenz"], 1),
                        "Differenz zur Bestellmenge (%)": (
                            round(_summe["differenz_prozent"], 1) if _summe["differenz_prozent"] is not None else None
                        ),
                        "Abweichung zum Vorjahres-Ist (kg)": (
                            round(_abweichung_vj, 1) if _abweichung_vj is not None else None
                        ),
                        "Abweichung zum Vorjahres-Ist (%)": (
                            round(_abweichung_vj_prozent, 1) if _abweichung_vj_prozent is not None else None
                        ),
                    }
                )
            _zusammenfassungstabelle = mo.ui.table(
                pl.DataFrame(_zusammenfassung), show_download=False, label="Wochensumme je Formel"
            )

            _diagramm_teile = [
                _basis.select(
                    pl.col("Datum"), pl.col("Bestellmenge").alias("Menge"), pl.lit("Bestellmenge").alias("Kategorie")
                )
            ]
            if _basis["Vorjahr (Ist)"].null_count() < len(_basis):
                _diagramm_teile.append(
                    _basis.filter(pl.col("Vorjahr (Ist)").is_not_null()).select(
                        pl.col("Datum"),
                        pl.col("Vorjahr (Ist)").alias("Menge"),
                        pl.lit("Vorjahr (Ist)").alias("Kategorie"),
                    )
                )
            for _name in _ergebnisse_je_formel:
                _diagramm_teile.append(
                    _basis.select(pl.col("Datum"), pl.col(_name).alias("Menge"), pl.lit(_name).alias("Kategorie"))
                )
            _diagramm_daten = pl.concat(_diagramm_teile)
            _diagramm = (
                alt.Chart(_diagramm_daten)
                .mark_line(point=True)
                .encode(
                    x=alt.X("Datum:T", title="Tag"),
                    y=alt.Y("Menge:Q", title="Menge (kg)"),
                    color=alt.Color("Kategorie:N", title=""),
                    tooltip=["Datum", "Kategorie", "Menge"],
                )
                .properties(title="Formelvergleich", width=600, height=320)
            )

            if _fehler:
                _fehler_anzeige = mo.callout(
                    mo.md("**Hinweis: Folgende Zeilen der Eingabetabelle wurden ignoriert:**\n\n"
                          + "\n".join(f"- {m}" for m in _fehler)),
                    kind="warn",
                )
            else:
                _fehler_anzeige = mo.md("")

            vergleich_ausgabe = mo.vstack(
                [
                    mo.md("### Vergleichsergebnis"),
                    _fehler_anzeige,
                    _vergleichstabelle,
                    _zusammenfassungstabelle,
                    _diagramm,
                ]
            )
    return (vergleich_ausgabe,)


@app.cell
def _(mo, multiselect_formeln, vergleich_ausgabe, vergleich_knopf):
    tab_vergleich = mo.vstack(
        [
            mo.md("## Formelvergleich"),
            multiselect_formeln,
            vergleich_knopf,
            mo.md("---"),
            vergleich_ausgabe,
        ]
    )
    return (tab_vergleich,)


@app.cell
def _(mo, tab_berechnung, tab_formeln_verwalten, tab_vergleich):
    mo.ui.tabs(
        {
            "Berechnung": tab_berechnung,
            "Formeln verwalten": tab_formeln_verwalten,
            "Vergleich": tab_vergleich,
        }
    )
    return


if __name__ == "__main__":
    app.run()
