"""Formelbibliothek, sicherer Formel-Auswerter und Berechnungslogik.

Dieses Modul enthält die gesamte fachliche Logik für die
Produktionsplanung Schweinefilet. Es ist absichtlich frei von
Marimo- und UI-Code, damit alle Funktionen einfach mit pytest
getestet werden können. Das Notebook ``planung.py`` importiert
dieses Modul und kümmert sich ausschließlich um Darstellung und
Bedienung.

Sicherheitshinweis
-------------------
Eigene Formeln werden **niemals** mit ``eval()`` oder ``exec()``
ausgewertet. Stattdessen kommt ``simpleeval`` mit einer festen
Positivliste erlaubter Variablen, Funktionen und Operatoren zum
Einsatz (siehe :func:`werte_formel_aus`). Zusätzlich wird die
Formel vorab über eine strenge, auf eine Whitelist von
AST-Knotentypen beschränkte Prüfung validiert (siehe
:func:`validiere_formel`), sodass z. B. ``__import__(...)`` oder
Attributzugriffe wie ``(1).__class__`` schon vor der Auswertung
abgelehnt werden.
"""

from __future__ import annotations

import ast
import csv
import difflib
import io
import json
import math
import random
import re
from dataclasses import asdict, dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

import simpleeval

# ---------------------------------------------------------------------------
# Verfügbare Variablen für eigene Formeln
# ---------------------------------------------------------------------------
# Name -> (Bedeutung, Beispielwert). Die Beispielwerte werden u. a. für die
# Live-Vorschau beim Erstellen eigener Formeln verwendet.

VARIABLEN_INFO: dict[str, tuple[str, Any]] = {
    "Bestellmenge": ("Bestellmenge der aktuellen Woche an diesem Tag (kg)", 500.0),
    "Bestellmenge_VJ": ("Bestellmenge des Vorjahres am vergleichbaren Tag (kg)", 480.0),
    "Ist_VJ": ("Tatsächlich produzierte Menge des Vorjahres am vergleichbaren Tag (kg)", 510.0),
    "Faktor_VJ": ("Abweichungsfaktor Vorjahr: Ist_VJ geteilt durch Bestellmenge_VJ", 1.06),
    "Gewichtung": ("Gewichtung der Historie: 0 = nur Bestellmenge, 1 = volle Historie", 0.5),
    "Zuschlag": ("Sicherheitszuschlag als Dezimalwert, 5 % entspricht 0,05", 0.05),
    "Wochentag": ("Wochentag als Kürzel: 'Mo', 'Di', 'Mi', 'Do', 'Fr' oder 'Sa'", "Fr"),
    "Tage_bis_Heiligabend": ("Anzahl Tage von diesem Tag bis Heiligabend (24.12.)", 3),
}


def beispielwerte() -> dict[str, Any]:
    """Beispielhafte Werte für alle Formel-Variablen, für die Live-Vorschau."""

    return {name: info[1] for name, info in VARIABLEN_INFO.items()}


# ---------------------------------------------------------------------------
# Erlaubte Funktionen (deutsche Excel-Namen)
# ---------------------------------------------------------------------------


def _runden_basis(zahl: float, stellen: float, modus: str) -> float:
    stellen_i = int(round(stellen))
    faktor = 10**stellen_i
    wert = zahl * faktor
    if modus == "kaufmaennisch":
        ergebnis = math.floor(wert + 0.5) if wert >= 0 else math.ceil(wert - 0.5)
    elif modus == "auf":
        ergebnis = math.ceil(wert) if wert >= 0 else math.floor(wert)
    elif modus == "ab":
        ergebnis = math.floor(wert) if wert >= 0 else math.ceil(wert)
    else:  # pragma: no cover - interner Programmierfehler
        raise ValueError(f"Unbekannter Rundungsmodus: {modus}")
    return ergebnis / faktor


def runden(zahl: float, stellen: float = 0) -> float:
    """Entspricht Excel RUNDEN: kaufmännisches Runden auf ``stellen`` Nachkommastellen."""

    return _runden_basis(zahl, stellen, "kaufmaennisch")


def aufrunden(zahl: float, stellen: float = 0) -> float:
    """Entspricht Excel AUFRUNDEN: rundet immer vom Nullpunkt weg."""

    return _runden_basis(zahl, stellen, "auf")


def abrunden(zahl: float, stellen: float = 0) -> float:
    """Entspricht Excel ABRUNDEN: rundet immer zum Nullpunkt hin (abschneiden)."""

    return _runden_basis(zahl, stellen, "ab")


def wenn(bedingung: Any, dann: Any, sonst: Any) -> Any:
    """Entspricht Excel WENN(Bedingung; Dann; Sonst).

    Hinweis: Da WENN hier als Funktion (nicht als Sprachkonstrukt)
    umgesetzt ist, werden beide Zweige vorab berechnet (keine
    Kurzschlussauswertung). Für die in diesem Projekt verwendeten
    arithmetischen Formeln ist das unkritisch.
    """

    return dann if bedingung else sonst


def mittelwert(*werte: float) -> float:
    """Entspricht Excel MITTELWERT: arithmetisches Mittel der Argumente."""

    if not werte:
        raise ValueError("MITTELWERT benötigt mindestens einen Wert.")
    return sum(werte) / len(werte)


# Positivliste der Funktionen, die in eigenen Formeln verwendet werden dürfen.
ERLAUBTE_FUNKTIONEN: dict[str, Any] = {
    "MIN": min,
    "MAX": max,
    "ABS": abs,
    "RUNDEN": runden,
    "AUFRUNDEN": aufrunden,
    "ABRUNDEN": abrunden,
    "WENN": wenn,
    "MITTELWERT": mittelwert,
}


# ---------------------------------------------------------------------------
# Normalisierung deutscher Excel-Schreibweise -> auswertbarer Ausdruck
# ---------------------------------------------------------------------------

# Ein Komma zwischen zwei Ziffern ist ein Dezimalkomma (z. B. "0,9" -> "0.9").
# Ein Komma, das NICHT zwischen zwei Ziffern steht, ist ein Argumenttrenner
# und bleibt (für Python/simpleeval) unverändert ein Komma.
_DEZIMALKOMMA = re.compile(r"(?<=\d),(?=\d)")

# "==", "!=", "<=", ">=" bleiben unverändert; ein einzelnes "=" (Excel-Gleichheit)
# wird zu "==" (Python-Gleichheit).
_VERGLEICHSOPERATOREN = re.compile(r"(==|!=|<=|>=)|(=)")


def _ersetze_vergleich(treffer: re.Match[str]) -> str:
    return treffer.group(1) if treffer.group(1) else "=="


def normalisiere_formel(text: str) -> str:
    """Wandelt deutsche Excel-Schreibweise in einen auswertbaren Ausdruck um.

    - Semikolon (Argumenttrenner in deutschem Excel) -> Komma
    - Dezimalkomma zwischen Ziffern -> Dezimalpunkt
    - "<>" (Excel "ungleich") -> "!="
    - einzelnes "=" (Excel "gleich") -> "==" (Python "gleich")
    """

    ergebnis = text.replace(";", ",")
    ergebnis = _DEZIMALKOMMA.sub(".", ergebnis)
    ergebnis = ergebnis.replace("<>", "NE")  # vor "=" -> "==" schützen
    ergebnis = _VERGLEICHSOPERATOREN.sub(_ersetze_vergleich, ergebnis)
    ergebnis = ergebnis.replace("NE", "!=")
    return ergebnis


# ---------------------------------------------------------------------------
# Validierung eigener Formeln
# ---------------------------------------------------------------------------


class FormelFehler(Exception):
    """Für Fachanwender verständliche Fehlermeldung zu einer Formel."""


# Strenge Whitelist erlaubter AST-Knotentypen. Alles, was hier nicht
# aufgeführt ist (Attribute, Subscript, Lambda, Comprehensions, Import, ...),
# führt zu einer Ablehnung der Formel, BEVOR simpleeval überhaupt aufgerufen wird.
_ERLAUBTE_KNOTEN = (
    ast.Expression,
    ast.BinOp,
    ast.UnaryOp,
    ast.Compare,
    ast.Call,
    ast.Name,
    ast.Load,
    ast.Constant,
    ast.Add,
    ast.Sub,
    ast.Mult,
    ast.Div,
    ast.USub,
    ast.UAdd,
    ast.Eq,
    ast.NotEq,
    ast.Lt,
    ast.LtE,
    ast.Gt,
    ast.GtE,
)


def _aehnlichster_name(name: str, kandidaten: Any) -> str | None:
    treffer = difflib.get_close_matches(name, list(kandidaten), n=1, cutoff=0.6)
    return treffer[0] if treffer else None


def validiere_formel(text: str) -> str:
    """Prüft eine Formel auf Syntax, erlaubte Variablen/Funktionen/Operatoren.

    Gibt bei Erfolg den normalisierten (auswertbaren) Ausdruck zurück.
    Wirft andernfalls :class:`FormelFehler` mit einer für Fachanwender
    verständlichen Meldung.
    """

    if not text or not text.strip():
        raise FormelFehler(
            "Die Formel ist leer. Bitte einen Ausdruck eingeben, "
            "z. B. 'Bestellmenge * Faktor_VJ'."
        )

    normalisiert = normalisiere_formel(text)

    offene = normalisiert.count("(")
    schliessende = normalisiert.count(")")
    if offene != schliessende:
        raise FormelFehler(
            f"Klammer nicht geschlossen: {offene} öffnende und "
            f"{schliessende} schließende Klammer(n) gefunden."
        )

    try:
        baum = ast.parse(normalisiert, mode="eval")
    except SyntaxError as exc:
        raise FormelFehler(
            f"Die Formel ist nicht korrekt aufgebaut (Syntaxfehler nahe "
            f"Position {exc.offset}): {exc.msg}"
        ) from exc
    except ValueError as exc:  # z. B. ungültige Zahlenliterale
        raise FormelFehler(f"Die Formel enthält ein ungültiges Zahlenformat: {exc}") from exc

    # Namen, die als Funktionsname direkt vor einer Klammer stehen (z. B. "MAX"
    # in "MAX(...)"), separat behandeln - das sind Funktionsnamen, keine Variablen.
    funktionsnamen_knoten = {
        id(knoten.func)
        for knoten in ast.walk(baum)
        if isinstance(knoten, ast.Call) and isinstance(knoten.func, ast.Name)
    }

    for knoten in ast.walk(baum):
        if not isinstance(knoten, _ERLAUBTE_KNOTEN):
            raise FormelFehler(
                f"Die Formel enthält ein nicht erlaubtes Konstrukt "
                f"({type(knoten).__name__}). Erlaubt sind Zahlen, Variablen, "
                f"+ - * / ( ), Vergleiche und die Funktionen "
                f"{', '.join(sorted(ERLAUBTE_FUNKTIONEN))}."
            )
        if isinstance(knoten, ast.Call):
            if not isinstance(knoten.func, ast.Name):
                raise FormelFehler("Funktionsaufrufe sind nur mit einfachen Funktionsnamen erlaubt.")
            if knoten.func.id not in ERLAUBTE_FUNKTIONEN:
                vorschlag = _aehnlichster_name(knoten.func.id, ERLAUBTE_FUNKTIONEN)
                hinweis = f" Meinten Sie '{vorschlag}'?" if vorschlag else ""
                raise FormelFehler(f"Unbekannte Funktion '{knoten.func.id}'.{hinweis}")
            if knoten.keywords:
                raise FormelFehler("Benannte Argumente (z. B. 'x=1') sind nicht erlaubt.")
        elif isinstance(knoten, ast.Name) and id(knoten) not in funktionsnamen_knoten:
            if knoten.id not in VARIABLEN_INFO:
                vorschlag = _aehnlichster_name(knoten.id, VARIABLEN_INFO)
                hinweis = f" Meinten Sie '{vorschlag}'?" if vorschlag else ""
                raise FormelFehler(f"Unbekannte Variable '{knoten.id}'.{hinweis}")

    return normalisiert


def werte_formel_aus(text: str, werte: dict[str, Any]) -> float:
    """Validiert und berechnet eine Formel mit den gegebenen Werten.

    Nutzt ausschließlich ``simpleeval`` mit einer festen Positivliste an
    Variablen (``werte``) und Funktionen (``ERLAUBTE_FUNKTIONEN``) -
    niemals ``eval()``/``exec()``.
    """

    normalisiert = validiere_formel(text)

    auswerter = simpleeval.SimpleEval(names=dict(werte), functions=dict(ERLAUBTE_FUNKTIONEN))
    try:
        ergebnis = auswerter.eval(normalisiert)
    except ZeroDivisionError as exc:
        raise FormelFehler("Division durch Null in der Formel.") from exc
    except simpleeval.InvalidExpression as exc:
        raise FormelFehler(f"Die Formel konnte nicht berechnet werden: {exc}") from exc

    if isinstance(ergebnis, bool) or not isinstance(ergebnis, (int, float)):
        raise FormelFehler(
            "Das Ergebnis der Formel ist keine Zahl (z. B. ein Vergleich ohne WENN). "
            "Bitte die Formel prüfen."
        )
    return float(ergebnis)


# ---------------------------------------------------------------------------
# Formelbibliothek (Datenstruktur, Standardformeln, eigene Formeln)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Formel:
    """Eine Produktionsformel: Name, Excel-ähnliche Schreibweise, Erklärung."""

    name: str
    formel: str
    beschreibung: str
    eingebaut: bool = False


STANDARD_FORMEL_NAME = "Standard (Historie gewichtet)"

STANDARDFORMELN: list[Formel] = [
    Formel(
        name=STANDARD_FORMEL_NAME,
        formel="Bestellmenge * (Gewichtung * Faktor_VJ + (1 - Gewichtung)) * (1 + Zuschlag)",
        beschreibung=(
            "Mischt Bestellmenge und Vorjahresabweichung je nach eingestellter "
            "Gewichtung und addiert den Sicherheitszuschlag. Bei Gewichtung 0 % "
            "entspricht das Ergebnis genau der Bestellmenge, bei 100 % wird die "
            "Bestellmenge komplett mit dem Vorjahresfaktor skaliert. Empfehlung: "
            "Standardformel für die meisten Wochen, da sie Planungssicherheit "
            "(Bestellmenge) und Erfahrungswerte (Vorjahr) kombiniert."
        ),
        eingebaut=True,
    ),
    Formel(
        name="Nur Bestellmenge",
        formel="Bestellmenge * (1 + Zuschlag)",
        beschreibung=(
            "Ignoriert die Vorjahreswerte komplett und produziert nur die "
            "Bestellmenge zzgl. Sicherheitszuschlag. Sinnvoll, wenn keine "
            "verlässlichen Vorjahreswerte vorliegen oder sich die Nachfrage "
            "stark verändert hat (z. B. neues Sortiment, neuer Kunde)."
        ),
        eingebaut=True,
    ),
    Formel(
        name="Voller Vorjahresfaktor",
        formel="Bestellmenge * Faktor_VJ",
        beschreibung=(
            "Skaliert die Bestellmenge vollständig mit der Vorjahresabweichung "
            "(ohne Zuschlag). Sinnvoll, wenn sich Nachbestellungen/Stornierungen "
            "im Muster sehr ähnlich zum Vorjahr verhalten."
        ),
        eingebaut=True,
    ),
    Formel(
        name="Maximum aus Bestellung und Vorjahr",
        formel="MAX(Bestellmenge; Bestellmenge * Faktor_VJ)",
        beschreibung=(
            "Produziert immer mindestens die Bestellmenge, zusätzlich aber auch "
            "mehr, falls der Vorjahresfaktor über 1,0 lag. Sinnvoll, um "
            "Lieferausfälle bei Mehrabnahme zu vermeiden, auf Kosten von "
            "möglichem Überbestand."
        ),
        eingebaut=True,
    ),
    Formel(
        name="Vorjahres-Ist mit Trend",
        formel="Ist_VJ * (Bestellmenge / Bestellmenge_VJ)",
        beschreibung=(
            "Nimmt die tatsächliche Vorjahresmenge als Basis und skaliert sie mit "
            "der Veränderung der Bestellmenge gegenüber dem Vorjahr. Sinnvoll, "
            "wenn der Vorjahres-Ist-Wert als 'wahrer' Bedarf angesehen wird und "
            "nur der diesjährige Trend (mehr/weniger bestellt) berücksichtigt "
            "werden soll."
        ),
        eingebaut=True,
    ),
    Formel(
        name="Begrenzter Faktor",
        formel="Bestellmenge * MIN(MAX(Faktor_VJ; 0,9); 1,3)",
        beschreibung=(
            "Wie 'Voller Vorjahresfaktor', aber der Faktor wird auf den Bereich "
            "0,9 bis 1,3 gekappt. Verhindert, dass ein einmaliger Ausreißer im "
            "Vorjahr (z. B. Reklamation oder Sonderaktion) die Planung zu stark "
            "verzerrt."
        ),
        eingebaut=True,
    ),
    Formel(
        name="Wochentagsabhängig",
        formel='WENN(Wochentag = "Sa"; Bestellmenge * 1,15; Bestellmenge * Faktor_VJ)',
        beschreibung=(
            "Setzt am Samstag (typischer Tag mit spontaner Mehrabnahme vor "
            "Weihnachten) pauschal 15 % Aufschlag auf die Bestellmenge, an den "
            "übrigen Tagen den Vorjahresfaktor. Sinnvoll, wenn einzelne "
            "Wochentage ein bekanntes, wiederkehrendes Abweichungsmuster haben."
        ),
        eingebaut=True,
    ),
]


def formel_by_name(name: str, eigene_formeln: list[Formel] | None = None) -> Formel | None:
    """Sucht eine Formel per Name in Standard- und eigenen Formeln."""

    for formel in STANDARDFORMELN + list(eigene_formeln or []):
        if formel.name == name:
            return formel
    return None


def alle_formelnamen(eigene_formeln: list[Formel] | None = None) -> list[str]:
    return [f.name for f in STANDARDFORMELN] + [f.name for f in (eigene_formeln or [])]


# -- Speichern / Laden eigener Formeln (formeln.json) ------------------------

PROJEKT_ORDNER = Path(__file__).parent
STANDARD_FORMELN_DATEI = PROJEKT_ORDNER / "formeln.json"


def lade_eigene_formeln(pfad: Path = STANDARD_FORMELN_DATEI) -> list[Formel]:
    """Lädt die eigenen (nutzerdefinierten) Formeln aus einer JSON-Datei."""

    if not pfad.exists():
        return []
    try:
        rohdaten = json.loads(pfad.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    formeln = []
    for eintrag in rohdaten:
        formeln.append(
            Formel(
                name=eintrag["name"],
                formel=eintrag["formel"],
                beschreibung=eintrag.get("beschreibung", ""),
                eingebaut=False,
            )
        )
    return formeln


def _schreibe_eigene_formeln(formeln: list[Formel], pfad: Path) -> None:
    rohdaten = [
        {"name": f.name, "formel": f.formel, "beschreibung": f.beschreibung}
        for f in formeln
        if not f.eingebaut
    ]
    pfad.write_text(json.dumps(rohdaten, ensure_ascii=False, indent=2), encoding="utf-8")


def speichere_eigene_formel(
    formel: Formel, pfad: Path = STANDARD_FORMELN_DATEI
) -> list[Formel]:
    """Speichert eine eigene Formel (neu oder Überschreiben einer gleichnamigen
    eigenen Formel). Vorgefertigte Formeln können nicht überschrieben werden -
    das Original bleibt beim 'Als Vorlage laden -> anpassen -> neu speichern'
    immer unverändert, da eine neue Formel mit neuem Namen entsteht.
    """

    if formel_by_name(formel.name) and any(
        f.name == formel.name and f.eingebaut for f in STANDARDFORMELN
    ):
        raise FormelFehler(
            f"Der Name '{formel.name}' ist bereits durch eine vorgefertigte Formel "
            f"belegt. Bitte einen anderen Namen wählen."
        )
    validiere_formel(formel.formel)

    formeln = [f for f in lade_eigene_formeln(pfad) if f.name != formel.name]
    formeln.append(Formel(formel.name, formel.formel, formel.beschreibung, eingebaut=False))
    _schreibe_eigene_formeln(formeln, pfad)
    return formeln


def loesche_eigene_formel(name: str, pfad: Path = STANDARD_FORMELN_DATEI) -> list[Formel]:
    """Entfernt eine eigene Formel per Name. Vorgefertigte Formeln können nicht
    gelöscht werden - für diese wird einfach nichts getan.
    """

    formeln = [f for f in lade_eigene_formeln(pfad) if f.name != name]
    _schreibe_eigene_formeln(formeln, pfad)
    return formeln


# ---------------------------------------------------------------------------
# Kalender-Hilfsfunktionen: Wochentag, Tage bis Heiligabend
# ---------------------------------------------------------------------------

WOCHENTAGSKUERZEL = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
ARBEITSTAGE = ["Mo", "Di", "Mi", "Do", "Fr", "Sa"]  # Montag bis Samstag


def wochentag_kuerzel(datum: date) -> str:
    return WOCHENTAGSKUERZEL[datum.weekday()]


def tage_bis_heiligabend(datum: date) -> int:
    """Anzahl Tage von ``datum`` bis Heiligabend (24.12.) desselben Jahres.
    Negative Werte bedeuten: Heiligabend liegt bereits in der Vergangenheit.
    """

    heiligabend = date(datum.year, 12, 24)
    return (heiligabend - datum).days


def erzeuge_wochen_vorlage(
    jahr: int, kalenderwoche: int, beispielmengen: dict[str, float] | None = None
) -> list[dict[str, Any]]:
    """Erzeugt die Montag-Samstag-Vorbelegung der Eingabetabelle für eine
    bestimmte Kalenderwoche (Datum + Wochentag vorbefüllt, Bestellmenge
    entweder aus ``beispielmengen`` oder 0.0).
    """

    zeilen = []
    for index, kuerzel in enumerate(ARBEITSTAGE, start=1):
        datum = date.fromisocalendar(jahr, kalenderwoche, index)
        menge = (beispielmengen or {}).get(kuerzel, 0.0)
        zeilen.append({"Datum": datum, "Wochentag": kuerzel, "Bestellmenge (kg)": float(menge)})
    return zeilen


# ---------------------------------------------------------------------------
# Historiendaten (Vorjahr)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class HistorienTag:
    """Ein historischer Tag des Vorjahres."""

    datum: date
    wochentag: str
    bestellmenge: float
    ist: float

    @property
    def tage_bis_heiligabend(self) -> int:
        return tage_bis_heiligabend(self.datum)

    @property
    def faktor_vj(self) -> float | None:
        if not self.bestellmenge:
            return None
        return self.ist / self.bestellmenge


def erzeuge_beispielhistorie(jahr: int, seed: int = 42) -> list[HistorienTag]:
    """Erzeugt realistische Beispiel-Historiendaten für KW 49-52 eines Jahres,
    falls noch keine echten Vorjahreswerte vorliegen. Die Mengen steigen zur
    Weihnachtswoche hin an, der Samstag vor Heiligabend hat eine etwas
    unruhigere (höhere) Mehrabnahme.
    """

    zufall = random.Random(seed)
    basis_je_wochentag = {"Mo": 380.0, "Di": 360.0, "Mi": 400.0, "Do": 420.0, "Fr": 480.0, "Sa": 520.0}

    ergebnisse: list[HistorienTag] = []
    for kalenderwoche in (49, 50, 51, 52):
        woche_vor_weihnachten = 52 - kalenderwoche  # 0 = Weihnachtswoche, 3 = KW49
        steigerung = 1.0 + (3 - woche_vor_weihnachten) * 0.12
        for index, kuerzel in enumerate(ARBEITSTAGE, start=1):
            try:
                datum = date.fromisocalendar(jahr, kalenderwoche, index)
            except ValueError:
                continue
            bestellmenge = round(basis_je_wochentag[kuerzel] * steigerung * zufall.uniform(0.92, 1.08), 1)
            abweichung = zufall.uniform(0.88, 1.18)
            if kuerzel == "Sa" and woche_vor_weihnachten == 0:
                abweichung *= zufall.uniform(1.0, 1.12)
            ist = round(bestellmenge * abweichung, 1)
            ergebnisse.append(HistorienTag(datum=datum, wochentag=kuerzel, bestellmenge=bestellmenge, ist=ist))
    return ergebnisse


def finde_vorjahresvergleichstag(
    wochentag: str, tage_bis_heiligabend_wert: int, historie: list[HistorienTag]
) -> HistorienTag | None:
    """Findet den zum aktuellen Tag passenden Vergleichstag im Vorjahr.

    Zuordnung erfolgt NICHT nach Kalenderdatum, sondern nach gleichem
    Wochentag und dem ähnlichsten Abstand zu Heiligabend - so verschiebt
    sich der Vergleich automatisch mit, wenn die Wochentage vor
    Weihnachten von Jahr zu Jahr anders liegen.
    """

    kandidaten = [h for h in historie if h.wochentag == wochentag]
    if not kandidaten:
        return None
    return min(kandidaten, key=lambda h: abs(h.tage_bis_heiligabend - tage_bis_heiligabend_wert))


# ---------------------------------------------------------------------------
# Eingabe-Prüfung (Zahlenformat, Plausibilität)
# ---------------------------------------------------------------------------


def parse_zahl(wert: Any) -> float:
    """Wandelt eine Benutzereingabe (Zahl oder Text) in eine float-Zahl um.

    Akzeptiert sowohl deutsche (1.234,56) als auch internationale
    (1234.56 bzw. 1234,56) Schreibweisen.
    """

    if wert is None:
        raise ValueError("Leerer Wert")
    if isinstance(wert, bool):
        raise ValueError(f"'{wert}' ist keine gültige Zahl")
    if isinstance(wert, (int, float)):
        if isinstance(wert, float) and math.isnan(wert):
            raise ValueError("Leerer Wert")
        return float(wert)

    text = str(wert).strip().replace(" ", "")
    if text == "":
        raise ValueError("Leerer Wert")

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError as exc:
        raise ValueError(f"'{wert}' ist keine gültige Zahl") from exc


def validiere_eingabetabelle(zeilen: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[str]]:
    """Prüft die Bestellmengen-Eingabetabelle (Datum, Wochentag, Bestellmenge).

    Gibt eine bereinigte Zeilenliste (Bestellmenge als float) sowie eine
    Liste verständlicher Fehlermeldungen zurück. Fehlerhafte Zeilen werden
    aus der bereinigten Liste ausgeschlossen, damit die restliche
    Berechnung trotzdem möglich ist.
    """

    bereinigt = []
    fehler = []
    for i, zeile in enumerate(zeilen, start=1):
        datum = zeile.get("Datum")
        wochentag = zeile.get("Wochentag", "?")
        roh = zeile.get("Bestellmenge (kg)", zeile.get("Bestellmenge"))
        bezeichner = f"Zeile {i} ({wochentag}, {datum})"
        try:
            menge = parse_zahl(roh)
        except ValueError:
            fehler.append(f"{bezeichner}: '{roh}' ist keine gültige Zahl.")
            continue
        if menge < 0:
            fehler.append(f"{bezeichner}: Bestellmenge darf nicht negativ sein ({menge:g} kg).")
            continue
        bereinigt.append({"Datum": datum, "Wochentag": wochentag, "Bestellmenge (kg)": menge})
    return bereinigt, fehler


# ---------------------------------------------------------------------------
# Rundung auf Gebinde-/Chargengröße
# ---------------------------------------------------------------------------


def runde_auf_gebinde(menge: float, gebindegroesse: float) -> float:
    """Rundet eine Menge auf das nächste Vielfache der Gebindegröße.

    Eine Gebindegröße von 0 (oder None) bedeutet 'keine Rundung'.
    """

    if not gebindegroesse or gebindegroesse <= 0:
        return float(menge)
    return float(round(menge / gebindegroesse) * gebindegroesse)


# ---------------------------------------------------------------------------
# Kernberechnung pro Tag und pro Woche
# ---------------------------------------------------------------------------


@dataclass
class Tagesergebnis:
    datum: date
    wochentag: str
    tage_bis_heiligabend: int
    bestellmenge: float
    bestellmenge_vj: float | None
    ist_vj: float | None
    faktor_vj: float | None
    produktionsmenge: float
    produktionsmenge_gerundet: float
    differenz: float
    differenz_prozent: float | None
    hinweise: list[str] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["hinweise"] = "; ".join(self.hinweise)
        return d


PLAUSIBLER_FAKTOR_MIN = 0.5
PLAUSIBLER_FAKTOR_MAX = 2.0
STARKE_ABWEICHUNG_SCHWELLE = 0.5  # 50 %


def berechne_tag(
    *,
    datum: date,
    wochentag: str,
    bestellmenge: float,
    bestellmenge_vj: float | None,
    ist_vj: float | None,
    gewichtung: float,
    zuschlag: float,
    formel_text: str,
    gebindegroesse: float,
    tage_bis_heiligabend_wert: int,
) -> Tagesergebnis:
    """Berechnet die Produktionsmenge für einen einzelnen Tag."""

    hinweise: list[str] = []

    if bestellmenge < 0:
        hinweise.append("Bestellmenge ist negativ - Wert bitte prüfen.")

    if bestellmenge_vj is not None and bestellmenge_vj < 0:
        hinweise.append("Vorjahres-Bestellmenge ist negativ - Wert bitte prüfen.")
    if ist_vj is not None and ist_vj < 0:
        hinweise.append("Vorjahres-Ist-Menge ist negativ - Wert bitte prüfen.")

    if bestellmenge_vj is None or ist_vj is None:
        hinweise.append(
            "Kein passender Vorjahreswert gefunden - es wird ein neutraler "
            "Faktor von 1,0 angenommen (Produktion = reine Bestellmenge)."
        )
        faktor_vj: float | None = None
        faktor_fuer_formel = 1.0
    elif not bestellmenge_vj:
        hinweise.append("Vorjahres-Bestellmenge ist 0 - Faktor Vorjahr kann nicht berechnet werden, Faktor 1,0 angenommen.")
        faktor_vj = None
        faktor_fuer_formel = 1.0
    else:
        faktor_vj = ist_vj / bestellmenge_vj
        faktor_fuer_formel = faktor_vj
        if not (PLAUSIBLER_FAKTOR_MIN <= faktor_vj <= PLAUSIBLER_FAKTOR_MAX):
            hinweise.append(
                f"Vorjahresfaktor {faktor_vj:.2f} liegt außerhalb des üblichen "
                f"Bereichs ({PLAUSIBLER_FAKTOR_MIN:g}-{PLAUSIBLER_FAKTOR_MAX:g}) - "
                f"bitte Vorjahreswerte prüfen."
            )

    werte = {
        "Bestellmenge": bestellmenge,
        "Bestellmenge_VJ": bestellmenge_vj if bestellmenge_vj is not None else 0.0,
        "Ist_VJ": ist_vj if ist_vj is not None else 0.0,
        "Faktor_VJ": faktor_fuer_formel,
        "Gewichtung": gewichtung,
        "Zuschlag": zuschlag,
        "Wochentag": wochentag,
        "Tage_bis_Heiligabend": tage_bis_heiligabend_wert,
    }

    try:
        roh_ergebnis = werte_formel_aus(formel_text, werte)
    except FormelFehler as exc:
        hinweise.append(f"Formel-Fehler, Bestellmenge wurde unverändert übernommen: {exc}")
        roh_ergebnis = bestellmenge

    if roh_ergebnis < 0:
        hinweise.append("Berechnetes Ergebnis ist negativ - auf 0 kg gesetzt.")
        roh_ergebnis = 0.0

    gerundet = runde_auf_gebinde(roh_ergebnis, gebindegroesse)

    if bestellmenge > 0 and abs(gerundet - bestellmenge) / bestellmenge > STARKE_ABWEICHUNG_SCHWELLE:
        hinweise.append(
            f"Ergebnis weicht stark von der Bestellmenge ab "
            f"({(gerundet - bestellmenge) / bestellmenge:+.0%}) - bitte Formel/Parameter prüfen."
        )

    differenz = gerundet - bestellmenge
    differenz_prozent = (differenz / bestellmenge * 100) if bestellmenge else None

    return Tagesergebnis(
        datum=datum,
        wochentag=wochentag,
        tage_bis_heiligabend=tage_bis_heiligabend_wert,
        bestellmenge=bestellmenge,
        bestellmenge_vj=bestellmenge_vj,
        ist_vj=ist_vj,
        faktor_vj=faktor_vj,
        produktionsmenge=roh_ergebnis,
        produktionsmenge_gerundet=gerundet,
        differenz=differenz,
        differenz_prozent=differenz_prozent,
        hinweise=hinweise,
    )


def berechne_woche(
    aktuelle_tage: list[dict[str, Any]],
    historie: list[HistorienTag],
    formel_text: str,
    gewichtung: float,
    zuschlag: float,
    gebindegroesse: float,
) -> list[Tagesergebnis]:
    """Berechnet die Produktionsmengen für eine komplette Woche (Mo-Sa)."""

    ergebnisse = []
    for tag in aktuelle_tage:
        datum = parse_datum(tag["Datum"])
        if datum is None:
            raise ValueError(f"Ungültiges Datum in der Eingabetabelle: {tag['Datum']!r}")
        wochentag = tag.get("Wochentag") or wochentag_kuerzel(datum)
        bestellmenge = parse_zahl(tag["Bestellmenge (kg)"])
        tage_heiligabend = tage_bis_heiligabend(datum)
        vergleichstag = finde_vorjahresvergleichstag(wochentag, tage_heiligabend, historie)
        ergebnisse.append(
            berechne_tag(
                datum=datum,
                wochentag=wochentag,
                bestellmenge=bestellmenge,
                bestellmenge_vj=vergleichstag.bestellmenge if vergleichstag else None,
                ist_vj=vergleichstag.ist if vergleichstag else None,
                gewichtung=gewichtung,
                zuschlag=zuschlag,
                formel_text=formel_text,
                gebindegroesse=gebindegroesse,
                tage_bis_heiligabend_wert=tage_heiligabend,
            )
        )
    return ergebnisse


def berechne_wochensumme(ergebnisse: list[Tagesergebnis]) -> dict[str, float | None]:
    """Fasst eine Liste von Tagesergebnissen zu Wochenkennzahlen zusammen."""

    summe_bestellung = sum(t.bestellmenge for t in ergebnisse)
    summe_produktion = sum(t.produktionsmenge_gerundet for t in ergebnisse)
    ist_vj_werte = [t.ist_vj for t in ergebnisse if t.ist_vj is not None]
    summe_ist_vj = sum(ist_vj_werte) if ist_vj_werte else None
    differenz = summe_produktion - summe_bestellung
    differenz_prozent = (differenz / summe_bestellung * 100) if summe_bestellung else None
    return {
        "summe_bestellung": summe_bestellung,
        "summe_produktion": summe_produktion,
        "summe_ist_vj": summe_ist_vj,
        "differenz": differenz,
        "differenz_prozent": differenz_prozent,
    }


def vergleiche_formeln(
    formelnamen: list[str],
    aktuelle_tage: list[dict[str, Any]],
    historie: list[HistorienTag],
    gewichtung: float,
    zuschlag: float,
    gebindegroesse: float,
    eigene_formeln: list[Formel] | None = None,
) -> dict[str, list[Tagesergebnis]]:
    """Berechnet eine Woche für mehrere Formeln gleichzeitig (Formelvergleich)."""

    ergebnis: dict[str, list[Tagesergebnis]] = {}
    for name in formelnamen:
        formel = formel_by_name(name, eigene_formeln)
        if formel is None:
            continue
        ergebnis[name] = berechne_woche(
            aktuelle_tage, historie, formel.formel, gewichtung, zuschlag, gebindegroesse
        )
    return ergebnis


# ---------------------------------------------------------------------------
# CSV / Excel: Einlesen von Datei-Uploads
# ---------------------------------------------------------------------------


def lese_csv_bytes(daten: bytes) -> list[dict[str, str]]:
    """Liest CSV-Rohdaten (Komma oder Semikolon als Trennzeichen) in eine
    Liste von Dictionaries ein (alle Werte zunächst als Text).
    """

    text = daten.decode("utf-8-sig")
    try:
        dialekt = csv.Sniffer().sniff(text.splitlines()[0], delimiters=";,")
        trennzeichen = dialekt.delimiter
    except (csv.Error, IndexError):
        trennzeichen = ";" if text.splitlines() and ";" in text.splitlines()[0] else ","
    reader = csv.DictReader(io.StringIO(text), delimiter=trennzeichen)
    return [dict(row) for row in reader]


def lese_excel_bytes(daten: bytes) -> list[dict[str, Any]]:
    """Liest das erste Tabellenblatt einer Excel-Datei (Rohbytes) über
    openpyxl in eine Liste von Dictionaries ein (erste Zeile = Spaltennamen).
    """

    import openpyxl

    arbeitsbuch = openpyxl.load_workbook(io.BytesIO(daten), data_only=True)
    blatt = arbeitsbuch.active
    zeilen_iter = blatt.iter_rows(values_only=True)
    try:
        kopfzeile = [str(k) if k is not None else "" for k in next(zeilen_iter)]
    except StopIteration:
        return []
    ergebnis = []
    for zeile in zeilen_iter:
        if all(wert is None for wert in zeile):
            continue
        ergebnis.append(dict(zip(kopfzeile, zeile)))
    return ergebnis


def zeilen_zu_aktueller_woche(zeilen: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Wandelt eingelesene Zeilen (Datum, Wochentag, Bestellmenge) in das von
    der Eingabetabelle erwartete Format um. Fehlende Wochentage werden aus
    dem Datum abgeleitet, fehlende Daten führen zum Überspringen der Zeile.
    """

    ergebnis = []
    for zeile in zeilen:
        roh_datum = zeile.get("Datum")
        datum = parse_datum(roh_datum)
        if datum is None:
            continue
        menge_roh = zeile.get("Bestellmenge (kg)", zeile.get("Bestellmenge"))
        try:
            menge = parse_zahl(menge_roh)
        except ValueError:
            menge = 0.0
        wochentag = zeile.get("Wochentag") or wochentag_kuerzel(datum)
        ergebnis.append({"Datum": datum, "Wochentag": wochentag, "Bestellmenge (kg)": menge})
    return ergebnis


def zeilen_zu_historie(zeilen: list[dict[str, Any]]) -> list[HistorienTag]:
    """Wandelt eingelesene Zeilen (Datum, Wochentag, Bestellmenge_VJ, Ist_VJ)
    in eine Liste von :class:`HistorienTag` um.
    """

    historie = []
    for zeile in zeilen:
        datum = parse_datum(zeile.get("Datum"))
        if datum is None:
            continue
        try:
            bestellmenge = parse_zahl(zeile.get("Bestellmenge_VJ"))
            ist = parse_zahl(zeile.get("Ist_VJ"))
        except ValueError:
            continue
        wochentag = zeile.get("Wochentag") or wochentag_kuerzel(datum)
        historie.append(HistorienTag(datum=datum, wochentag=wochentag, bestellmenge=bestellmenge, ist=ist))
    return historie


def parse_datum(wert: Any) -> date | None:
    """Wandelt ein Datum (date-Objekt, ISO-Text oder deutsches Datumsformat)
    in ein ``date``-Objekt um. Wird u. a. benötigt, da UI-Elemente wie
    ``mo.ui.data_editor`` ein Datum ggf. als Text statt als ``date``-Objekt
    zurückgeben.
    """

    if wert is None:
        return None
    if isinstance(wert, date):
        return wert
    text = str(wert).strip()
    for muster in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            from datetime import datetime

            return datetime.strptime(text, muster).date()
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# Excel-Export des Ergebnisses
# ---------------------------------------------------------------------------


def exportiere_ergebnis_als_excel(
    ergebnisse: list[Tagesergebnis],
    wochensumme: dict[str, float | None],
    formel_name: str,
    formel_text: str,
    parameter: dict[str, Any],
) -> bytes:
    """Erstellt eine Excel-Datei (zwei Tabellenblätter) mit dem Ergebnis der
    Wochenplanung sowie der verwendeten Formel und den Parametern, damit die
    Planung für andere nachvollziehbar bleibt. Gibt die Datei als Bytes
    zurück (für ``mo.download``).
    """

    import openpyxl
    from openpyxl.styles import Font

    arbeitsbuch = openpyxl.Workbook()

    # -- Blatt 1: Ergebnis pro Tag -------------------------------------------------
    blatt_ergebnis = arbeitsbuch.active
    blatt_ergebnis.title = "Ergebnis"
    spalten = [
        "Datum",
        "Wochentag",
        "Bestellmenge (kg)",
        "Bestellmenge Vorjahr (kg)",
        "Ist Vorjahr (kg)",
        "Faktor Vorjahr",
        "Produktionsmenge (kg)",
        "Produktionsmenge gerundet (kg)",
        "Differenz (kg)",
        "Differenz (%)",
        "Hinweise",
    ]
    blatt_ergebnis.append(spalten)
    for zelle in blatt_ergebnis[1]:
        zelle.font = Font(bold=True)

    for t in ergebnisse:
        blatt_ergebnis.append(
            [
                t.datum.isoformat() if isinstance(t.datum, date) else t.datum,
                t.wochentag,
                t.bestellmenge,
                t.bestellmenge_vj,
                t.ist_vj,
                round(t.faktor_vj, 4) if t.faktor_vj is not None else None,
                round(t.produktionsmenge, 2),
                round(t.produktionsmenge_gerundet, 2),
                round(t.differenz, 2),
                round(t.differenz_prozent, 2) if t.differenz_prozent is not None else None,
                "; ".join(t.hinweise),
            ]
        )

    letzte_zeile = len(ergebnisse) + 2
    blatt_ergebnis.append(
        [
            "Wochensumme",
            "",
            wochensumme.get("summe_bestellung"),
            "",
            wochensumme.get("summe_ist_vj"),
            "",
            "",
            wochensumme.get("summe_produktion"),
            wochensumme.get("differenz"),
            wochensumme.get("differenz_prozent"),
            "",
        ]
    )
    for zelle in blatt_ergebnis[letzte_zeile]:
        zelle.font = Font(bold=True)

    for spalte in blatt_ergebnis.columns:
        breite = max((len(str(z.value)) for z in spalte if z.value is not None), default=10)
        blatt_ergebnis.column_dimensions[spalte[0].column_letter].width = min(max(breite + 2, 10), 40)

    # -- Blatt 2: Formel & Parameter -----------------------------------------------
    blatt_formel = arbeitsbuch.create_sheet("Formel & Parameter")
    blatt_formel.append(["Verwendete Formel", formel_name])
    blatt_formel.append(["Formel (Excel-Schreibweise)", formel_text])
    blatt_formel.append([])
    blatt_formel.append(["Parameter", "Wert"])
    for zelle in blatt_formel[4]:
        zelle.font = Font(bold=True)
    for name, wert in parameter.items():
        blatt_formel.append([name, wert])
    blatt_formel.column_dimensions["A"].width = 30
    blatt_formel.column_dimensions["B"].width = 60

    puffer = io.BytesIO()
    arbeitsbuch.save(puffer)
    return puffer.getvalue()
