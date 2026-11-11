# Produktionsplanung Schweinefilet - Weihnachtswoche

Marimo-Notebook, das die Excel/Sheets-Planung der Produktionsmenge für
Schweinefilet während der Weihnachtszeit ersetzt. Der Fachanwender trägt
die Bestellmengen Montag-Samstag ein, wählt eine Formel (oder schreibt
eine eigene) und klickt auf **Berechnen** - wie in Excel, nur reaktiv und
nachvollziehbar.

## Setup (mit uv)

```bash
# Projekt wurde bereits mit folgenden Befehlen angelegt:
uv init schweinefilet-planung
cd schweinefilet-planung
uv add marimo polars altair openpyxl simpleeval
uv add --dev pytest
```

Zum Weiterarbeiten reicht:

```bash
cd schweinefilet-planung
uv sync
```

## Notebook starten

**Zum Bearbeiten/Entwickeln** (Code-Zellen sichtbar, volle Reaktivität):

```bash
uv run marimo edit planung.py
```

**Als App für Fachanwender** (keine Code-Zellen sichtbar, nur die Oberfläche):

```bash
uv run marimo run planung.py
```

**Ohne lokale Installation**, direkt per `uvx` und den im Notebook
hinterlegten PEP-723-Abhängigkeiten (lädt automatisch eine isolierte
Umgebung):

```bash
uvx marimo edit --sandbox planung.py
# oder als App:
uvx marimo run --sandbox planung.py
```

## Tests ausführen

```bash
uv run pytest
```

Die Tests prüfen die komplette Berechnungslogik und den Formel-Auswerter
in `formeln.py`, unter anderem auch, dass unerlaubte Ausdrücke wie
`__import__(...)`, `(1).__class__`, Listen-Comprehensions oder `eval(...)`
zuverlässig abgelehnt werden.

## Projektdateien

| Datei | Zweck |
|---|---|
| `planung.py` | Das Marimo-Notebook (Oberfläche, Bedienung, Darstellung) |
| `formeln.py` | Formelbibliothek, sicherer Formel-Auswerter, Berechnungslogik (reine Python-Funktionen, kein UI-Code) |
| `test_formeln.py` | pytest-Tests für `formeln.py` |
| `historie_vorjahr.csv` | Beispiel-Historiendaten des Vorjahres (KW 49-52) |
| `bestellung_aktuell.csv` | Beispiel-Bestellmengen einer aktuellen Planungswoche |
| `formeln.json` | Beispiele für gespeicherte eigene Formeln |

## Wie die Produktionsmenge berechnet wird

### 1. Abgleich mit dem Vorjahr

Die Tage der aktuellen Woche werden **nicht** über das Kalenderdatum dem
Vorjahr zugeordnet, sondern über **gleichen Wochentag** und **ähnlichem
Abstand zu Heiligabend** (24.12.). Das ist nötig, weil sich die
Wochentage vor Weihnachten jedes Jahr verschieben - der "Freitag vor
Heiligabend" fällt nicht jedes Jahr auf dasselbe Kalenderdatum.

### 2. Abweichungsfaktor Vorjahr

Für jeden passenden Vorjahrestag wird berechnet:

```
Faktor_VJ = Ist_VJ / Bestellmenge_VJ
```

Ein Faktor von `1,1` bedeutet: Im Vorjahr wurde an diesem Tag 10 % mehr
produziert als ursprünglich bestellt (z. B. durch Nachbestellungen).

### 3. Formel anwenden

Die Standardformel lautet:

```
Produktionsmenge = Bestellmenge * (Gewichtung * Faktor_VJ + (1 - Gewichtung)) * (1 + Zuschlag)
```

- **Gewichtung** (0-100 %, per Schieberegler einstellbar) steuert, wie
  stark die Vorjahreserfahrung einfließt. Bei 0 % zählt nur die
  Bestellmenge, bei 100 % wird die Bestellmenge komplett mit dem
  Vorjahresfaktor skaliert.
- **Zuschlag** ist der Sicherheitszuschlag in Prozent (z. B. `5` für 5 %).

Im Reiter **Berechnung** kann statt der Standardformel auch jede andere
Formel aus der Formelbibliothek (Dropdown) oder eine einmalige eigene
Formel (Textfeld) gewählt werden - siehe nächster Abschnitt.

### 4. Rundung auf Gebinde-/Chargengröße

Das Ergebnis jeder Formel wird anschließend einheitlich auf die gewählte
Gebindegröße gerundet (z. B. auf volle 10 kg oder auf Kisten à 25 kg).

### 5. Wochensumme und Abweichung

Am Ende werden die Wochensumme der Bestellmenge und der geplanten
Produktion sowie deren Differenz (absolut und in %) ausgewiesen.

## Kurzanleitung: Eigene Formel erstellen (für Fachanwender)

Im Reiter **Formeln verwalten**:

1. **Verfügbare Variablen ansehen.** Oben steht eine Übersicht aller
   Variablen, die in einer Formel verwendet werden dürfen (z. B.
   `Bestellmenge`, `Faktor_VJ`, `Gewichtung`, `Zuschlag`, `Wochentag`,
   `Tage_bis_Heiligabend`) - mit Bedeutung und einem Beispielwert.
2. **Vorlage wählen (optional).** Im Dropdown "Formel als Vorlage wählen"
   eine vorhandene Formel auswählen und mit "Als Vorlage in das Textfeld
   laden" übernehmen. Die Original-Formel bleibt dabei unverändert -
   es wird nur eine Kopie in das Textfeld geladen.
3. **Formel schreiben oder anpassen.** Im Textfeld die Formel in
   Excel-ähnlicher Schreibweise eintragen, z. B.:

   ```
   WENN(Wochentag = "Sa"; Bestellmenge * 1,2; Bestellmenge * Faktor_VJ)
   ```

   Erlaubt sind die Grundrechenarten `+ - * / ( )`, Vergleiche
   (`=`, `<>`, `<`, `<=`, `>`, `>=`) sowie die Funktionen `MIN`, `MAX`,
   `RUNDEN`, `AUFRUNDEN`, `ABRUNDEN`, `WENN`, `ABS`, `MITTELWERT`.
   Semikolon als Argumenttrenner und Komma als Dezimaltrennzeichen (wie
   im deutschen Excel) werden genauso verstanden wie Punkt und Komma in
   internationaler Schreibweise.
4. **Live-Vorschau prüfen.** Direkt unter dem Textfeld erscheint sofort
   das Ergebnis der Formel für einen Beispieltag. Bei einem Fehler in der
   Formel (z. B. Tippfehler oder fehlende Klammer) erscheint stattdessen
   eine verständliche Fehlermeldung mit Korrekturvorschlag, z. B.:
   *"Unbekannte Variable 'Bestelmenge'. Meinten Sie 'Bestellmenge'?"*
5. **Formel speichern.** Einen Namen und eine kurze Beschreibung
   eintragen (wann ist diese Formel sinnvoll?) und auf **"Formel
   speichern"** klicken. Die Formel erscheint danach im Dropdown
   "Formel auswählen" im Reiter **Berechnung**, unter "Eigene Formeln".
6. **Formel löschen.** Weiter unten kann eine eigene Formel über das
   Dropdown ausgewählt und mit **"Formel löschen"** wieder entfernt
   werden. Vorgefertigte Formeln können nicht gelöscht werden.

**Sicherheitshinweis:** Eigene Formeln werden niemals mit `eval()` oder
`exec()` ausgeführt. Jede Formel wird zunächst geprüft (erlaubte
Variablen, Funktionen und Operatoren) und danach mit der Bibliothek
`simpleeval` berechnet. Nicht erlaubte Ausdrücke (z. B. Funktionsaufrufe,
Attributzugriffe oder Python-Sprachkonstrukte wie Listen-Comprehensions)
werden abgelehnt, bevor überhaupt eine Berechnung versucht wird.

## Formelvergleich

Im Reiter **Vergleich** können mehrere Formeln gleichzeitig ausgewählt
werden. Nach einem Klick auf "Vergleich berechnen" erscheinen die
Produktionsmengen aller gewählten Formeln nebeneinander in einer Tabelle
und in einem Diagramm, zusammen mit der Wochensumme je Formel und der
Abweichung zum tatsächlichen Vorjahresergebnis (Ist_VJ).

## Eigene Daten verwenden

- **Historiendaten des Vorjahres**: CSV oder Excel mit den Spalten
  `Datum`, `Wochentag`, `Bestellmenge_VJ`, `Ist_VJ` hochladen (siehe
  `historie_vorjahr.csv` als Vorlage). Ohne Upload werden automatisch
  plausible Beispieldaten für KW 49-52 des Vorjahres erzeugt.
- **Bestellmengen der aktuellen Woche**: entweder direkt in die Tabelle
  eintragen, oder CSV/Excel mit den Spalten `Datum`, `Wochentag`,
  `Bestellmenge (kg)` hochladen (siehe `bestellung_aktuell.csv`).

## Export

Nach dem Klick auf "Berechnen" steht ein Download-Button für eine
Excel-Datei zur Verfügung. Diese enthält zwei Tabellenblätter:

1. **Ergebnis**: Tagesergebnisse und Wochensumme.
2. **Formel & Parameter**: die verwendete Formel sowie alle Parameter
   (Gewichtung, Zuschlag, Rundung), damit die Planung für andere
   nachvollziehbar bleibt.
