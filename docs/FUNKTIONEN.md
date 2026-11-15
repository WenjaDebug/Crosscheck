# Gegenprobe: alles im Detail

Die kurze Fassung mit Bild steht im [README](../README.de.md).

Gegenprobe rechnet deine Excel-Mappe selbst nach und sagt, was nicht aufgeht. Sie arbeitet ohne KI, ohne Internet und ohne Schlüssel, und die Daten bleiben auf deinem Rechner. Deshalb passt sie auch zu heiklen Daten, etwa aus Spital, Labor oder Buchhaltung. Am Mac steht sie als Seitenfenster direkt in Excel.



Sie liest den Namen einer Kennzahl („Umsatz Ost“, „Durchschnitt Umsatz Nord“, „Umsatz Januar, Februar und Juli“) oder die Köpfe einer Übersicht und rechnet nach, welche Zeilen dazugehören. Stimmt eine Formel nicht, nennt sie die Stelle in der Formel, die das verursacht:
- eine Bedingung zu viel oder ein falscher Wert;
- ein Bereich, der eine Zeile zu früh aufhört;
- eine doppelt gezählte Zelle oder ein falscher Teiler;
- eine Bedingung aus der falschen Kopfzelle;
- eine Zahl oder ein Datum, das als Text dasteht.

Dazu gibt es die korrigierte Formel, die sie vorher selbst nachgerechnet hat. Außerdem prüft sie die ganze Rechenkette, Bereiche, die beim Herunterziehen mitrutschen, Quoten und Mittelwerte, Finanzmodelle, Daten für Auswertungen und Unterschiede zu einer früheren Version der Mappe.

*English: [FEATURES.md](FEATURES.md)*

Das Projekt hieß früher VoiceAgent und begann als Sprach-Assistent. Davon ist nichts mehr übrig; nur die Startdateien heißen noch **VoiceAgent.command**, **VoiceAgent.bat**, **ExcelAddin.command** und **ExcelAddin.bat**, und das Python-Paket heißt `voiceagent`.

## Stand

| Teil | Inhalt | Status |
|---|---|---|
| Prüfer in Excel | Seitenfenster am Mac und unter Windows, rechnet nach jeder Änderung neu, ein Klick markiert die Zelle | ✅ |
| Ursache in der Formel | Kleinste Änderung, die die Formel richtig macht, auch zwei Fehler auf einmal | ✅ |
| Datei prüfen und vergleichen | .xlsx prüfen, ohne sie zu öffnen, und mit einer früheren Version vergleichen | ✅ |
| Prüfbericht | Zum Weitergeben, auf Wunsch nur mit Zelladressen, ohne Daten | ✅ |
| Excel-Helfer | Übersichten und Tabellen aus Dateien auf Zuruf, als echte Formeln | ✅ |
| Windows | Prüfer im Browser und als Seitenfenster in Excel (ExcelAddin.bat, neu: auf einem echten Windows-PC noch nicht erprobt) | ✅ |

## Installation (macOS)

**Am einfachsten:** wie im README [in 5 Schritten](../README.de.md#in-5-schritten). **ExcelAddin.command** richtet beim ersten Mal alles ein und dann das Seitenfenster in Excel; **VoiceAgent.command** richtet beim ersten Mal alles ein und startet Gegenprobe im Browser. Beide suchen zuerst ein Python 3.10 oder neuer, das schon da ist (von python.org, von Homebrew oder sonst auf dem PATH). Das Python von macOS (`/usr/bin/python3`, meist 3.9) fragen sie nur, wenn die Befehlszeilentools installiert sind, damit macOS nicht anbietet, sie zu installieren. Fehlt ein passendes Python, installieren sie Python 3.13 mit Homebrew, falls es da ist, sonst öffnen sie python.org und sagen, was zu tun ist. Dann legen sie eine eigene Umgebung im Ordner `.venv` an und laden die drei Pakete (FastAPI, Uvicorn, pypdf). Danach startet es nur noch. Ohne Internet beim ersten Start sagen sie das. Die ZIP von GitHub behält das Ausführungsrecht der .command-Dateien (geprüft mit `git archive`, so packt GitHub). Die Sperre von macOS für Dateien aus dem Internet umgeht Gegenprobe nicht: Einmal je Datei bestätigst du sie selbst (siehe Schritt 3).

**Im Terminal:** Gegenprobe braucht Python 3.10 oder neuer, empfohlen ist 3.13.

```bash
brew install python@3.13
git clone https://github.com/WenjaDebug/Crosscheck.git Gegenprobe
cd Gegenprobe
python3.13 -m venv .venv
source .venv/bin/activate
pip install -e .
```

Unter Linux statt `brew`: `sudo apt install python3-venv`.

## Installation (Windows)

Windows 10 oder 11, wie im README [in 5 Schritten](../README.de.md#in-5-schritten): ZIP laden, mit **Alle extrahieren…** auspacken und im Explorer **ExcelAddin.bat** (Seitenfenster in Excel) oder **VoiceAgent.bat** (im Browser) doppelklicken. Beim ersten Start sucht es ein Python 3.10 oder neuer, das schon da ist. Ein `python.exe` unter `WindowsApps` zählt nicht, das ist nur der Verweis auf den Microsoft Store. Fehlt eins, installiert es Python 3.13 über winget, nur für dich und ohne Adminrechte. Fehlt winget (ältere Windows-10-Versionen), öffnet es python.org und sagt, was zu tun ist. Dann richtet es die Pakete ein; danach startet es nur noch. Meldet Windows „Der Computer wurde durch Windows geschützt“, auf **Weitere Informationen → Trotzdem ausführen** klicken. Der Ordner darf Leerzeichen und Umlaute im Namen haben (`C:\Users\Jürg Müller\Crosscheck-main`). Liegt er noch in der ZIP, sagen die Startdateien, dass du zuerst auspacken musst.

### Updates

Wer Gegenprobe mit `git clone` geholt hat, bekommt Updates von selbst: **VoiceAgent.command** und **VoiceAgent.bat** holen vor jedem Start die neueste Version und installieren neue Pakete nach. Pakete, die Gegenprobe nicht mehr braucht (etwa die frühere Spracherkennung), entfernt das Update dabei aus dem Ordner `.venv`. Ohne Internet startet einfach die bisherige Version. Hast du selbst etwas im Ordner geändert oder testest du gerade einen anderen Zweig, lässt Gegenprobe alles, wie es ist, und sagt das im Terminal. Ausschalten lässt sich das mit `VOICEAGENT_NO_UPDATE=1`.

Wer den Ordner als ZIP geladen hat, bekommt nie Code aus dem Netz. Höchstens einmal am Tag fragt Gegenprobe beim Start GitHub ohne Konto, wann die neueste Version entstanden ist, und sagt, wenn es eine neuere gibt (im Fenster und unter **Zustand**). Woher die ZIP stammt, schreibt GitHub beim Packen in `src/voiceagent/stand.txt`. Eine neue Version lädst du wieder als ZIP; wie du sie an die Stelle der alten legst, steht in der [Hilfe](HILFE.md#neue-version). Offline oder mit `VOICEAGENT_NO_UPDATE=1` fragt es nicht.

## Starten

Im Finder **VoiceAgent.command** doppelklicken (Windows: **VoiceAgent.bat**), oder im Terminal:

```bash
source .venv/bin/activate
voiceagent
```

Der Browser öffnet `http://127.0.0.1:8765`. Die Startseite sagt, ob das Fenster in Excel eingerichtet ist, nimmt eine Excel-Datei zum Prüfen an (hineinziehen oder wählen, mit Prüfbericht) und hat die Übungsmappe. Darunter stehen unter **Mehr** der Excel-Helfer und der Zustand.

### Excel-Prüfer (ohne KI)

Auf der Startseite **Excel-Prüfer im Browser** öffnen und in Excel wie gewohnt arbeiten. Gegenprobe liest die offene Mappe direkt mit (am Mac fragt macOS einmal, ob Terminal Excel steuern darf: Erlauben) und zeigt alle paar Sekunden, was nicht aufgeht: Fehlerwerte wie #DIV/0!, eine SUMME, die eine Zeile der Tabelle auslässt oder anders rechnet, gerundete Werte, die zusammen nicht mehr stimmen, eine feste Zahl statt einer Formel, abgetippte Zahlen statt Zellbezügen und dieselbe feste Zahl in vielen Formeln. Dazu prüft er die ganze Rechenkette mit einem eigenen Formelrechner: Formeln, die aus dem Muster ihrer Spalte fallen (mit der Formel, die dort stehen müsste), Kreisbezüge, doppelt gezählte Zwischensummen, Bezüge auf leere Zellen, Werte, die Excel anders zeigt, als die Formel ergibt (Berechnung auf manuell?), Ausreißer, doppelte Namen und Anteile, die nicht 100 % ergeben. Unter „Wovon deine Ergebnisse abhängen“ steht für jedes Ergebnis, wie stark es sich ändert, wenn eine Eingabe oder eine feste Zahl in den Formeln um 1 % steigt. Zu jedem Hinweis steht, was man meistens tun kann. Er liest auch den Namen einer Kennzahl („Umsatz alle Nord“, „Durchschnitt Umsatz Nord“, „Umsatz März 2026“) und rechnet nach, welche Zeilen dazugehören: Fehlt eine Zeile im Bereich, ist eine zu viel, teilt eine Formel durch eine feste Anzahl, die nicht mehr stimmt, oder findet SUMMEWENN wegen eines Leerzeichens oder Tippfehlers zu wenig, sagt er es genau (welche Zeile fehlt, welche zu viel ist und warum) und nennt die Stelle in der Formel, die das verursacht: eine Bedingung zu viel, ein falscher Wert, ein Bereich, der zu früh aufhört, ein falscher Teiler oder eine doppelt gezählte Zelle, auch zwei solche Fehler in einer Formel. In einer Kreuztabelle (Regionen untereinander, Produkte nebeneinander) liest er den Namen aus Zeile und Spalte und findet eine Bedingung aus der falschen Kopfzelle. Zahlen und Daten, die als Text dastehen (oft nach einem Import), findet er mit den Kennzahlen, die dadurch falsch rechnen. Dazu kommt die korrigierte Formel, die er vorher selbst nachgerechnet hat. Er findet auch Bereiche, die beim Herunterziehen ohne $ mitrutschen, einen Mittelwert über Quoten je Zeile statt der Quote über alle Zeilen, Mittelwerte von verschieden großen Gruppen und Wachstum, das durch den neuen statt den alten Wert teilt. Das braucht keine KI, keinen Schlüssel und kein Internet; die Daten bleiben auf dem Rechner.

### Der Prüfer als Seitenfenster in Excel (Mac)

Statt im Browser kann der Prüfer rechts in Excel stehen, wie das Fenster für Kommentare. Einmal einrichten: im Finder **ExcelAddin.command** doppelklicken (oder im Terminal `./ExcelAddin.command`). Fehlt die Installation noch, legt es sie zuerst an wie VoiceAgent.command. Es legt ein Zertifikat nur für diesen Mac an, denn Excel öffnet Add-ins nur über eine gesicherte Verbindung. Dafür fragt der Mac einmal nach deinem Passwort. Danach meldet es das Add-in bei Excel an. Dafür muss Terminal einmal in den Ordner von Excel schreiben dürfen; lässt macOS das nicht zu, öffnet es dir die Einstellung **Festplattenvollzugriff**: dort Terminal einschalten, Terminal ganz beenden (Befehl-Q) und ExcelAddin.command noch einmal starten. Danach kannst du den Festplattenvollzugriff wieder ausschalten; später lässt ExcelAddin.command den Ordner von Excel in Ruhe, solange sich am Add-in nichts ändert.

Zum Schluss richtet es den Prüfer im Hintergrund ein (macOS meldet dafür vielleicht ein neues Hintergrundobjekt „python“): macOS hält den Port 8766 selbst offen und startet ihn, sobald Excel das Seitenfenster öffnet, auch nach einem Neustart des Macs. Bis dahin braucht er keinen Arbeitsspeicher, und er braucht nie ein Terminal oder einen Browser. Ob das geht, probiert ExcelAddin.command gleich aus: Es lädt ihn so, wie macOS es bei der Anmeldung tut, und klopft dann an wie Excel. Startet er nicht von selbst, steht in data/pruefer-dienst.log, was macOS über ihn sagt. Zeigt Excel nach einem Neustart „Add-In-Fehler“, ExcelAddin.command noch einmal doppelklicken: Es sagt, warum er nicht lief (nicht gestartet, abgestürzt, Ordner umbenannt, in macOS abgeschaltet), startet ihn wieder und öffnet, wenn nötig, die Anmeldeobjekte, wo macOS ihn im Hintergrund erlauben muss. Die Startseite sagt es auch. Ab dann genügt: Excel öffnen (beim ersten Mal ganz beenden und neu öffnen), unter **Start** ganz rechts auf **Add-ins** klicken und **Gegenprobe Prüfer** wählen; danach gibt es unter **Start** den Knopf **Prüfer**, ein Klick und er steht rechts in Excel. Ohne den Hintergrund liefert VoiceAgent.command das Seitenfenster, solange es läuft. Nach einem Update durch VoiceAgent.command startet der Prüfer im Hintergrund mit der neuen Version neu, und ein offenes Seitenfenster lädt sich dann selbst neu. Das Seitenfenster liest die Mappe direkt über Excel, prüft nach jeder Änderung neu und markiert eine Zelle, wenn du auf ihren Hinweis klickst. Hat ein Hinweis eine nachgerechnete Korrektur, schreibt **Einsetzen** sie in die Zellen (es zeigt vorher, welche Formel wohin) und **Rückgängig** wieder heraus. Das geht auch für Zahlen als Text (sie werden echte Zahlen im Format ihrer Spalte), zu kurze Bereiche und SVERWEIS ohne FALSCH. Haben mehrere Hinweise eine Korrektur, setzt **Alle einsetzen** sie zusammen ein, nachdem Gegenprobe das Blatt so nachgerechnet hat, wie es danach rechnet; was sich gegenseitig stört, bleibt draußen. Sonst ändert es nichts. Die Zellen gehen nur an Gegenprobe auf diesem Rechner; aus dem Internet lädt Excel nur office.js von Microsoft. Entfernen (auch aus dem Hintergrund): `./ExcelAddin.command --entfernen`.

Das Seitenfenster rechnet die ganze Mappe nach, auch Bezüge auf andere Blätter (Tarife!B3), SVERWEIS, WVERWEIS, XVERWEIS, INDEX und VERGLEICH, Text- und Datumsfunktionen. Typische Fehler beim Nachschlagen meldet es eigens, mit der Ursache und einer nachgerechneten Korrektur zum Einfügen: die falsche Spaltennummer (der Spaltenkopf passt nicht zum Namen der Zelle), ein Bereich, der zu früh aufhört oder zu schmal ist, eine ungefähre Suche (ohne FALSCH am Ende) in einer unsortierten Tabelle oder bei einem Text, der fehlt, eine Tabelle ohne $, die beim Herunterkopieren mitrutscht, und Schlüssel, die nur wegen eines Leerzeichens oder einer Zahl als Text nicht gefunden werden. Dazu Schlüssel, die WENNFEHLER oder WENNNV still mit einem Ersatzwert überdecken, und doppelte Schlüssel.

Summen prüft es gegen ihre Teile: Summenzeilen unter einer Übersicht, Summenspalten rechts daneben und die Gesamtsumme in der Ecke. Es sagt, welcher Teil fehlt, doppelt zählt, abgezogen wird oder nicht dazugehört, auch wenn eine Summe als feste Zahl dasteht, und auch dann, wenn der fehlende Teil gerade 0 ist und die Summe darum noch stimmt. Rechnen die Teile mit SUMMEWENN je Gruppe aus der Tabelle („Umsatz Nord“, „Umsatz Ost“), prüft es, ob sie zusammen alle Zeilen erfassen: eine Gruppe ohne eigene Zeile, ein vertippter Wert („Nrod“) oder eine leere Zelle. Jede Korrektur ist auf dem Blatt nachgerechnet. Bei Statistik achtet es auf Werte, die keine sind: 999 oder -99 für „fehlt“ in MITTELWERT, MEDIAN oder STABW, Nullen aus WENNFEHLER(…;0), leere Zellen oder Text, die eine Summe geteilt durch ANZAHL2 als 0 mitzählt. Es sagt, wenn STABW.N statt STABW.S rechnet, wo „Stichprobe“ steht, wenn unter „Standardabweichung“ die Varianz steht, wenn dem Standardfehler die Wurzel aus n fehlt und wenn Anteile durch die falsche Summe teilen. Bei Tests und Intervallen prüft es, ob sie richtig aufgesetzt sind: T.TEST gepaart über zwei Gruppen oder ungepaart über dieselben Fälle vorher und nachher, mit gleichen Varianzen, wo sie weit auseinanderliegen, oder einseitig, ohne dass es dasteht; KONFIDENZ mit 0,95 statt 0,05, mit dem Standardfehler statt der Standardabweichung oder der Normalverteilung bei wenigen Werten; ein Konfidenzintervall von Hand ohne Wurzel aus n oder mit dem falschen z. Es nennt den richtigen p-Wert und sagt, ob sich die Signifikanz ändert. Die Differenz zweier Prozentwerte unter „Veränderung %“ sind Prozentpunkte, das sagt es auch. Ein Alter von Hand (Tage durch 365 oder 365,25, JAHR minus JAHR) liegt um den Geburtstag herum ein Jahr daneben; es sagt, in welchen Zeilen, und setzt DATEDIF ein. HEUTE() in einer Tabelle von Fällen mit Eintritt oder Austritt ändert sich jeden Tag, gemeint ist meist das Alter bei Eintritt oder die Dauer bis zum Austritt. Tage „inklusive“ ohne +1, „Nächte“ mit +1 und Arbeitstage mit Wochenenden findet es auch. Pivot-Tabellen vergleicht es mit ihren Daten: eine, die nach einer Änderung nicht aktualisiert wurde (mit den Zahlen, die nicht mehr stimmen), eine feste Quelle, unter der die Daten weitergehen, und „Anzahl von …“, wo wegen einer leeren Zelle die Summe gemeint ist. Bei Diagrammen merkt es, wenn eine Reihe vor dem Ende der Tabelle aufhört, die Summenzeile mitnimmt oder gegen ihre Beschriftungen verschoben ist. Im Seitenfenster geht das mit „Aktualisieren“, „Summe nehmen“ und „Einsetzen“. Unter „Auf den anderen Blättern“ steht, wo es sonst noch Fehler gibt; ein Klick öffnet das Blatt.

Im Reiter **Herleitung** steht für die Zelle, die du in Excel markierst, woraus ihre Zahl entsteht: die Formel (auch in Worten), der Wert aus Excel und ob er nachgerechnet stimmt, darunter Stufe für Stufe die Zellen, aus denen er kommt, auch auf anderen Blättern. Bei SVERWEIS, XVERWEIS und INDEX steht die Zelle dabei, die gefunden wurde. Große Bereiche fasst er zusammen (wie viele Zahlen, ihre Summe) und öffnet nur die Zellen, die nicht stimmen oder auf etwas aufbauen, das nicht stimmt. So findest du von einer falschen Summe den Weg bis zur Ursache. Rechnet die Formel aus der Tabelle, steht in einem Satz dabei, welche Zeilen sie wirklich nimmt, zum Beispiel „Summiert Umsatz (€) aus den 9 Zeilen mit Region „Nord“ außer D11 (D2:D10)“. So siehst du auch ohne Hinweis sofort, ob eine Formel rechnet, was sie soll. Unter „Fließt in“ steht, welche Formeln die Zelle lesen. Ein Klick auf eine Zelle im Baum markiert sie in Excel.

Für Datenauswertungen (Fälle, Patientinnen, Messwerte) prüft er zusätzlich: fehlende Werte je Spalte und Platzhalter wie „k.A.“, Codes wie 999 oder -99 für „unbekannt“, mit denen MITTELWERT still rechnet, unmögliche Werte (Alter über 120, negative Liegedauer, Anteile über 100 %, Puls, BMI, Temperatur), ein Austritt vor dem Eintritt oder ein Geburtsdatum in der Zukunft, doppelte Fallnummern und genau gleiche Zeilen, zwei Einheiten in einer Spalte (1,72 neben 172, mmol/l neben mg/dl), MITTELWERT, MEDIAN oder ANZAHL über einen Bereich, der die letzten Zeilen auslässt, einen Nenner, der beim Kopieren mitrutscht (C10 statt C$10), und SUMME oder MITTELWERT über eine gefilterte Tabelle, die die ausgeblendeten Zeilen mitzählt (dann TEILERGEBNIS). Zum Ausprobieren gibt es die Übungsmappe [pruefer-uebung.xlsx](../src/voiceagent/web/demo/pruefer-uebung.xlsx) mit erfundenen Daten und eingebauten Fehlern (auch unten im Seitenfenster).

Der Reiter **Mappe** erklärt die ganze Mappe auf einer Seite, von den Daten bis zum Ergebnis: wo die Daten stehen, welches Blatt aus welchem liest, ob jede Formel nachgerechnet dasselbe ergibt wie in Excel, und wo es Hinweise gibt. Darunter stehen die Ergebnisse (Formeln und gerechnete Spalten, die nichts mehr weiterverwendet), je mit dem, woraus sie über alle Stufen entstehen, auch auf anderen Blättern, mit festen Zahlen und HEUTE() unterwegs und einer Zahl unterwegs, die nicht stimmt. Je Blatt steht jede gerechnete Spalte mit ihrer Formel in Worten („=Betrag * Satz“, in allen 7 Zeilen gleich) und jede einzelne Formel mit dem Satz, welche Zeilen sie nimmt. Ein Bezug ohne $, der beim Herunterziehen mitwandert, steht dabei („E3 nimmt D10“). Ein Klick markiert die Zelle in Excel. So verstehst du auch eine Mappe, die jemand anderes gebaut hat. Der Prüfbericht mit Werten hat denselben Abschnitt.

Unter **Eigene Regeln je Spalte** legst du fest, was in einer Spalte stehen darf: Markiere eine Zelle der Spalte, klicke **Regel für die markierte Spalte** und wähle Pflichtfeld, Art (Zahl, ganze Zahl, Datum, Text), von und bis, nicht vor oder nach einem Tag, nicht in der Zukunft, erlaubte Werte („m, w, d“), jeder Wert nur einmal, ein Muster („PAT-####“: # eine Ziffer, @ ein Buchstabe) oder einen Vergleich mit einer anderen Spalte (Austritt nicht vor Eintritt). Das Formular zeigt dabei, was jetzt in der Spalte steht. Die Regel gilt in jeder Mappe mit einer Spalte dieses Namens (oder nur auf einem Blatt) und bleibt auf dem Rechner. Jeder Verstoß kommt als Fehler mit Zelle und Grund, auch was nur anders geschrieben ist („M“ statt „m“, ein Leerzeichen am Ende), eine Zahl als Text oder ein Datum ohne Datumsformat. Ein Klick markiert alle Zellen, **Regel ändern** öffnet die Regel.

**Große Mappen** bleiben bedienbar: Das Seitenfenster liest große Blätter in Stücken und zeigt, wie weit es ist. Ein Blatt, das sich nicht geändert hat, schickt es nicht noch einmal, und andere große Blätter prüft Gegenprobe im Hintergrund, solange du nichts änderst („wird noch geprüft“). Gelesen werden bis 500 000 Zellen je Blatt und 1,5 Millionen in der ganzen Mappe, das aktive Blatt zuerst. Ist ein Blatt größer, liest er es bis dorthin in ganzen Zeilen und sagt, bis zu welcher Zeile. Zu Formeln, die mit den Zeilen darunter rechnen, sagt er dann nichts, statt einen falschen Unterschied zu Excel zu melden.

Unter **Prüfbericht** im Seitenfenster öffnet sich ein Bericht über alle Blätter zum Drucken oder Sichern als PDF (in Chrome oder Safari Drucken → Als PDF sichern). **Mit Werten** stehen die ganzen Hinweise drin, mit den Zahlen und Namen aus der Mappe. **Ohne Daten** stehen nur Stufe, Zelle und die Art des Hinweises da („Code für ‚fehlt‘ wird mitgerechnet“), die Blätter heißen Blatt 1, Blatt 2, und der Name der Mappe fehlt. Den kannst du weitergeben, auch wenn in der Mappe Patientendaten stehen. Die Berichte liegen im Datenordner unter `berichte`, nur die letzten 20 bleiben.

Im Reiter **Datei** ziehst du eine frühere Version der Mappe ins Seitenfenster (oder klickst und wählst die .xlsx-Datei). Gegenprobe vergleicht sie mit der offenen Mappe und zeigt zuerst die Ergebnisse, die jetzt anders sind, mit dem Grund: welche Eingabe sich davor geändert hat, welche Formel anders ist, ob neue Zeilen mitzählen oder Zeilen fehlen, was sich auf einem anderen Blatt geändert hat, oder dass eine Zahl in Excel nicht neu gerechnet ist. Darunter Formeln, die jetzt feste Zahlen sind, geänderte Formeln (je Spalte zusammengefasst), geänderte Eingaben, neue und gelöschte Zeilen und Spalten, umbenannte Blätter. Eingefügte und umsortierte Zeilen erkennt er an Zeilennamen und Inhalt, so dass verschobene Formeln nicht als geändert gelten. Mit **Prüfen** prüft er die Datei selbst, ohne dass du sie in Excel öffnest, alle Blätter, auch mit Prüfbericht (ohne Daten oder mit Werten). Ist die Datei keine Version der offenen Mappe (kein Blatt gleich), prüft er sie gleich so. Die Datei bleibt im Arbeitsspeicher.

### Der Prüfer als Seitenfenster in Excel (Windows)

Unter Windows geht dasselbe mit einem Doppelklick auf **ExcelAddin.bat** im Explorer, ohne Adminrechte und ohne weitere Pakete. Fehlt die Installation noch, legt es sie zuerst an wie VoiceAgent.bat. Dann: (1) ein Zertifikat nur für diesen PC, mit Python selbst, ohne openssl; (2) Windows fragt einmal in einer Sicherheitswarnung, ob es der Zertifizierungsstelle „VoiceAgent lokal (nur dieser Rechner)“ vertrauen soll: **Ja** klicken (sie kommt nur zu den Zertifikaten deines Benutzers); (3) es meldet das Add-in als Entwickler-Add-in bei Excel an (Registry `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`, das Manifest liegt in `%USERPROFILE%\.voiceagent\addin`); (4) der Prüfer startet bei jeder Anmeldung ohne Fenster im Hintergrund (`HKCU\…\CurrentVersion\Run`, pythonw) und gleich jetzt einmal, und ExcelAddin.bat probiert aus, ob er antwortet. Danach Excel ganz schließen und neu öffnen: Unter **Start** steht der Knopf **Prüfer**. Fehlt er, unter **Start → Add-ins** (oder **Einfügen → Meine Add-ins**) bei den Entwickler-Add-ins **Gegenprobe Prüfer** wählen. Im Trust Center musst du nichts einstellen. Läuft VoiceAgent.bat, liefert es das Seitenfenster selbst; nach einem Update durch VoiceAgent.bat startet der Prüfer im Hintergrund mit der neuen Version neu. Was er meldet, steht in `data\pruefer-dienst.log`. Entfernen: in der Eingabeaufforderung im Ordner `ExcelAddin.bat --entfernen`. Das Zertifikat bleibt dabei in Windows, damit ein neues Einrichten nicht wieder fragt; herausnehmen lässt es sich unter **Benutzerzertifikate verwalten**.

Laut Microsoft zeigen Excel 2016 und 2019 Add-ins noch mit dem Internet Explorer; damit läuft das Seitenfenster nicht, dort bleibt der Prüfer im Browser. Excel 2021, 2024 und Microsoft 365 nehmen WebView2 (Edge). Ältere Microsoft-365-Versionen mit dem alten Edge (EdgeHTML) brauchen für localhost eine Ausnahme (`CheckNetIsolation LoopbackExempt`). Auf einem echten Windows-PC ist das Seitenfenster noch nicht erprobt: Registry, certutil und Prozessstart sind in den Tests nachgebaut, das Zertifikat ist mit einer echten, strengen TLS-Verbindung geprüft.

### Excel-Helfer: Pivot-Übersicht auf Zuruf

Auf der Startseite unter **Mehr** oder im Excel-Prüfer **Excel-Helfer** öffnen. Es geht ein schmales Fenster auf, das neben Excel stehen bleibt. Dort schreibst du , was du willst, zum Beispiel „Summe der Sitze pro Partei“, „Durchschnitt Alter pro Station und Geschlecht“, „Anzahl pro Kanton“ oder „Summe Umsatz pro Region, Jahre in den Spalten“. Gegenprobe nimmt das aktive Blatt, findet die Spalten (auch ungefähr: „Parteien“ für „Partei“) und zeigt zuerst eine Vorschau mit den Zahlen, die herauskommen. Erst nach **In Excel schreiben** legt er ein neues Blatt an; bestehende Zellen ändert er nie.

Die Übersicht besteht aus echten Formeln (SUMMEWENNS, ZÄHLENWENNS, MITTELWERTWENNS, MINWENNS, MAXWENNS) auf deine Daten, mit Gesamtzeile, auf Wunsch als Kreuztabelle und sortiert („größte zuerst“). Sie rechnet also mit, wenn sich die Daten ändern, und jede Zahl lässt sich in Excel nachvollziehen. Eine Gesamtsumme unter den Daten lässt er weg. Nach dem Schreiben liest er das neue Blatt zurück und vergleicht jede Zahl mit seiner eigenen Rechnung. Eine echte Pivot-Tabelle von Excel kann er zusätzlich versuchen (unter Windows über COM, am Mac über AppleScript); klappt es nicht, steht die Übersicht mit Formeln da.

Den Satz versteht Gegenprobe ohne KI, auf dem Rechner. Rückgängig (Befehl-Z) geht für das neue Blatt nicht: Gefällt es nicht, Rechtsklick auf den Blattnamen, Löschen.

### Excel-Helfer: Tabellen aus Dateien einfügen

Im Excel-Helfer unter „Oder eine Tabelle aus einer Datei“ eine Datei wählen oder ins Fenster ziehen: CSV oder Text (mit Semikolon, Komma, Tab oder |), Excel (.xlsx, .xlsm, jedes Blatt), Word (.docx, jede Tabelle) oder PDF (Tabellen mit Text, keine Scans). Hat die Datei mehrere Tabellen, wählst du eine aus. Gegenprobe zeigt die Tabelle zuerst und schreibt sie erst nach **In Excel schreiben** in ein neues Blatt der offenen Mappe, mit fetter Kopfzeile.

Zahlen werden echte Zahlen, auch deutsch (1.234,50), schweizerisch (1'234.50) oder englisch (1,234.50), mit Prozent (12,5 %) und Währung (CHF 80.–). Datumsangaben (1.3.2024, 2024-03-01) werden echte Daten. Was nur wie eine Zahl aussieht, bleibt Text, damit nichts verloren geht: Postleitzahlen und Nummern mit Null vorne (08001), Telefonnummern, IBAN und lange Kennnummern. Nach dem Schreiben liest er das Blatt zurück und vergleicht jeden Wert mit der Datei. Die Datei bleibt auf dem Rechner. Danach kann man auf dem neuen Blatt gleich eine Übersicht verlangen.

### Zustand und Fehlerbericht

Unter **Zustand** (Startseite, unter Mehr) steht, ob alles bereit ist: Version, Verbindung zu Excel und Datenordner. Fehlt etwas, steht dabei, was zu tun ist.

Geht etwas nicht, lade dort den **Fehlerbericht** herunter und schick ihn mit, wenn du um Hilfe fragst. Darin stehen der Zustand, die Versionen der Pakete und die letzten Warnungen und Fehler. Zellinhalte stehen nie darin, Ordnernamen unter deinem Benutzer werden zu `~`, und E-Mail, IBAN und Ähnliches sind geschwärzt. Auf der Seite siehst du vorher, was genau darin steht.

## Datenschutz

- Alles bleibt auf deinem Rechner. Gegenprobe braucht keine KI, kein Konto und kein Internet; nur Updates holt sie von GitHub. Aus einer ZIP fragt sie höchstens einmal am Tag GitHub, ob es eine neuere Version gibt, ohne etwas aus dem Ordner oder den Mappen mitzuschicken.
- Der Server hört nur auf `127.0.0.1` und ist aus dem Netz nicht erreichbar. Auch eine Webseite, die du im Browser offen hast, kommt nicht an deine Daten: Gegenprobe antwortet nur Seiten auf diesem Rechner (`127.0.0.1` und `localhost`).
- Das Seitenfenster in Excel schickt die Zellen nur an Gegenprobe auf diesem Rechner; aus dem Internet lädt Excel nur office.js von Microsoft.
- Eine Datei, die du prüfst oder vergleichst, bleibt im Arbeitsspeicher und wird nicht gespeichert.
- Der Prüfbericht **ohne Daten** enthält nur Stufe, Zelle und Art des Hinweises. Den kannst du weitergeben, auch wenn in der Mappe Patientendaten stehen.
- Gegenprobe speichert keine Inhalte deiner Mappen. Das Protokoll (`data/pruefer-dienst.log`) enthält keine Zellen, Namen oder Werte, auch nicht im Fehlerbericht.
- Nur drei Dinge bleiben, alle von dir ausgelöst: **Passt so** steht in der Mappe selbst, als Prüfsumme ohne Inhalt. **Eigene Spaltenregeln** liegen in `data/spaltenregeln.json`. Ein **Prüfbericht** liegt in `data/berichte`, damit der Browser ihn öffnen kann, und wird nach einem Tag gelöscht. Wer ihn behalten will, sichert ihn im Browser als PDF.

## Grenzen und Haftung

Gegenprobe rechnet nach und sucht typische Fehler. Es kann Fehler übersehen und auch etwas melden, das richtig ist. Es ersetzt die eigene Kontrolle nicht: Für die Zahlen und Entscheidungen bleibt verantwortlich, wer die Mappe nutzt. Jede Korrektur zeigt vor dem Einsetzen, was sie wohin schreibt, und **Rückgängig** nimmt sie wieder heraus. Gegenprobe ist kostenlos und quelloffen und wird ohne jede Gewähr bereitgestellt (siehe [Lizenz](../LICENSE)). Das Seitenfenster zeigt diesen Hinweis beim ersten Öffnen und danach jederzeit über **i**.

## Sparsam mit Speicher

Gegenprobe braucht kein Modell und keine großen Pakete: Python mit FastAPI, Uvicorn und pypdf, zusammen etwa 60 MB auf der Festplatte. Der Prüfer im Hintergrund am Mac startet erst, wenn Excel das Seitenfenster öffnet.

Wer eine ältere Version mit Spracherkennung benutzt hat, kann deren Modell löschen (etwa 1,6 GB). Auch die Ordner `data/sessions` (Sitzungen mit Bildschirmfotos) und `data/anweisungen` stammen vom früheren Lehrling; Gegenprobe liest sie nicht mehr, du kannst sie löschen.

```bash
rm -rf ~/.cache/huggingface/hub/models--mobiuslabsgmbh--faster-whisper-large-v3-turbo
```

## Einstellungen

| Umgebungsvariable | Standard | Bedeutung |
|---|---|---|
| `VOICEAGENT_DATA_DIR` | `data` | Wo Gegenprobe Prüfberichte, eigene Regeln und das Protokoll ablegt |
| `VOICEAGENT_NO_UPDATE` | leer | `1`: vor dem Start nicht nach einer neuen Version schauen |
| `VOICEAGENT_PYTHON` | leer | Dieses Python für die Installation nehmen statt selbst eins zu suchen |
| `VOICEAGENT_NO_EXCEL` | leer | `1`: nicht mit dem offenen Excel verbinden (zum Beispiel für Tests) |

## Entwickeln

```bash
pip install -e ".[dev]"
pytest
```

Der Code liegt in `src/voiceagent`. Die Prüfungen stehen in `analyse.py` (Einstieg), `formel.py` (Formelrechner), `workbook.py` (Tabellen erkennen), `kopf.py` (Köpfe gegen Formeln), `finanz.py` und `fallen.py` (Finanz- und Statistikfallen); `pruefer.py` sammelt die Hinweise, `app.py` liefert Seiten und Schnittstellen aus, `web/` enthält Startseite, Prüfer und Seitenfenster.

## Lizenz

MIT, siehe [LICENSE](../LICENSE).
