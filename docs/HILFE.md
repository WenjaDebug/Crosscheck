# Wenn etwas nicht geht

Hier steht für jedes bekannte Problem, woran du es erkennst und was du tust. Gegenprobe zeigt im Fenster (Terminal am Mac, das schwarze Fenster unter Windows) fast immer einen Satz dazu, was fehlt. Hilft nichts davon: Mach ein Bildschirmfoto von diesem Fenster und frag damit nach. Läuft Gegenprobe schon, gibt es auf der Startseite unter **Mehr → Zustand** einen **Fehlerbericht** zum Mitschicken; er enthält nichts aus deinen Mappen.

*English: [HELP.md](HELP.md)*

**Inhalt:** [Python fehlt oder ist zu alt](#python-fehlt-oder-ist-zu-alt) · [Kein Internet beim ersten Start](#kein-internet-beim-ersten-start) · [Mac öffnet die Datei nicht](#mac-öffnet-die-datei-nicht) · [Windows warnt vor der Datei](#windows-warnt-vor-der-datei) · [Windows: „noch in der ZIP-Datei“](#windows-noch-in-der-zip-datei) · [Mac fragt nach dem Passwort](#mac-fragt-nach-dem-passwort) · [Mac: Terminal darf nicht in den Ordner von Excel schreiben](#mac-terminal-darf-nicht-in-den-ordner-von-excel-schreiben) · [Mac fragt nach Downloads, Schreibtisch oder Dokumente](#mac-fragt-nach-downloads-schreibtisch-oder-dokumente) · [Windows-Sicherheitswarnung zum Zertifikat](#windows-sicherheitswarnung-zum-zertifikat) · [Der Knopf Prüfer fehlt in Excel](#der-knopf-prüfer-fehlt-in-excel) · [Seitenfenster leer, Add-In-Fehler oder alte Version](#seitenfenster-leer-add-in-fehler-oder-alte-version) · [Excel 2016 oder 2019](#excel-2016-oder-2019) · [Port belegt](#port-belegt) · [Virenscanner oder gesperrte Skripte](#virenscanner-oder-gesperrte-skripte) · [Ordner verschoben oder umbenannt](#ordner-verschoben-oder-umbenannt) · [Neue Version](#neue-version) · [Add-in wieder entfernen](#add-in-wieder-entfernen) · [Alles deinstallieren](#alles-deinstallieren)

## Python fehlt oder ist zu alt

**Woran du es erkennst:** Im Fenster steht „Gegenprobe braucht Python (ein kostenloses Programm), und das fehlt auf diesem Mac/PC noch.“, und der Browser öffnet python.org.

Gegenprobe braucht Python 3.10 oder neuer. Es nimmt eins, das schon da ist: am Mac von python.org (`/Library/Frameworks/Python.framework`), von Homebrew oder sonst auf dem PATH; unter Windows aus den üblichen Ordnern von python.org, über den Starter `py` oder auf dem PATH. Das Python, das zu macOS gehört (`/usr/bin/python3`), ist meist 3.9 und damit zu alt. Gegenprobe fragt es nur, wenn die Befehlszeilentools installiert sind, damit macOS nicht fragt, ob es sie installieren soll. Unter Windows zählt ein `python.exe` im Ordner `WindowsApps` nicht: Das ist nur ein Verweis auf den Microsoft Store.

**Was du tust:**

- **Mac:** Auf der Seite, die sich geöffnet hat (python.org/downloads/macos), bei „Python 3.13…“ auf **macOS 64-bit universal2 installer** klicken. Die geladene Datei doppelklicken und mit **Fortfahren** und **Installieren** durchklicken; der Mac fragt nach deinem Passwort. Das braucht ein Konto mit Adminrechten; auf einem Mac der Firma oder des Spitals macht das sonst die IT. Danach **ExcelAddin.command** (oder **VoiceAgent.command**) noch einmal doppelklicken. Ist Homebrew installiert, installiert Gegenprobe Python 3.13 selbst damit.
- **Windows:** Zuerst versucht Gegenprobe, Python 3.13 mit winget zu installieren (nur für dich, ohne Adminrechte). Geht das nicht (ältere Windows-10-Versionen haben kein winget), öffnet es python.org/downloads/windows: Bei „Python 3.13…“ auf **Windows installer (64-bit)** klicken, die Datei doppelklicken. Hast du keine Adminrechte, unten das Häkchen bei **Use admin privileges when installing py.exe** wegnehmen. Dann **Install Now**. Danach **ExcelAddin.bat** (oder **VoiceAgent.bat**) noch einmal doppelklicken.

Ein bestimmtes Python lässt sich mit der Umgebungsvariable `VOICEAGENT_PYTHON` vorgeben.

## Kein Internet beim ersten Start

**Woran du es erkennst:** „Ich komme nicht ins Internet. Beim ersten Start lädt Gegenprobe einmal ihre drei Pakete (FastAPI, Uvicorn, pypdf).“

**Was du tust:** Den Rechner mit dem Internet verbinden und die Datei noch einmal doppelklicken. Nur die Installation braucht Internet (Python und die drei Pakete von pypi.org); danach arbeitet Gegenprobe ganz ohne. Lässt ein Firmen- oder Spitalnetz pypi.org nicht zu, muss die IT das einmal freigeben oder die Installation an einem anderen Netz machen.

## Mac öffnet die Datei nicht

**Woran du es erkennst:** Nach dem Doppelklick auf **ExcelAddin.command** oder **VoiceAgent.command** sagt macOS, die Datei wurde nicht geöffnet, weil Apple nicht prüfen konnte, ob sie frei von Schadsoftware ist, oder sie stamme von einem nicht verifizierten Entwickler. Das passiert bei jeder Datei aus dem Internet, die nicht bei Apple angemeldet ist, einmal je Datei.

**Was du tust:**

- **macOS 15 (Sequoia) und neuer:** Im Hinweis auf **Fertig** klicken. Dann **Systemeinstellungen → Datenschutz & Sicherheit** öffnen, ganz nach unten scrollen und bei der Datei auf **Dennoch öffnen** klicken, mit deinem Passwort bestätigen und noch einmal **Öffnen**.
- **macOS 14 und älter:** Mit der rechten Maustaste (oder Control-Klick) auf die Datei klicken, **Öffnen** wählen und im Hinweis noch einmal **Öffnen**.

Danach geht der Doppelklick normal. Gegenprobe schaltet diese Prüfung nie selbst ab.

**„… konnte nicht ausgeführt werden, da Sie nicht über die nötigen Zugriffsrechte verfügen“:** Das Entpacken hat das Recht zum Ausführen verloren. Die ZIP von GitHub hat es (mit dem Archivierungsprogramm von macOS, also einem Doppelklick auf die ZIP, bleibt es erhalten); andere Entpack-Programme lassen es manchmal weg. Am einfachsten die ZIP noch einmal mit einem Doppelklick entpacken. Sonst: Terminal öffnen (Programme → Dienstprogramme → Terminal), `chmod +x ` tippen (mit Leerzeichen), die Datei aus dem Finder ins Terminal-Fenster ziehen und Enter drücken.

## Windows warnt vor der Datei

**Woran du es erkennst:** Nach dem Doppelklick auf **ExcelAddin.bat** oder **VoiceAgent.bat** erscheint „Der Computer wurde durch Windows geschützt“ oder „Datei öffnen – Sicherheitswarnung: Möchten Sie diese Datei ausführen?“. Windows warnt so bei jeder Datei aus dem Internet ohne Unterschrift.

**Was du tust:** Bei „Der Computer wurde durch Windows geschützt“ auf **Weitere Informationen** und dann auf **Trotzdem ausführen** klicken. Bei der Sicherheitswarnung auf **Ausführen**. Wer die Warnung gar nicht sehen will: vor dem Entpacken Rechtsklick auf die ZIP-Datei → **Eigenschaften** → unten **Zulassen** ankreuzen → **OK**, dann entpacken.

## Windows: „noch in der ZIP-Datei“

**Woran du es erkennst:** „Gegenprobe steckt noch in der ZIP-Datei. Bitte zuerst entpacken“. Ein Doppelklick auf eine ZIP zeigt unter Windows nur ihren Inhalt, ohne ihn auszupacken.

**Was du tust:** Rechtsklick auf die ZIP-Datei → **Alle extrahieren…** → einen Ort wählen (am besten deinen Benutzerordner) → **Extrahieren**. Dann im ausgepackten Ordner doppelklicken.

## Mac fragt nach dem Passwort

**Woran du es erkennst:** ExcelAddin.command schreibt „2/4 Jetzt fragt der Mac nach deinem Passwort, damit er dem Zertifikat vertraut“, und ein Fenster von macOS will dein Passwort.

**Was du tust:** Dein Mac-Passwort eingeben (das, mit dem du dich anmeldest). Excel öffnet Add-ins nur über eine gesicherte Verbindung. Dafür legt Gegenprobe eine eigene Zertifizierungsstelle nur für diesen Mac an (im Ordner `.voiceagent` in deinem Benutzerordner; ihr Schlüssel verlässt den Mac nie) und trägt sie in deinen Schlüsselbund ein. Gegenprobe stellt damit nur ein Zertifikat für `localhost` aus. Das ist einmalig. Hast du abgebrochen, ExcelAddin.command noch einmal doppelklicken.

## Mac: Terminal darf nicht in den Ordner von Excel schreiben

**Woran du es erkennst:** „macOS lässt das Terminal nicht in den Ordner von Excel schreiben“, und die Einstellung **Festplattenvollzugriff** geht auf.

**Was du tust:** Wie im Fenster beschrieben: dort Terminal einschalten (fehlt es: unten auf **+**, dann Programme → Dienstprogramme → Terminal), Terminal ganz beenden (Befehl-Q) und ExcelAddin.command noch einmal doppelklicken. Danach kannst du den Festplattenvollzugriff wieder ausschalten.

## Mac fragt nach Downloads, Schreibtisch oder Dokumente

**Woran du es erkennst:** „Terminal möchte auf Dateien in deinem Ordner „Downloads“ zugreifen“, oder beim ersten Einrichten der Tipp „Dieser Ordner liegt in Downloads, auf dem Schreibtisch, in Dokumente oder in iCloud“.

**Was du tust:** Am besten den Ordner von Gegenprobe vor dem ersten Start in deinen Benutzerordner legen (im Finder: **Gehe zu → Benutzerordner**). Diese Ordner schützt macOS, und der Prüfer im Hintergrund kann dort hängen bleiben. Ist schon alles eingerichtet, siehe [Ordner verschoben oder umbenannt](#ordner-verschoben-oder-umbenannt). Sonst die Frage mit **OK** beantworten.

## Windows-Sicherheitswarnung zum Zertifikat

**Woran du es erkennst:** ExcelAddin.bat schreibt „2/4 Jetzt fragt Windows einmal, ob es dem Zertifikat „VoiceAgent lokal (nur dieser Rechner)“ vertrauen soll“, und Windows zeigt eine **Sicherheitswarnung**, in der dieser Name steht.

**Was du tust:** **Ja** klicken. Die Zertifizierungsstelle ist auf diesem PC entstanden, ihr Schlüssel verlässt ihn nie, und Gegenprobe stellt damit nur ein Zertifikat für `localhost` aus. Sie kommt nur zu den Zertifikaten deines Benutzers, Adminrechte brauchst du nicht. Hast du **Nein** geklickt, ExcelAddin.bat noch einmal doppelklicken.

## Der Knopf Prüfer fehlt in Excel

**Woran du es erkennst:** ExcelAddin.command oder ExcelAddin.bat sagte „Fertig“, aber unter **Start** gibt es keinen Knopf **Prüfer**.

**Was du tust:**

1. Excel ganz beenden, nicht nur das Fenster schließen: am Mac Befehl-Q, unter Windows alle Excel-Fenster schließen. Dann Excel neu öffnen und eine Mappe öffnen.
2. **Mac:** Unter **Start** ganz rechts auf **Add-ins** klicken und **Gegenprobe Prüfer** wählen. Danach steht unter **Start** der Knopf **Prüfer**.
3. **Windows:** Unter **Start** auf **Add-ins** (oder **Einfügen → Meine Add-ins**) klicken, dort bei den Entwickler-Add-ins **Gegenprobe Prüfer** wählen.
4. Hilft das nicht, ExcelAddin noch einmal doppelklicken und lesen, was es sagt.

## Seitenfenster leer, Add-In-Fehler oder alte Version

**Woran du es erkennst:** Das Fenster rechts bleibt weiß, Excel meldet „Add-In-Fehler“, oder es fehlt etwas, das es nach einem Update geben sollte.

**Was du tust:**

1. **ExcelAddin noch einmal doppelklicken.** Es prüft, ob der Prüfer im Hintergrund läuft, sagt, warum nicht (nicht gestartet, abgestürzt, Ordner umbenannt, in macOS abgeschaltet), und startet ihn neu. Am Mac öffnet es, wenn nötig, **Systemeinstellungen → Allgemein → Anmeldeobjekte**: dort unter **Im Hintergrund erlauben** python3.13 (oder Python) einschalten. Was der Prüfer selbst meldet, steht in `data/pruefer-dienst.log` im Ordner von Gegenprobe.
2. **Zwischenspeicher von Excel leeren.** Gegenprobe gibt jeder neuen Fassung des Fensters eine neue Adresse, das genügt meistens. Sonst:
   - **Mac:** Im Kopf des Seitenfensters oben rechts das Menü von Excel öffnen (nicht das **i** von Gegenprobe) und **Clear Web Cache** wählen (je nach Sprache und Version heißt es etwas anders), dann das Fenster schließen und wieder öffnen.
   - **Windows:** Excel ganz schließen und neu öffnen. Hilft das nicht: Excel schließen, im Explorer oben in die Adresszeile `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` tippen, Enter, alles darin löschen, Excel wieder öffnen.

## Excel 2016 oder 2019

**Woran du es erkennst:** Unter Windows mit Excel 2016 oder 2019 bleibt das Seitenfenster leer oder zeigt eine Fehlermeldung.

**Was du tust:** Diese Versionen zeigen Add-ins noch mit dem alten Internet Explorer, damit läuft das Seitenfenster nicht. Nimm den Prüfer im Browser: **VoiceAgent.bat** doppelklicken, auf der Startseite **Excel-Prüfer im Browser**. Er liest die offene Mappe direkt aus Excel. Excel 2021, 2024 und Microsoft 365 können das Seitenfenster.

## Port belegt

**Woran du es erkennst:**

- „Port 8766 ist belegt. Läuft noch VoiceAgent.command?“ (Mac) oder „Läuft noch VoiceAgent.bat?“ (Windows) beim Einrichten mit ExcelAddin. Port 8766 gehört dem Seitenfenster.
- „Port 8765 hat schon ein anderes Programm. Gegenprobe nimmt diesmal Port …“ beim Start von VoiceAgent. Dann nimmt Gegenprobe selbst den nächsten freien Port; der Browser öffnet die richtige Adresse.
- „Gegenprobe läuft schon, in einem anderen Fenster.“ Dann öffnet ein zweiter Doppelklick nur den Browser.

**Was du tust:** Bei Port 8766 das Fenster von VoiceAgent schließen (oder dort Ctrl + C drücken) und ExcelAddin noch einmal doppelklicken. Sonst nichts: Gegenprobe weicht selbst aus.

## Virenscanner oder gesperrte Skripte

**Woran du es erkennst:** Der Virenscanner meldet `ExcelAddin.bat`, `powershell` oder `pythonw.exe`, oder im Fenster steht „… kann nicht geladen werden, da die Ausführung von Skripts auf diesem System deaktiviert ist“.

**Was du tust:** Auf einem Rechner der Firma oder des Spitals die IT fragen. Ihr kannst du sagen, was Gegenprobe tut: Alle Startdateien und der Code sind lesbare Textdateien. Unter Windows startet `VoiceAgent.bat` das Skript `scripts\start-windows.ps1` mit `powershell -ExecutionPolicy Bypass` (nur für diesen einen Aufruf). Es installiert Python mit winget und drei Pakete von pypi.org in den Ordner `.venv`. ExcelAddin.bat legt ein Zertifikat für `localhost` an (`%USERPROFILE%\.voiceagent\addin`), trägt es mit `certutil -user` ein und meldet das Add-in unter `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` an. Unter `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` („Gegenprobe Pruefer“) startet es den Prüfer bei jeder Anmeldung mit `pythonw.exe -m voiceagent.dienst`, ohne Fenster. Er hört nur auf `127.0.0.1`. Am Mac richtet ExcelAddin.command dasselbe mit einem LaunchAgent (`~/Library/LaunchAgents/io.github.wenjadebug.voiceagent.pruefer.plist`) und dem Schlüsselbund ein. Daten aus Mappen gehen nirgends hin.

## Ordner verschoben oder umbenannt

**Woran du es erkennst:** Nach dem Verschieben oder Umbenennen des Ordners startet VoiceAgent nicht mehr (zum Beispiel „bad interpreter“ oder „No such file“), oder das Seitenfenster meldet einen Add-In-Fehler.

**Was du tust:** Die eigene Python-Umgebung hängt am Ort des Ordners. Am einfachsten den Ordner zurückschieben. Sonst im Ordner von Gegenprobe den Ordner `.venv` löschen (am Mac ist er versteckt: im Finder Befehl-Umschalt-Punkt zeigt ihn) und **ExcelAddin** (oder **VoiceAgent**) noch einmal doppelklicken. Das richtet alles am neuen Ort neu ein.

## Neue Version

**Woran du es erkennst:** Beim Start von VoiceAgent steht „Es gibt eine neuere Version von Gegenprobe (vom …)“, auch auf der Startseite unter **Mehr → Zustand**.

Hast du Gegenprobe als ZIP geladen, lädt es nie selbst Code aus dem Internet. Höchstens einmal am Tag fragt es GitHub, wann die neueste Version entstanden ist, ohne Konto und ohne etwas aus dem Ordner oder deinen Mappen mitzuschicken. Ausschalten mit der Umgebungsvariable `VOICEAGENT_NO_UPDATE=1`. Wer Gegenprobe mit `git clone` geholt hat, bekommt neue Versionen beim Start von selbst.

**Was du tust:**

1. Den alten Ordner umbenennen, zum Beispiel in „Gegenprobe alt“.
2. Die neue ZIP wie beim ersten Mal laden und entpacken und den neuen Ordner genau dorthin legen und genau so nennen, wie der alte hieß.
3. Eigene Spaltenregeln mitnehmen: die Datei `data/spaltenregeln.json` aus dem alten Ordner in den Ordner `data` des neuen kopieren (falls es sie gibt).
4. **ExcelAddin** im neuen Ordner doppelklicken. Danach den alten Ordner löschen.

## Add-in wieder entfernen

Das nimmt Gegenprobe aus Excel und aus dem Hintergrund. Es braucht ein Wort extra, `--entfernen`:

- **Mac:** Terminal öffnen (Programme → Dienstprogramme → Terminal). **ExcelAddin.command** aus dem Finder ins Terminal-Fenster ziehen, ein Leerzeichen und `--entfernen` tippen, Enter. Danach Excel neu starten. Das Zertifikat bleibt im Schlüsselbund. Herausnehmen: **Schlüsselbundverwaltung** öffnen, links **Anmeldung**, oben **Zertifikate**, „VoiceAgent lokal (nur dieser Rechner)“ löschen.
- **Windows:** Den Ordner von Gegenprobe im Explorer öffnen, oben in die Adresszeile `cmd` tippen, Enter. Im schwarzen Fenster `ExcelAddin.bat --entfernen` tippen, Enter. Danach Excel ganz schließen und neu öffnen. Das Zertifikat bleibt in Windows, damit ein neues Einrichten nicht noch einmal fragt. Herausnehmen: Windows-Taste, „Benutzerzertifikate verwalten“, **Vertrauenswürdige Stammzertifizierungsstellen → Zertifikate**, „VoiceAgent lokal (nur dieser Rechner)“ löschen.

## Alles deinstallieren

1. Zuerst das [Add-in entfernen](#add-in-wieder-entfernen), solange der Ordner noch da ist.
2. Den Ordner von Gegenprobe löschen (am Mac in den Papierkorb). Darin liegt alles, auch die eigene Python-Umgebung `.venv` und der Ordner `data` mit deinen Spaltenregeln, Prüfberichten und dem Protokoll.
3. Den Ordner `.voiceagent` in deinem Benutzerordner löschen (Zertifikate und Schlüssel). Am Mac ist er versteckt: im Finder Befehl-Umschalt-Punkt zeigt ihn. Unter Windows: im Explorer in die Adresszeile `%USERPROFILE%` tippen.
4. Das Zertifikat aus dem Schlüsselbund oder aus Windows nehmen, wie oben beschrieben.
5. Python selbst kann bleiben, es stört nicht. Wer es loswerden will: unter Windows **Einstellungen → Apps**, „Python 3.13“ deinstallieren; am Mac liegt es unter **Programme → Python 3.13** und in `/Library/Frameworks/Python.framework`.
