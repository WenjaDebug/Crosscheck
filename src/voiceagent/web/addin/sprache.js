// Sprache des Seitenfensters: Deutsch oder Englisch. Am Anfang wie die Sprache des Macs (und damit von Excel),
// danach wie im Fenster unter „i“ gewählt. tr("Eingesetzt in {0}.", "C5") gibt den Satz in der gewählten Sprache;
// die Werte darin (Zellen, Namen, Zahlen) übersetzt es nie. Die festen Texte im HTML übersetzt applyStatic() einmal
// beim Laden, bevor Daten aus der Mappe im Fenster stehen.
// Formeln zeigt das Fenster so, wie Excel sie in seiner Sprache zeigt: =SUMME(A1;B1) im deutschen Excel,
// =SUM(A1,B1) in jedem anderen. Eingesetzt wird ohnehin die Grundform, die Excel in jeder Sprache versteht.

const LANG_KEY = "gegenprobe-sprache";
const LANGS = ["de", "en"];

function chosenLang() {
  try {
    const stored = window.localStorage.getItem(LANG_KEY);
    if (LANGS.includes(stored)) return stored;
  } catch (_) {
    // Ohne Speicher im Fenster gilt die Sprache des Rechners.
  }
  return String(navigator.language || "de").toLowerCase().startsWith("de") ? "de" : "en";
}

const LANG = chosenLang();

function chooseLang(lang) {
  try {
    window.localStorage.setItem(LANG_KEY, lang);
  } catch (_) {
    // Dann gilt die Wahl nur bis zum nächsten Öffnen.
  }
  window.location.reload();
}

// Die Sprache, in der Excel Formeln zeigt. Erst nach Office.onReady bekannt; davor wie das Fenster.
function formulaLang() {
  try {
    const shown = Office.context.displayLanguage;
    if (shown) return String(shown).toLowerCase().startsWith("de") ? "de" : "en";
  } catch (_) {
    // office.js ist noch nicht bereit.
  }
  return LANG;
}

// Eine Korrektur, wie Excel sie in seiner Sprache zeigt.
function shownFormula(fix) {
  return formulaLang() === "en" && fix && fix.english ? fix.english : (fix && fix.german) || "";
}

// „Hinweise|Gruppe“: dasselbe deutsche Wort, im Englischen je nach Stelle anders.
// Sagt Gegenprobe, in welcher Sprache es schreiben soll und wie Excel Formeln zeigt.
function langHeaders() {
  return { "X-Gegenprobe-Sprache": LANG, "X-Gegenprobe-Formeln": formulaLang() };
}

function tr(text, ...args) {
  const template = (LANG === "en" && EN[text]) || text.split("|")[0];
  return template.replace(/\{(\d+)\}/g, (all, i) => (args[Number(i)] === undefined ? all : String(args[Number(i)])));
}

function applyStatic(root) {
  document.documentElement.lang = LANG;
  if (LANG !== "en") return;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  for (const node of nodes) {
    const parent = node.parentElement;
    if (!parent || parent.closest("script, style, [data-en]")) continue;
    const text = node.nodeValue.replace(/\s+/g, " ");
    const key = text.trim();
    if (!key || !EN[key]) continue;
    const lead = text.startsWith(" ") ? " " : "";
    const tail = text.endsWith(" ") && key ? " " : "";
    node.nodeValue = lead + EN[key] + tail;
  }
  for (const node of root.querySelectorAll("[data-en]")) node.textContent = node.dataset.en;
  for (const attr of ["title", "aria-label", "placeholder"]) {
    for (const node of root.querySelectorAll(`[${attr}]`)) {
      const key = node.getAttribute(attr).replace(/\s+/g, " ").trim();
      if (EN[key]) node.setAttribute(attr, EN[key]);
    }
  }
}

const EN = {
  // --- Fenster (pruefer.html) ---
  "Gegenprobe Prüfer": "Gegenprobe checker",
  "Über Gegenprobe: Grenzen und Datenschutz": "About Gegenprobe: limits and privacy",
  "Über Gegenprobe": "About Gegenprobe",
  "Neu prüfen": "Check again",
  "Ich verbinde mich mit Excel …": "Connecting to Excel …",
  "Bevor du loslegst": "Before you start",
  "Gegenprobe kann sich irren": "Gegenprobe can be wrong",
  "Gegenprobe rechnet deine Mappe nach und sucht typische Fehler. Es kann Fehler übersehen und auch etwas melden, das richtig ist. Es ersetzt deine eigene Kontrolle nicht: Für deine Zahlen und Entscheidungen bleibst du verantwortlich. Schau dir jede Korrektur an, bevor du sie einsetzt. „Rückgängig“ nimmt sie wieder heraus.":
    "Gegenprobe recalculates your workbook and looks for typical mistakes. It can miss mistakes and can flag something that is correct. It does not replace your own review: you stay responsible for your numbers and decisions. Look at every fix before you apply it. “Undo” takes it back out.",
  "Deine Daten bleiben bei dir": "Your data stays with you",
  "Alles läuft auf diesem Rechner, ohne KI, ohne Konto und ohne Internet. Die Zellen gehen nur an Gegenprobe hier, nie an einen Server. Aus dem Internet lädt Excel nur die Office-Bibliothek von Microsoft (office.js), ohne deine Daten.":
    "Everything runs on this computer, with no AI, no account and no internet. Cells go only to Gegenprobe here, never to a server. The only thing Excel loads from the internet is Microsoft's Office library (office.js), without your data.",
  "Gegenprobe speichert keine Inhalte deiner Mappen. Das Protokoll enthält keine Zellen, Namen oder Werte.":
    "Gegenprobe stores no content of your workbooks. The log holds no cells, names or values.",
  "Nur drei Dinge bleiben, alle von dir ausgelöst: „Passt so“ steht in der Mappe selbst, als Prüfsumme ohne Inhalt. Eigene Spaltenregeln liegen auf diesem Rechner. Einen Prüfbericht, den du anforderst, löscht Gegenprobe nach einem Tag.":
    "Only three things are kept, all triggered by you: “Looks right” is stored in the workbook itself, as a checksum without content. Your own column rules stay on this computer. A report you ask for is deleted after one day.",
  "Ohne Gewähr": "No warranty",
  "Gegenprobe ist kostenlos und quelloffen (MIT-Lizenz) und wird ohne jede Gewähr bereitgestellt.":
    "Gegenprobe is free and open source (MIT license) and is provided without any warranty.",
  "Sprache": "Language",
  "Verstanden": "Got it",
  "Schließen": "Close",
  "Hinweise": "Findings",
  "Herleitung": "Trace",
  "Mappe": "Workbook",
  "Datei": "File",
  "Excel- oder CSV-Datei hierher ziehen": "Drop an Excel or CSV file here",
  "oder hier klicken und wählen (.xlsx oder .csv). Eine frühere Version vergleiche ich mit der offenen Mappe: was sich geändert hat und warum ein Ergebnis jetzt anders ist. Jede Datei kann ich auch prüfen, ohne sie zu öffnen, eine CSV aus einem Export auf ihre Daten.":
    "or click here to choose one (.xlsx or .csv). I compare an earlier version with the open workbook: what changed and why a result is different now. I can also check any file without opening it, and a CSV from an export for its data.",
  "Mit offener Mappe vergleichen": "Compare with open workbook",
  "Prüfbericht über diese Datei:": "Report on this file:",
  "Ohne Daten": "Without data",
  "Mit Werten": "With values",
  "Die Datei liest Gegenprobe nur in den Arbeitsspeicher, sie wird nicht gespeichert und geht nicht ins Internet.":
    "Gegenprobe reads the file into memory only. It is not saved and does not go to the internet.",
  "Hier steht, was die ganze Mappe rechnet: woher die Daten kommen, was jede gerechnete Spalte und jede Formel tut und woraus die Ergebnisse entstehen.":
    "This shows what the whole workbook calculates: where the data comes from, what each calculated column and each formula does, and what the results are made of.",
  "Markiere in Excel eine Zelle, dann zeige ich hier, woraus ihre Zahl entsteht und wohin sie fließt.":
    "Select a cell in Excel and I show here what its number is made of and where it flows.",
  "Auf den anderen Blättern": "On the other sheets",
  "Eigene Regeln je Spalte": "Your own rules per column",
  "Leg fest, was in einer Spalte stehen darf, zum Beispiel Alter von 0 bis 120, kein Datum in der Zukunft oder Pflichtfeld. Gegenprobe meldet jeden Verstoß mit Zelle und Grund, in jeder Mappe mit einer Spalte dieses Namens. Die Regeln bleiben auf diesem Rechner.":
    "Set what a column may contain, for example age from 0 to 120, no date in the future, or required. Gegenprobe reports every violation with cell and reason, in every workbook with a column of that name. The rules stay on this computer.",
  "Regel für die markierte Spalte": "Rule for the selected column",
  "Pflichtfeld: darf nicht leer sein": "Required: must not be empty",
  "Art": "Type",
  "beliebig": "any",
  "Zahl": "Number",
  "ganze Zahl": "Whole number",
  "Datum": "Date",
  "Text": "Text",
  "von": "from",
  "bis": "to",
  "nicht vor dem": "not before",
  "nicht nach dem": "not after",
  "nicht in der Zukunft": "not in the future",
  "Erlaubte Werte": "Allowed values",
  "z. B. m, w, d": "e.g. m, f, d",
  "Mit Komma getrennt, Zahlen mit Komma mit Strichpunkt („0,5; 1; 1,5“). Groß und klein zählt.":
    "Separated by commas; numbers with a decimal comma by semicolons (“0,5; 1; 1,5”). Upper and lower case count.",
  "Jeder Wert nur einmal (Fallnummer, Patienten-ID)": "Each value only once (case number, patient ID)",
  "Muster": "Pattern",
  "z. B. PAT-#####": "e.g. PAT-#####",
  "# eine Ziffer, @ ein Buchstabe, ? ein Zeichen, * beliebig viele, alles andere genau so.":
    "# a digit, @ a letter, ? one character, * any number of characters, everything else exactly as written.",
  "Vergleich": "Comparison",
  "Spalte": "Column",
  "Speichern": "Save",
  "Abbrechen": "Cancel",
  "Wovon die Ergebnisse abhängen": "What the results depend on",
  "Jede Eingabe und jede feste Zahl in den Formeln um 1 % erhöht und alles neu gerechnet: So viel Prozent ändert sich das Ergebnis.":
    "Each input and each fixed number in the formulas raised by 1 % and everything recalculated: this is how many percent the result changes.",
  "Was ich sonst nachgerechnet habe": "What else I recalculated",
  "Prüfbericht zum Weitergeben": "Report to pass on",
  "Alle Blätter auf einer Seite. Sie geht im Browser auf; dort drucken oder als PDF sichern.":
    "All sheets on one page. It opens in the browser; print it there or save it as PDF.",
  "„Ohne Daten“ zeigt nur Art, Stufe und Zelle jedes Hinweises, ohne Werte, Namen, Blatt- und Mappennamen. Den kannst du weitergeben, auch wenn in der Mappe Patientendaten stehen.":
    "“Without data” shows only kind, level and cell of each finding, without values, names, sheet or workbook names. You can pass it on even when the workbook holds patient data.",
  "Zum Ausprobieren:": "To try it out:",
  "Übungsmappe mit eingebauten Fehlern": "Practice workbook with built-in mistakes",
  "(erfundene Daten) in Excel öffnen.": "(made-up data), open it in Excel.",
  "Ohne KI und ohne Internet: Die Zellen gehen nur an Gegenprobe auf diesem Rechner, gespeichert wird nichts davon (":
    "No AI and no internet: cells go only to Gegenprobe on this computer, and none of it is stored (",
  "mehr dazu": "more",
  "). Ein Klick auf einen Hinweis oder eine Zelle in der Herleitung oder unter Mappe markiert sie in Excel. Geändert wird nur, was du mit einem Knopf wie „Einsetzen“ änderst, und „Rückgängig“ nimmt es wieder heraus („Aktualisieren“ rechnet eine Pivot-Tabelle nur neu aus ihren Daten).":
    "). Clicking a finding, or a cell in Trace or Workbook, selects it in Excel. Only what you change with a button like “Apply” is changed, and “Undo” takes it back out (“Refresh” only recalculates a pivot table from its data).",

  // --- Meldungen (pruefer.js) ---
  " Die Mappe hat mehr als {0} Blätter: 1 Blatt habe ich nicht gelesen. Wechsle auf ein Blatt, dann prüfe ich es.":
    " The workbook has more than {0} sheets: I did not read 1 sheet. Switch to a sheet and I check it.",
  " Die Mappe hat mehr als {0} Blätter: {1} Blätter habe ich nicht gelesen. Wechsle auf ein Blatt, dann prüfe ich es.":
    " The workbook has more than {0} sheets: I did not read {1} sheets. Switch to a sheet and I check it.",
  " Eine Korrektur habe ich weggelassen, weil sie nicht mit den anderen zusammenpasst. Was noch fehlt, steht unten.":
    " I left out one fix because it does not fit with the others. What is still missing is below.",
  " auf Blatt {0}": " on sheet {0}",
  " mit {0}": " with {0}",
  " nur auf „{0}“": " only on “{0}”",
  " {0} Formeln kann ich noch nicht nachrechnen, dort zählt der Wert aus Excel.":
    " I cannot recalculate {0} formulas yet; there the value from Excel counts.",
  " {0} Korrekturen habe ich weggelassen, weil sie nicht mit den anderen zusammenpassen. Was noch fehlt, steht unten.":
    " I left out {0} fixes because they do not fit with the others. What is still missing is below.",
  "1 Formel nachgerechnet": "1 formula recalculated",
  "{0} Formeln nachgerechnet": "{0} formulas recalculated",
  "1 Formel": "1 formula",
  "{0} Formeln": "{0} formulas",
  "1 Hinweis": "1 note",
  "{0} Hinweise": "{0} notes",
  "1 frühere Korrektur": "1 earlier fix",
  "{0} frühere Korrekturen": "{0} earlier fixes",
  "1 Hinweis mit „Passt so“ ausgeblendet.": "1 finding hidden with “Looks right”.",
  "{0} Hinweise mit „Passt so“ ausgeblendet.": "{0} findings hidden with “Looks right”.",
  "Passt so": "Looks right",
  "Wieder zeigen": "Show again",
  "Aktualisieren": "Refresh",
  "Alte .xls-Dateien kann ich nicht lesen. Öffne sie in Excel und speichere sie als .xlsx.":
    "I cannot read old .xls files. Open it in Excel and save it as .xlsx.",
  "Am Diagramm „{0}“ hat sich inzwischen etwas geändert. Ich prüfe neu, dann kannst du es noch einmal versuchen.":
    "Chart “{0}” has changed in the meantime. I check again, then you can try once more.",
  "Berichtsfilter {0} auf „{1}“": "Report filter {0} set to “{1}”",
  "Blatt {0} öffnen": "Open sheet {0}",
  "Blendet diesen Hinweis in dieser Mappe aus. Ändert sich an der Stelle etwas, zeigt Gegenprobe ihn wieder.":
    "Hides this finding in this workbook. If something changes there, Gegenprobe shows it again.",
  "Das Blatt ist geschützt. Unter Überprüfen → Blattschutz aufheben geht es wieder.":
    "The sheet is protected. Review → Unprotect Sheet makes it work again.",
  "Das Blatt ist sehr lang: Ich vergleiche die ersten 20 000 Zeilen.": "The sheet is very long: I compare the first 20,000 rows.",
  "Das Blatt oder die Zelle gibt es so nicht mehr. Prüf nochmals mit „Neu prüfen“ (der Pfeil oben rechts).":
    "That sheet or cell no longer exists. Check again with “Check again” (the arrow at the top right).",
  "Das Blatt {0} kann ich nicht öffnen.": "I cannot open sheet {0}.",
  "Das Blatt „{0}“ ist ausgeblendet. Zum Ansehen: Rechtsklick auf einen Blattreiter → Einblenden.":
    "Sheet “{0}” is hidden. To see it: right-click a sheet tab → Unhide.",
  "Das ging nicht. Setz die Korrekturen einzeln ein.": "That did not work. Apply the fixes one by one.",
  "Das ging nicht. {0} Der Tipp oben sagt, wie es in Excel von Hand geht.":
    "That did not work. {0} The tip above says how to do it by hand in Excel.",
  "Das ist eine Numbers-Datei. Exportiere sie in Numbers als Excel-Datei (Ablage → Exportieren → Excel), dann prüfe ich sie.":
    "This is a Numbers file. Export it from Numbers as an Excel file (File → Export To → Excel), then I check it.",
  "Das ist eine Tabelle aus LibreOffice oder OpenOffice (.ods). Speichere sie dort als Excel-Datei (.xlsx), dann prüfe ich sie.":
    "This is a LibreOffice or OpenOffice spreadsheet (.ods). Save it there as an Excel file (.xlsx), then I check it.",
  "Das ist für Excel auf einmal zu viel. Versuch es in einem Moment noch einmal.": "That is too much for Excel at once. Try again in a moment.",
  "Dazwischen": "In between",
  "Den Bericht konnte ich nicht schreiben. Läuft Gegenprobe noch?": "I could not write the report. Is Gegenprobe still running?",
  "Der Baum ist sehr groß, ich zeige die ersten 80 Zellen.": "The tree is very large; I show the first 80 cells.",
  "Der Bericht ist im Browser aufgegangen.": "The report opened in the browser.",
  "Der Filter {0} hat sich inzwischen geändert. Ich prüfe neu.": "Filter {0} has changed in the meantime. I check again.",
  "Der Name {0} hat sich inzwischen geändert. Ich prüfe neu.": "Name {0} has changed in the meantime. I check again.",
  "Der Name {0} zeigt jetzt auf {1}.": "Name {0} now refers to {1}.",
  "Diagramm „{0}“: {1}": "Chart “{0}”: {1}",
  "Die Blätter": "The sheets",
  "Die Datei hat kein Blatt wie die offene Mappe, darum prüfe ich sie für sich.":
    "The file has no sheet like the open workbook, so I check it on its own.",
  "Die Mappe kann ich gerade nicht beschreiben.": "I cannot describe the workbook right now.",
  "Die Pivot-Tabelle „{0}“ ist aktualisiert.": "Pivot table “{0}” is refreshed.",
  "Die Pivot-Tabelle „{0}“ nimmt jetzt die Daten aus {1}.": "Pivot table “{0}” now takes its data from {1}.",
  "Die Pivot-Tabelle „{0}“ zeigt jetzt nur {1} = „{2}“.": "Pivot table “{0}” now shows only {1} = “{2}”.",
  "Die Quelle der Pivot-Tabelle „{0}“ hat sich inzwischen geändert. Ich prüfe neu.":
    "The source of pivot table “{0}” has changed in the meantime. I check again.",
  "Die Regel für „{0}“ ist gelöscht. Ich prüfe neu.": "The rule for “{0}” is deleted. I check again.",
  "Die Regel für „{0}“ ist gespeichert. Ich prüfe neu.": "The rule for “{0}” is saved. I check again.",
  "Die Tabelle {0} geht jetzt bis {1}, die Pivot-Tabelle „{2}“ ist aktualisiert.":
    "Table {0} now reaches {1}, and pivot table “{2}” is refreshed.",
  "Die Tabelle {0} hat sich inzwischen geändert. Ich prüfe neu.": "Table {0} has changed in the meantime. I check again.",
  "Die Zelle ist leer.": "The cell is empty.",
  "Die Zelle {0} kann ich nicht markieren.": "I cannot select cell {0}.",
  "Die Übungsmappe kommt über den Browser, in Downloads.": "The practice workbook comes through the browser, into Downloads.",
  "Die Übungsmappe konnte ich nicht öffnen. Läuft Gegenprobe noch?": "I could not open the practice workbook. Is Gegenprobe still running?",
  "Diese Excel-Version kann den Filter nicht von hier aus stellen. Der Tipp oben sagt, wie es von Hand geht.":
    "This Excel version cannot set the filter from here. The tip above says how to do it by hand.",
  "Diese Excel-Version kann die Tabelle nicht von hier aus erweitern. Der Tipp oben sagt, wie es von Hand geht.":
    "This Excel version cannot extend the table from here. The tip above says how to do it by hand.",
  "Dieses Blatt gibt es in der offenen Mappe nicht mehr.": "This sheet no longer exists in the open workbook.",
  "Dieses Blatt ist neu.": "This sheet is new.",
  "Dieses Fenster gehört in Excel: Start, dann Prüfer.": "This panel belongs in Excel: Home, then Checker.",
  "Du schreibst gerade in einer Zelle. Drück Enter oder Esc, dann geht es.": "You are editing a cell. Press Enter or Esc, then it works.",
  "Eingabe": "Input",
  "Eingabe: {0}": "Input: {0}",
  "Eingesetzt im Diagramm „{0}“.": "Applied to chart “{0}”.",
  "Eingesetzt in {0}.": "Applied to {0}.",
  "Eingesetzt in {0}{1}: ": "Applied to {0}{1}: ",
  "Einsetzen ging nicht. {0} Du kannst die Korrektur oben auch selbst eintippen.":
    "Applying did not work. {0} You can also type the fix above yourself.",
  "Einsetzen": "Apply",
  "Entsteht aus": "Made from",
  "Ergebnis": "Result",
  "Ergebnisse": "Results",
  "Erst muss ich die Datei prüfen: Prüfen, dann noch einmal.": "I need to check the file first: Check, then try again.",
  "Erst muss ich die Mappe prüfen: Neu prüfen, dann noch einmal.": "I need to check the workbook first: Check again, then try again.",
  "Excel hat das gerade nicht zugelassen. Versuch es in einem Moment noch einmal.": "Excel did not allow that just now. Try again in a moment.",
  "Excel lässt sich gerade nicht lesen. {0} Sonst versuche ich es bei der nächsten Änderung wieder.":
    "I cannot read Excel right now. {0} Otherwise I try again at the next change.",
  "Excel nimmt das so nicht an. Prüf nochmals mit „Neu prüfen“ (der Pfeil oben rechts), dann sehe ich den neuen Stand.":
    "Excel does not accept that. Check again with “Check again” (the arrow at the top right), then I see the new state.",
  "Excel zeigt {0}, nachgerechnet ergibt die Formel aber {1}.": "Excel shows {0}, but recalculated the formula gives {1}.",
  "Excel zeigt {0}. Diese Formel kann ich noch nicht nachrechnen.": "Excel shows {0}. I cannot recalculate this formula yet.",
  "Fehler": "Error",
  "Fehler|Gruppe": "Errors",
  "Hinweise|Gruppe": "Notes",
  "Feste Zahl": "Fixed number",
  "Feste Zahlen in der Formel: {0}": "Fixed numbers in the formula: {0}",
  "Fließt in": "Flows into",
  "Formel": "Formula",
  "Gegenprobe antwortet nicht. Doppelklick auf ExcelAddin.command im Ordner VoiceAgent startet den Prüfer im Hintergrund wieder und sagt, was fehlte.":
    "Gegenprobe does not answer. Double-click ExcelAddin.command in the VoiceAgent folder to restart the checker in the background; it says what was missing.",
  "Gegenprobe antwortet nicht. Doppelklick auf ExcelAddin.bat im Ordner VoiceAgent startet den Prüfer im Hintergrund wieder und sagt, was fehlte.":
    "Gegenprobe does not answer. Double-click ExcelAddin.bat in the VoiceAgent folder to restart the checker in the background; it says what was missing.",
  "Geleert: {0}.": "Cleared: {0}.",
  "Geleert: {0}{1}": "Cleared: {0}{1}",
  "Geprüft: {0} (nicht in Excel geöffnet)": "Checked: {0} (not opened in Excel)",
  "Herleitung {0}": "Trace {0}",
  "Hinweis": "Note",
  "Hängt vom heutigen Datum ab (HEUTE()).": "Depends on today's date (TODAY()).",
  "Ich konnte das Ergebnis nicht ganz zeigen. Prüf nochmals mit „Neu prüfen“ (der Pfeil oben rechts).":
    "I could not show the whole result. Check again with “Check again” (the arrow at the top right).",
  "Ich lese die ganze Mappe …": "Reading the whole workbook …",
  "Ich lese „{0}“: {1} von {2} Zellen …": "Reading “{0}”: {1} of {2} cells …",
  "Ich prüfe Excel-Mappen (.xlsx) und CSV-Dateien.": "I check Excel workbooks (.xlsx) and CSV files.",
  "Ich prüfe {0} …": "Checking {0} …",
  "Ich prüfe „{0}“ mit {1} Zellen, das dauert einen Moment …": "Checking “{0}” with {1} cells, this takes a moment …",
  "Ich rechne alle zusammen nach …": "Recalculating them all together …",
  "Ich schreibe den Bericht …": "Writing the report …",
  "Ich vergleiche {0} mit der offenen Mappe …": "Comparing {0} with the open workbook …",
  "Ich öffne die Übungsmappe …": "Opening the practice workbook …",
  "In Gegenprobe ist dabei ein Fehler passiert, deine Mappe ist unverändert. Versuch es noch einmal.":
    "Something went wrong in Gegenprobe; your workbook is unchanged. Try again.",
  "In {0} hat sich seit dem Einsetzen etwas geändert, darum lasse ich es so.":
    "Something in {0} has changed since the fix was applied, so I leave it as it is.",
  "In {0} steht inzwischen etwas anderes. Ich prüfe neu, dann kannst du es noch einmal versuchen.":
    "{0} holds something else by now. I check again, then you can try once more.",
  "Jetzt in der Spalte: {0}": "Now in the column: {0}",
  "Keine Formel liest diese Zelle.": "No formula reads this cell.",
  "Liest aus „{0}“{1}: {2}.": "Reads from “{0}”{1}: {2}.",
  "Löschen": "Delete",
  "Macht {0} leer": "Clears {0}",
  "Markiere in Excel eine einzelne Zelle der Spalte, dann noch einmal.": "Select a single cell of the column in Excel, then try again.",
  "Name {0} auf {1}": "Name {0} set to {1}",
  "Neue Regel für „{0}“": "New rule for “{0}”",
  "Nur Text oder Zahlen, keine Formeln.": "Only text or numbers, no formulas.",
  "Pivot-Tabelle „{0}“ mit den Daten aus {1}": "Pivot table “{0}” with the data from {1}",
  "Prüfen geht gerade nicht.": "Checking does not work right now.",
  "Prüfen": "Review",
  "Regel für „{0}“ ändern": "Change rule for “{0}”",
  "Regel ändern": "Change rule",
  "Rückgängig gemacht: In {0} steht wieder, was vorher dort war.": "Undone: {0} holds what was there before.",
  "Rückgängig gemacht: {0}.": "Undone: {0}.",
  "Rückgängig ging nicht. {0}": "Undo did not work. {0}",
  "Rückgängig": "Undo",
  "Schreibt die Korrektur in {0}": "Writes the fix into {0}",
  "Seit dem Einsetzen hat sich dort etwas geändert, darum lasse ich es so.": "Something has changed there since the fix was applied, so I leave it as it is.",
  "Sonst unverändert.": "Otherwise unchanged.",
  "Spalten": "Columns",
  "Stelle zum Ansehen": "spot to look at",
  "Stellen zum Ansehen": "spots to look at",
  "Summe nehmen": "Use sum",
  "Summe statt Anzahl in „{0}“": "Sum instead of count in “{0}”",
  "Tabelle {0} auf {1}": "Table {0} extended to {1}",
  "Tabelle {0} mit {1} Zeilen.": "Table {0} with {1} rows.",
  "Und {0} weitere Formeln.": "And {0} more formulas.",
  "Und {0} weitere Änderungen auf diesem Blatt.": "And {0} more changes on this sheet.",
  "Unverändert: {0}.": "Unchanged: {0}.",
  "Verbundene Zellen sind im Weg. Heb den Verbund auf (Start → Verbinden und zentrieren), dann geht es.":
    "Merged cells are in the way. Unmerge them (Home → Merge & Center), then it works.",
  "Vergleichen geht gerade nicht. Neu prüfen, dann noch einmal.": "Comparing does not work right now. Check again, then try once more.",
  "Wirklich?": "Sure?",
  "Zeilen": "Rows",
  "Zum Prüfen": "To review",
  "aktualisiert die Pivot-Tabelle „{0}“ aus ihren Daten": "refreshes pivot table “{0}” from its data",
  "baut die Pivot-Tabelle „{0}“ mit den Daten aus {1} neu auf, mit denselben Feldern":
    "rebuilds pivot table “{0}” with the data from {1}, with the same fields",
  "erweitert die Tabelle {0} auf {1} und aktualisiert die Pivot-Tabelle „{2}“": "extends table {0} to {1} and refreshes pivot table “{2}”",
  "erweitert die Tabelle {0} auf {1} und baut die Pivot-Tabelle „{2}“ damit neu auf, mit denselben Feldern":
    "extends table {0} to {1} and rebuilds pivot table “{2}” from it, with the same fields",
  "fasst „{0}“ als Summe zusammen": "summarizes “{0}” as a sum",
  "geprüft um {0}": "checked at {0}",
  "in {0}": "into {0}",
  "keine": "none",
  "leert {0}": "clears {0}",
  "lässt den Namen {0} auf {1} zeigen; die Formeln damit bleiben, wie sie sind":
    "makes name {0} refer to {1}; the formulas using it stay as they are",
  "nicht größer als": "not greater than",
  "nicht kleiner als": "not less than",
  "nicht nach / nicht größer als": "not after / not greater than",
  "nicht nach": "not after",
  "nicht nachgerechnet": "not recalculated",
  "nicht vor / nicht kleiner als": "not before / not less than",
  "nicht vor": "not before",
  "nur auf dem Blatt „{0}“": "only on sheet “{0}”",
  "nur auf diesem Blatt": "only on this sheet",
  "office.js von Microsoft lädt nicht. Ist der Mac mit dem Internet verbunden?":
    "Microsoft's office.js does not load. Is the computer connected to the internet?",
  "setzt die Korrekturen zusammen ein, nachgerechnet; ein Rückgängig nimmt alle zurück":
    "applies the fixes together, recalculated; one Undo takes them all back",
  "setzt im Diagramm „{0}“ {1}": "sets {1} in chart “{0}”",
  "stellt den Berichtsfilter {0} auf „{1}“": "sets report filter {0} to “{1}”",
  "und weitere": "and more",
  "und {0} weitere": "and {0} more",
  "wird noch geprüft": "still being checked",
  "{0} Fehler": "{0} errors",
  "{0} ({1}): 1 % mehr ergibt {2}": "{0} ({1}): 1 % more gives {2}",
  "Daraus lesen {0}.": "Read by {0}.",
  ", ergibt ": ", gives ",
  "Excel zeigt {0}. Nachgerechnet stimmt es.": "Excel shows {0}. Recalculated, it is right.",
  "Nachgerechnet: {0}.": "Recalculated: {0}.",
  "{0} bis {1} ({2} Zellen)": "{0} to {1} ({2} cells)",
  "Weiter unten rechnet {0} anders als Excel zeigt.": "Further down, {0} calculates differently from what Excel shows.",
  "„{0}“": "“{0}”",
  "kommt weiter oben schon vor": "already appears further up",
  "{0} Formeln, {1} Eingaben, {2} Ergebnisse, Rechenkette bis {3} Stufen tief.":
    "{0} formulas, {1} inputs, {2} results, calculation chain up to {3} levels deep.",
  "{0} auf- oder zuklappen": "expand or collapse {0}",
  "{0} eingesetzt": "{0} applied",
  "{0} in Excel markieren": "select {0} in Excel",
  "{0} zum Prüfen": "{0} to review",
  "Ändern": "Change",
  "Änderung": "Change",
  "Öffnen ging nicht, die Mappe liegt hier: {0}. In Excel über Datei → Öffnen.":
    "Opening did not work; the workbook is here: {0}. In Excel use File → Open.",
  "„{0}“ ist in Excel aufgegangen. Klick dort auf Start → Prüfer.": "“{0}” opened in Excel. Click Home → Checker there.",
  "„{0}“ summiert jetzt.": "“{0}” now sums.",
  "„{0}“ zählt inzwischen nicht mehr. Ich prüfe neu.": "“{0}” no longer counts by now. I check again.",
};

applyStatic(document.body);
