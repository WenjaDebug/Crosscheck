// Startseite des Prüfers: ob das Excel-Fenster eingerichtet ist, und eine Datei prüfen, ohne sie zu öffnen.

const LEVELS = { fehler: "Fehler", pruefen: "Prüfen", info: "Hinweis" };
const fileState = document.getElementById("file-state");

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

async function addinState() {
  const note = document.getElementById("addin-state");
  let data;
  try {
    data = await (await fetch("/api/start", { cache: "no-store" })).json();
  } catch (_) {
    note.textContent = "Gegenprobe antwortet nicht. Läuft das Fenster mit VoiceAgent.command noch?";
    return;
  }
  // Unter Windows heißen die Starter .bat statt .command.
  const setup = data.windows ? "ExcelAddin.bat" : "ExcelAddin.command";
  const start = data.windows ? "VoiceAgent.bat" : "VoiceAgent.command";
  if (data.addin && data.dienst && (data.laeuft === false || data.neustart === false)) {
    // Ohne den Dienst zeigt Excel „Add-In-Fehler“, sobald dieses Fenster zu ist.
    const why = data.laeuft === false
      ? `Der Prüfer im Hintergrund ${data.abruf ? "ist gerade nicht bereit" : "läuft gerade nicht"} (${data.grund}).`
      : "Nach dem letzten Neustart des Macs ist der Prüfer im Hintergrund nicht von selbst gestartet.";
    note.textContent = `${why} Ohne dieses Fenster zeigt Excel dann „Add-In-Fehler“. Doppelklick auf ${setup} ` +
      (data.windows ? "startet ihn wieder und sagt, was fehlte." : "repariert das und sagt, ob macOS ihn im Hintergrund erlauben muss.");
  } else if (data.addin && data.dienst) {
    const how = data.abruf
      ? "Eingerichtet: Der Prüfer startet im Hintergrund von selbst, sobald du ihn in Excel öffnest, auch nach einem Neustart."
      : "Eingerichtet: Der Prüfer läuft still im Hintergrund.";
    note.textContent = `${how} Einfach Excel öffnen und unter Start auf Prüfer ` +
      `klicken (beim ersten Mal Start, Add-ins, Gegenprobe Prüfer). Diese Seite und ${data.windows ? "dieses Fenster" : "das Terminal"} ` +
      "brauchst du dafür nicht.";
  } else if (data.addin) {
    note.textContent = `Das Excel-Fenster ist eingerichtet, läuft aber nur, solange ${start} offen ist. Damit es ` +
      `immer da ist: ${setup} noch einmal doppelklicken.`;
  } else if (data.mac) {
    note.textContent = "Einmal einrichten: Im Programmordner ExcelAddin.command doppelklicken (fragt einmal nach deinem " +
      "Mac-Passwort), dann Excel ganz beenden und neu öffnen. Danach gibt es in Excel unter Start den Knopf Prüfer, " +
      "ohne Terminal und ohne Browser.";
  } else if (data.windows) {
    note.textContent = "Einmal einrichten: Im Programmordner ExcelAddin.bat doppelklicken (Windows fragt einmal, ob es dem " +
      "Zertifikat vertrauen soll: Ja), dann Excel ganz schließen und neu öffnen. Danach gibt es in Excel unter Start den " +
      "Knopf Prüfer, ohne Fenster und ohne Browser.";
  } else {
    note.textContent = "Das Fenster direkt in Excel gibt es nur am Mac und unter Windows. Prüf hier unten eine Datei oder " +
      "nimm den Excel-Prüfer im Browser.";
  }
}

async function checkFile(file) {
  if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
    fileState.textContent = /\.xls$/i.test(file.name)
      ? "Alte .xls-Dateien kann ich nicht lesen. Öffne sie in Excel und speichere sie als .xlsx."
      : "Das ist keine Excel-Datei (.xlsx).";
    return;
  }
  fileState.textContent = `Ich prüfe ${file.name} …`;
  document.getElementById("file-report").hidden = true;
  let data;
  try {
    const response = await fetch(`/api/addin/datei?name=${encodeURIComponent(file.name)}`, {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", "X-VoiceAgent": "1" },
      body: file,
    });
    if (!response.ok) throw new Error(String(response.status));
    data = await response.json();
  } catch (_) {
    fileState.textContent = "Gegenprobe antwortet nicht. Läuft das Fenster mit VoiceAgent.command noch?";
    return;
  }
  if (data.state !== "ok") {
    fileState.textContent = data.text || "Prüfen geht gerade nicht.";
    document.getElementById("file-result").replaceChildren();
    return;
  }
  fileState.textContent = `${data.book}: ${data.text}`;
  const parts = [];
  for (const sheet of data.sheets) {
    parts.push(el("h3", "", sheet.name));
    if (!sheet.hints.length) {
      parts.push(el("p", "sub", sheet.text));
      continue;
    }
    const list = el("ul", "rules");
    for (const h of sheet.hints) {
      const li = el("li", "rule-card");
      const head = el("p", "");
      head.append(el("span", `badge ${h.level === "fehler" ? "fix" : "reason"}`, LEVELS[h.level] || "Hinweis"));
      if (h.cell) head.append(" ", el("code", "", h.cell.split(",").join(", ")));
      li.append(head, el("p", "", h.text));
      if (h.tip) li.append(el("p", "sub", h.tip));
      list.append(li);
    }
    parts.push(list);
  }
  document.getElementById("file-result").replaceChildren(...parts);
  document.getElementById("file-report").hidden = !data.sheets.length;
  document.getElementById("file-report-state").textContent = "";
}

async function report(withData) {
  const note = document.getElementById("file-report-state");
  note.textContent = "Ich schreibe den Bericht …";
  try {
    const response = await fetch("/api/addin/bericht", {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-VoiceAgent": "1" },
      body: JSON.stringify({ daten: withData, datei: true }),
    });
    const data = await response.json();
    if (data.state !== "ok") {
      note.textContent = "Erst muss ich die Datei prüfen.";
      return;
    }
    window.open(data.url, "_blank");
    note.textContent = "Der Bericht ist in einem neuen Tab aufgegangen.";
  } catch (_) {
    note.textContent = "Den Bericht konnte ich nicht schreiben.";
  }
}

const drop = document.getElementById("drop");
drop.addEventListener("dragover", (e) => {
  e.preventDefault();
  drop.classList.add("over");
});
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  drop.classList.remove("over");
  const file = e.dataTransfer && e.dataTransfer.files[0];
  if (file) checkFile(file);
});
document.getElementById("file").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) checkFile(file);
  e.target.value = "";
});
document.getElementById("file-plain").addEventListener("click", () => report(false));
document.getElementById("file-full").addEventListener("click", () => report(true));
// Eine Datei, die neben dem Feld landet, öffnet der Browser sonst selbst.
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => e.preventDefault());

addinState();
