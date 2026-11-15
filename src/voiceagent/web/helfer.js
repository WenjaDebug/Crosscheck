// Excel-Helfer: Auftrag eintippen, Vorschau zeigen, nach dem Klick in Excel schreiben.

const NUMBER = /^[-+−]?(?!0\d)[\d.,'’ ]+%?$/;  // 044 123 45 67 ist eine Telefonnummer, keine Zahl
const EXAMPLES = [
  "Summe von Betrag pro Monat",
  "Anzahl pro Station",
  "Durchschnitt Alter pro Station und Geschlecht",
  "Summe Umsatz pro Region, Jahre in den Spalten",
  "Maximum der Wartezeit je Abteilung, größte zuerst",
];

const text = document.getElementById("text");
const state = document.getElementById("state");
const plan = document.getElementById("plan");
const done = document.getElementById("done");
const run = document.getElementById("run");
const show = document.getElementById("show");
let token = "";

function el(tag, className, content) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (content !== undefined) node.textContent = content;
  return node;
}

function examples(list) {
  // Mit den echten Spaltennamen, sobald Gegenprobe die Tabelle kennt.
  if (!list || !list.length) list = EXAMPLES;
  document.getElementById("examples").replaceChildren(...list.map((example) => {
    const chip = el("button", "chip", example);
    chip.type = "button";
    chip.addEventListener("click", () => {
      text.value = example;
      text.focus();
    });
    return chip;
  }));
}

async function post(url, body) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-VoiceAgent": "1" },
    body: JSON.stringify(body),
  });
  return response.json();
}

function renderPlan(data) {
  token = data.token;
  plan.hidden = false;
  document.getElementById("plan-text").textContent = data.text;
  document.getElementById("plan-notes").replaceChildren(...(data.notes || []).map((n) => el("li", "", n)));
  const table = document.getElementById("plan-table");
  table.replaceChildren(...(data.preview || []).map((row, i) => {
    const tr = el("tr", i === 0 || row[0] === "Gesamt" ? "head" : "");
    row.forEach((cell) => tr.append(el(i === 0 ? "th" : "td", NUMBER.test(cell) ? "num" : "", cell)));
    return tr;
  }));
  document.getElementById("real-box").hidden = !data.pivot_possible;
  document.getElementById("real-note").hidden = !data.pivot_possible;
  document.getElementById("real").checked = false;
  run.disabled = false;
}

async function preview(event) {
  if (event) event.preventDefault();
  if (!text.value.trim()) {
    text.focus();
    return;
  }
  show.disabled = true;
  done.hidden = true;
  state.textContent = "Ich lese die Daten aus Excel und rechne …";
  try {
    const data = await post("/api/helfer/vorschau", { text: text.value });
    if (data.examples) examples(data.examples);
    if (data.ok) {
      state.textContent = `Aus ${data.book}, Blatt ${data.source}`;
      renderPlan(data);
    } else {
      plan.hidden = true;
      state.textContent = data.text;
    }
  } catch (_) {
    state.textContent = "Gegenprobe antwortet nicht. Läuft es noch?";
  }
  show.disabled = false;
}

async function write() {
  run.disabled = true;
  state.textContent = "Ich schreibe in Excel und prüfe danach jede Zahl nach …";
  try {
    const data = await post("/api/helfer/ausfuehren", { token, pivot: document.getElementById("real").checked });
    if (data.changed && data.ok) {
      state.textContent = "Neue Vorschau, weil sich die Daten geändert haben.";
      renderPlan(data);
      return;
    }
    if (data.ok) {
      plan.hidden = true;
      done.hidden = false;
      document.getElementById("done-text").textContent = data.text;
      state.textContent = "Für den nächsten Auftrag klick in Excel wieder auf das Blatt mit deinen Daten.";
    } else {
      state.textContent = data.text;
      run.disabled = false;
    }
  } catch (_) {
    state.textContent = "Gegenprobe antwortet nicht. Läuft es noch?";
    run.disabled = false;
  }
}

document.getElementById("ask").addEventListener("submit", preview);
text.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    preview();
  }
});
run.addEventListener("click", write);

// Datei mit Tabellen: hochladen (nur zu Gegenprobe auf diesem Rechner), Tabelle wählen, Vorschau.
const tables = document.getElementById("tables");
let file = "";

function showTable(data) {
  if (data.ok) {
    state.textContent = `In die Mappe ${data.book}`;
    renderPlan(data);
  } else {
    plan.hidden = true;
    state.textContent = data.text;
  }
}

async function upload(chosen) {
  if (!chosen) return;
  done.hidden = true;
  state.textContent = `Ich lese ${chosen.name} …`;
  try {
    const response = await fetch(`/api/helfer/datei?name=${encodeURIComponent(chosen.name)}`, {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", "X-VoiceAgent": "1" },
      body: chosen,
    });
    const data = await response.json();
    file = data.file || "";
    const list = data.tables || [];
    tables.hidden = list.length < 2;
    tables.replaceChildren(...list.map((t, i) => new Option(`${t.name}: ${t.rows} Zeilen, ${t.columns} Spalten`, String(i))));
    showTable(data);
  } catch (_) {
    state.textContent = "Gegenprobe antwortet nicht. Läuft es noch?";
  }
}

document.getElementById("file").addEventListener("change", (e) => upload(e.target.files[0]));
tables.addEventListener("change", async () => {
  try {
    showTable(await post("/api/helfer/tabelle", { file, index: Number(tables.value) }));
  } catch (_) {
    state.textContent = "Gegenprobe antwortet nicht. Läuft es noch?";
  }
});
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => {
  e.preventDefault();
  upload(e.dataTransfer.files[0]);
});
examples();
