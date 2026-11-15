// Excel-Prüfer: alle paar Sekunden die offene Mappe nachrechnen lassen und die Hinweise zeigen.

const INTERVAL = 3000;
const LEVELS = { fehler: "Fehler", pruefen: "Prüfen" };

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function render(data) {
  document.getElementById("state").textContent = data.text;
  document.getElementById("where").textContent = data.book ? `${data.book}, Blatt ${data.sheet}` : "";
  const hints = data.hints || [];
  document.getElementById("hints").replaceChildren(...hints.map((h) => {
    const li = el("li", "rule-card");
    li.append(el("span", `badge ${h.level === "fehler" ? "fix" : "reason"}`, LEVELS[h.level] || "Hinweis"), el("p", "", h.text));
    if (h.tip) li.append(el("p", "sub", h.tip));
    return li;
  }));
  renderDrivers(data.drivers || []);
  const st = data.stats;
  document.getElementById("stats").textContent = st
    ? `${st.formeln} Formeln, ${st.eingaben} Eingaben, ${st.ergebnisse} Ergebnisse, Rechenkette bis ${st.tiefe} Stufen tief` +
      (st.nicht_nachrechenbar ? `. ${st.nicht_nachrechenbar} Formeln kann ich nicht nachrechnen (z. B. SVERWEIS oder andere Blätter), die zählen mit dem Wert aus Excel.` : ".")
    : "";
  document.getElementById("overview").replaceChildren(...(data.overview || []).map((o) => el("li", "", o.text)));
}

function percent(value) {
  const text = Math.abs(value).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return `${value < 0 ? "−" : "+"}${text} %`;
}

function renderDrivers(drivers) {
  const box = document.getElementById("drivers-box");
  box.hidden = !drivers.length;
  const groups = new Map();
  for (const d of drivers) {
    if (!groups.has(d.output)) groups.set(d.output, { value: d.output_value, items: [] });
    groups.get(d.output).items.push(d);
  }
  document.getElementById("drivers").replaceChildren(...[...groups].map(([output, g]) => {
    const card = el("div", "rule-card");
    card.append(el("strong", "", `${output} = ${g.value}`));
    const list = el("ul");
    for (const d of g.items) {
      list.append(el("li", "", `${d.input} (${d.input_value}): 1 % mehr ergibt ${percent(d.elasticity)}`));
    }
    card.append(list);
    return card;
  }));
}

async function check() {
  try {
    const response = await fetch("/api/pruefen");
    render(await response.json());
  } catch (_) {
    document.getElementById("state").textContent = "Gegenprobe antwortet nicht. Läuft es noch?";
  }
  setTimeout(check, INTERVAL);
}

check();
