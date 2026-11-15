// Zustand: jeder Baustein mit ✓, ✗ oder – und einem Satz dazu. Darunter der Fehlerbericht zum Nachlesen.

const MARKS = { true: ["✓", "ok"], false: ["✗", "bad"], null: ["–", "neutral"] };

async function load() {
  const list = document.getElementById("checks");
  const report = document.getElementById("report");
  try {
    const response = await fetch("/api/zustand");
    const data = await response.json();
    if (!response.ok) throw new Error(data.detail || response.statusText);
    list.replaceChildren();
    for (const check of data.checks) {
      const [mark, kind] = MARKS[String(check.ok)];
      const li = document.createElement("li");
      li.className = `check ${kind}`;
      const sign = document.createElement("span");
      sign.className = "mark";
      sign.textContent = mark;
      const name = document.createElement("strong");
      name.textContent = check.name;
      const text = document.createElement("span");
      text.textContent = check.text;
      li.append(sign, name, text);
      list.append(li);
    }
  } catch (err) {
    list.replaceChildren(Object.assign(document.createElement("li"), { textContent: `Der Zustand konnte nicht geprüft werden: ${err.message}` }));
  }
  try {
    report.textContent = await (await fetch("/api/zustand/bericht")).text();
  } catch (err) {
    report.textContent = "";
  }
}

document.getElementById("again").addEventListener("click", load);
load();
