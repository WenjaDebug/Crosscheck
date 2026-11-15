// Gegenprobe Prüfer als Seitenfenster in Excel: die Mappe über Office.js lesen, von Gegenprobe auf diesem
// Rechner nachrechnen lassen und die Hinweise zeigen. Ein Klick auf einen Hinweis markiert die Zelle, „Einsetzen“
// schreibt die nachgerechnete Korrektur hinein, „Rückgängig“ wieder heraus.

// So viele Zellen liest das Fenster je Blatt und in der ganzen Mappe, in Stücken von CHUNK_CELLS (ein großes Stück
// auf einmal macht Excel zäh). Wie xlsx.MAX_CELLS und xlsx.MAX_BOOK_CELLS.
const MAX_CELLS = 500000;
const MAX_BOOK_CELLS = 1500000;
const CHUNK_CELLS = 20000;
const MAX_SHEETS = 20;
// So viele Diagramme, Reihen und Pivot-Tabellen je Blatt liest das Fenster höchstens.
const MAX_PARTS = 20;
// So viele Namen aus dem Namensmanager schickt das Fenster höchstens.
const MAX_NAMES = 500;
// So viele verbundene Bereiche je Blatt (wie xlsx.MAX_MERGED).
const MAX_MERGED = 5000;
// „Passt so“: je Blatt die Hinweise, die jemand angeschaut und für richtig befunden hat. Sie stehen in den
// Einstellungen der Mappe, gehen also mit der Datei mit, und kommen wieder, sobald sich an der Stelle etwas ändert.
const PASST = "gegenprobe.passt";
const MAX_PASST = 500;
let passt = {};
const QUIET_MS = 1200;
// Bei einem großen Blatt länger warten, bis es ruhig ist: Jede Prüfung liest es ganz.
const QUIET_BIG_MS = 3000;
const BIG_CELLS = 100000;
// So oft fragt das Fenster nach, solange Gegenprobe andere große Blätter im Hintergrund prüft.
const OTHERS_MS = 3000;
const LEVELS = { fehler: tr("Fehler"), pruefen: tr("Prüfen"), info: tr("Hinweis") };

const state = document.getElementById("state");
const where = document.getElementById("where");
const again = document.getElementById("again");
// Gelesene Blätter nach Name, damit nach einer Änderung nur das geänderte Blatt neu gelesen wird. Ein Blatt, das
// Gegenprobe schon hat (sent), geht danach nur noch mit seinem Stempel mit.
const sheets = new Map();
let stamps = 0;
// Große andere Blätter prüft Gegenprobe im Hintergrund: Solange, fragt das Fenster ab und zu nach, wie weit.
let othersTimer = null;
let book = "";
let active = "";
let tables = [];
let bookNames = [];
// Ausgeblendete Blätter: Gegenprobe prüft sie mit, Excel kann sie aber nicht zeigen, solange sie ausgeblendet sind.
const hiddenSheets = new Set();
let unread = 0;
let busy = false;
let waiting = false;
let timer = null;
// Herleitung: die gezeigte Zelle, und eine Zelle, die wir selbst markiert haben (dann bleibt der Baum stehen).
let shownPlace = null;
let ownSelect = "";
let selectTimer = null;
// Eingesetzte Korrekturen, die neueste zuerst, mit dem, was vorher in den Zellen stand. Sie bleiben stehen, wenn
// danach neu geprüft wird und der Hinweis verschwindet. Excel selbst kann Änderungen aus Add-ins nicht immer
// rückgängig machen.
const inserted = [];
const MAX_UNDO = 10;
// Antwortet Gegenprobe nicht, ist meist der Prüfer im Hintergrund nicht bereit; ExcelAddin.command (unter Windows
// ExcelAddin.bat) repariert ihn.
const SILENT = /^Win/.test(navigator.platform || "")
  ? tr("Gegenprobe antwortet nicht. Doppelklick auf ExcelAddin.bat im Ordner VoiceAgent startet den Prüfer im Hintergrund wieder und sagt, was fehlte.")
  : tr("Gegenprobe antwortet nicht. Doppelklick auf ExcelAddin.command im Ordner VoiceAgent startet den Prüfer im Hintergrund wieder und sagt, was fehlte.");
// Die Fassung dieses Fensters, wie Gegenprobe sie beim ersten Prüfen nennt. Läuft nach einem Update eine neue,
// lädt sich das Fenster neu, statt mit dem alten weiterzumachen.
let version = "";

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

// Wie xlsx.py: Prozent- und Datumsformate erkennen, damit Werte gleich aussehen wie aus der Datei.
function kindOf(format) {
  const code = String(format || "").replace(/"[^"]*"|\[[^\]]*\]|\\./g, "").toLowerCase();
  if (code.includes("%")) return "p";
  if (/[dy]|h+:m|m+:s/.test(code)) return "d";
  // Ganze Zahlen (#,##0): Eine 0,4 zeigt Excel dann als 0.
  const first = code.split(";")[0];
  if (/[0#]/.test(first) && !/[.e@]|general|standard/.test(first)) return "i";
  return null;
}

// Direkt nach einer Änderung rechnet Excel noch: Sonst läse das Prüfen alte Ergebnisse (J2 zeigt noch 216 statt 240).
async function settled(context) {
  if (!supports("1.9")) return;
  try {
    const app = context.workbook.application;
    for (let i = 0; i < 20; i++) {
      app.load("calculationState");
      await context.sync();
      if (app.calculationState === "Done") return;
      await new Promise((done) => setTimeout(done, 150));
    }
  } catch (_) {}
}

async function readSheet(context, ws, budget, progress) {
  const used = ws.getUsedRangeOrNullObject(true);
  used.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
  await context.sync();
  const stamp = `${Date.now()}-${++stamps}`;
  if (used.isNullObject) return { name: ws.name, row: 1, col: 1, values: [], formulas_at: [], kinds_at: [], stamp, size: 0 };
  // Riesige Blätter: nur der obere Teil (ganze Zeilen), sonst wird Excel zäh. Gegenprobe sagt dann, was fehlt.
  const width = used.columnCount;
  const rows = Math.min(used.rowCount, Math.max(1, Math.floor(Math.min(MAX_CELLS, Math.max(budget, width)) / width)));
  const step = Math.max(1, Math.floor(CHUNK_CELLS / width));
  const values = [];
  const formulas = [];
  const kinds = [];
  for (let start = 0; start < rows; start += step) {
    const count = Math.min(step, rows - start);
    const part = ws.getRangeByIndexes(used.rowIndex + start, used.columnIndex, count, width);
    part.load("values,formulas,numberFormat");
    await context.sync();
    for (let i = 0; i < count; i++) {
      const row = part.values[i];
      values.push(row);
      for (let j = 0; j < width; j++) {
        const f = part.formulas[i][j];
        // Eine Formel als Text (Zelle als Text formatiert): Excel gibt bei „formulas“ dann den Text selbst zurück.
        if (typeof f === "string" && f.startsWith("=") && f !== row[j]) formulas.push([start + i, j, f]);
        if (typeof row[j] === "number") {
          const kind = kindOf(part.numberFormat[i][j]);
          if (kind) kinds.push([start + i, j, kind]);
        }
      }
    }
    if (rows > step) progress(ws.name, (start + count) * width, rows * width);
  }
  const filtered = await readFilters(context, ws);
  const charts = await readCharts(context, ws);
  const pivots = await readPivots(context, ws);
  const merged = await readMerged(context, ws.getRangeByIndexes(used.rowIndex, used.columnIndex, rows, width));
  const hiddenRows = await readHiddenRows(context, ws, used.rowIndex + 1, used.rowIndex + rows);
  const sheet = { name: ws.name, row: used.rowIndex + 1, col: used.columnIndex + 1, values, formulas_at: formulas, kinds_at: kinds,
    cut: rows < used.rowCount, filtered, charts, pivots, merged, stamp, size: rows * width };
  // null: nicht gelesen (dann weiß Gegenprobe nicht, welche Zeilen man sieht), nicht „keine“.
  if (hiddenRows) sheet.hiddenRows = hiddenRows;
  return sheet;
}

// Ausgeblendete Zeilen (von Hand oder durch einen Filter) von first bis last: rowHidden ist true, wenn alle Zeilen
// eines Bereichs ausgeblendet sind, false, wenn keine, sonst null. Meist ist es gleich beim ersten Mal false. Sonst
// wird geteilt, bis jeder Teil ganz oder gar nicht ausgeblendet ist, alle Teile einer Stufe in einem Durchgang.
const MAX_HIDDEN = 50000;
const MAX_PROBES = 2000;
async function readHiddenRows(context, ws, first, last) {
  try {
    const hidden = [];
    let todo = [[first, last]];
    while (todo.length) {
      if (todo.length > MAX_PROBES) return null;
      const probes = todo.map(([a, b]) => {
        const range = ws.getRange(`${a}:${b}`);
        range.load("rowHidden");
        return [a, b, range];
      });
      await context.sync();
      const next = [];
      for (const [a, b, range] of probes) {
        if (range.rowHidden === true) {
          for (let r = a; r <= b; r++) hidden.push(r);
          if (hidden.length > MAX_HIDDEN) return null;
        } else if (range.rowHidden !== false && b > a) {
          const middle = Math.floor((a + b) / 2);
          next.push([a, middle], [middle + 1, b]);
        }
      }
      todo = next;
    }
    return hidden;
  } catch (err) {
    return null;
  }
}

// Verbundene Zellen (ExcelApi 1.13): Excel zeigt den Wert über alle Zeilen, er steht aber nur in der ersten. Formeln,
// Filter und Pivot-Tabellen sehen in den übrigen eine leere Zelle. Ohne 1.13 eben ohne diese Prüfung.
async function readMerged(context, range) {
  if (!supports("1.13")) return [];
  try {
    const areas = range.getMergedAreasOrNullObject();
    areas.load("isNullObject,areaCount");
    await context.sync();
    if (areas.isNullObject || !areas.areaCount) return [];
    areas.areas.load("items/address");
    await context.sync();
    return areas.areas.items.slice(0, MAX_MERGED).map((r) => r.address.slice(r.address.lastIndexOf("!") + 1).replace(/\$/g, ""));
  } catch (_) {
    return [];
  }
}

// Wie pruefer._number: ab 10 000 mit schmalem Abstand zwischen den Tausendern.
function cells(count) {
  return count >= 10000 ? String(count).replace(/\B(?=(\d{3})+$)/g, "\u202f") : String(count);
}

function supports(version) {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", version);
  } catch (_) {
    return false;
  }
}

// Diagramme: je Reihe die Bereiche der Werte und Beschriftungen, damit Gegenprobe sie mit der Tabelle vergleicht.
// Excel kennt das in Office.js erst ab ExcelApi 1.15; vorher ohne diese Prüfung.
async function readCharts(context, ws) {
  const found = [];
  if (!supports("1.15")) return found;
  try {
    const charts = ws.charts;
    charts.load("items/name");
    await context.sync();
    for (const chart of charts.items.slice(0, MAX_PARTS)) {
      const list = chart.series;
      list.load("items/name");
      await context.sync();
      let title = "";
      try {
        chart.title.load("text");
        await context.sync();
        title = chart.title.text || "";
      } catch (_) {
        // Ein Diagramm ohne Titel.
      }
      const series = [];
      for (const item of list.items.slice(0, MAX_PARTS)) {
        const one = { name: item.name, values: "", categories: "" };
        // Eine Reihe mit eingetippten Werten hat keinen Bereich: dann eben ohne.
        try {
          const values = item.getDimensionDataSourceString(Excel.ChartSeriesDimension.values);
          await context.sync();
          one.values = values.value;
        } catch (_) {}
        try {
          const categories = item.getDimensionDataSourceString(Excel.ChartSeriesDimension.categories);
          await context.sync();
          one.categories = categories.value;
        } catch (_) {}
        series.push(one);
      }
      found.push({ name: chart.name, title, series });
    }
  } catch (_) {}
  return found;
}

// Pivot-Tabellen: wo ihre Zahlen stehen, ihre Felder, was ausgeblendet ist, und woher die Daten kommen (ExcelApi 1.8,
// die Quelle ab 1.15, Beschriftungs- und Wertfilter ab 1.12).
async function readPivots(context, ws) {
  const found = [];
  if (!supports("1.8")) return found;
  try {
    const tables = ws.pivotTables;
    tables.load("items/name");
    await context.sync();
    for (const table of tables.items.slice(0, MAX_PARTS)) {
      try {
        found.push(await readPivot(context, table));
      } catch (_) {}
    }
  } catch (_) {}
  return found;
}

async function readPivot(context, table) {
  const layout = table.layout;
  layout.load("showRowGrandTotals,showColumnGrandTotals");
  const all = layout.getRange();
  all.load("address");
  const body = layout.getDataBodyRange();
  body.load("address");
  const groups = { rows: table.rowHierarchies, columns: table.columnHierarchies, filters: table.filterHierarchies };
  for (const group of Object.values(groups)) group.load("items/name");
  table.dataHierarchies.load("items/name,items/summarizeBy,items/showAs");
  const source = supports("1.15") ? table.getDataSourceString() : null;
  const kind = supports("1.15") ? table.getDataSourceType() : null;
  await context.sync();
  for (const group of Object.values(groups)) for (const h of group.items) h.fields.load("items/name");
  for (const d of table.dataHierarchies.items) d.field.load("name");
  await context.sync();
  const names = {};
  const fields = [];
  for (const [key, group] of Object.entries(groups)) {
    names[key] = group.items.map((h) => (h.fields.items[0] ? h.fields.items[0].name : h.name));
    for (const h of group.items) for (const f of h.fields.items) {
      f.items.load("items/name,items/visible");
      fields.push(f);
    }
  }
  const tests = supports("1.12")
    ? fields.flatMap((f) => [Excel.PivotFilterType.label, Excel.PivotFilterType.value, Excel.PivotFilterType.date].map((t) => f.isFiltered(t)))
    : null;
  await context.sync();
  const skipped = {};
  for (const f of fields) {
    const off = f.items.items.filter((i) => !i.visible).map((i) => i.name);
    if (off.length) skipped[f.name] = off.slice(0, 500);
  }
  const local = kind && kind.value === "LocalTable";
  return {
    name: table.name,
    area: all.address,
    body: body.address,
    rows: names.rows,
    columns: names.columns,
    filters: names.filters,
    values: table.dataHierarchies.items.map((d) => ({
      caption: d.name, field: d.field.name, function: d.summarizeBy, showAs: d.showAs ? d.showAs.calculation : "None",
    })),
    skipped,
    grand: [layout.showRowGrandTotals, layout.showColumnGrandTotals],
    // null: Ob ein Filter Einträge ausblendet, weiß diese Excel-Version nicht.
    filtered: tests ? tests.some((t) => t.value) : null,
    source: source && !local ? source.value : "",
    table: source && local ? source.value : "",
  };
}

// Die Excel-Tabellen der Mappe: für Formeln wie =[@Menge]*[@Preis] und für Pivot-Tabellen, die ihre Daten aus einer nehmen.
async function readTables(context) {
  try {
    const list = context.workbook.tables;
    list.load("items/name,items/showTotals,items/showHeaders");
    await context.sync();
    const parts = list.items.slice(0, MAX_PARTS * 5).map((t) => {
      const range = t.getRange();
      range.load("address");
      t.worksheet.load("name");
      t.columns.load("items/name");
      return [t, range];
    });
    await context.sync();
    return parts.map(([t, range]) => ({
      name: t.name, sheet: t.worksheet.name, address: range.address, totals: t.showTotals, headers: t.showHeaders,
      columns: t.columns.items.map((c) => c.name),
    }));
  } catch (_) {
    return [];
  }
}

// Namen aus dem Namensmanager (MwSt → =Daten!$I$1), für die ganze Mappe und je Blatt: Formeln wie =Netto*MwSt.
async function readNames(context, list) {
  const found = [];
  try {
    const book = context.workbook.names;
    book.load("items/name,items/formula");
    const own = list.map((ws) => {
      const names = ws.names;
      names.load("items/name,items/formula");
      return [ws.name, names];
    });
    await context.sync();
    for (const item of book.items.slice(0, MAX_NAMES)) found.push({ name: item.name, formula: String(item.formula || "") });
    for (const [sheet, names] of own) {
      for (const item of names.items.slice(0, MAX_NAMES)) found.push({ name: item.name, formula: String(item.formula || ""), sheet });
    }
  } catch (_) {
    // Ältere Excel-Versionen kennen formula bei Namen noch nicht: dann ohne Namen.
  }
  return found.slice(0, MAX_NAMES);
}

// Bereiche, in denen ein Filter gerade Zeilen ausblendet (auf dem Blatt oder in einer Excel-Tabelle).
async function readFilters(context, ws) {
  const found = [];
  try {
    const sheetFilter = ws.autoFilter;
    sheetFilter.load("isDataFiltered");
    const area = sheetFilter.getRangeOrNullObject();
    area.load("address");
    const tables = ws.tables;
    tables.load("items/name");
    await context.sync();
    if (sheetFilter.isDataFiltered && !area.isNullObject) found.push(area.address);
    const parts = tables.items.map((table) => {
      const filter = table.autoFilter;
      filter.load("isDataFiltered");
      const range = table.getRange();
      range.load("address");
      return [filter, range];
    });
    await context.sync();
    for (const [filter, range] of parts) if (filter.isDataFiltered) found.push(range.address);
  } catch (_) {
    // Ältere Excel-Versionen kennen Filter in Office.js noch nicht: dann ohne diese Prüfung.
  }
  return found;
}

async function readBook(onlyActive) {
  return Excel.run(async (context) => {
    const wb = context.workbook;
    wb.load("name");
    const list = wb.worksheets;
    list.load("items/name,items/visibility");
    const current = list.getActiveWorksheet();
    current.load("name");
    await context.sync();
    if (book !== wb.name) sheets.clear();
    book = wb.name;
    active = current.name;
    const names = list.items.map((ws) => ws.name);
    hiddenSheets.clear();
    for (const ws of list.items) if (ws.visibility && ws.visibility !== "Visible") hiddenSheets.add(ws.name);
    for (const name of [...sheets.keys()]) if (!names.includes(name)) sheets.delete(name);
    await settled(context);
    let budget = MAX_BOOK_CELLS;
    // Das aktive Blatt zuerst: Es bekommt seine Zellen auch dann, wenn die Mappe davor schon viele hat.
    // Mehr als 20 Blätter liest er nicht; das aktive aber immer, auch wenn es weiter hinten steht.
    let first = list.items.slice(0, MAX_SHEETS);
    const own = list.items.find((ws) => ws.name === active);
    if (own && !first.includes(own)) first = [own, ...first.slice(0, MAX_SHEETS - 1)];
    unread = Math.max(0, list.items.length - first.length);
    first.sort((a, b) => (b.name === active) - (a.name === active));
    for (const ws of first) {
      if (onlyActive && ws.name !== active && sheets.has(ws.name)) {
        budget -= sheets.get(ws.name).size || 0;
        continue;
      }
      // Bei großen Blättern sieht man, wie weit das Lesen ist.
      const read = await readSheet(context, ws, budget, (name, done, all) => {
        state.textContent = tr("Ich lese „{0}“: {1} von {2} Zellen …", name, cells(done), cells(all));
      });
      budget -= read.size;
      sheets.set(ws.name, read);
    }
    // In der Reihenfolge der Mappe: Gegenprobe zählt die Zellen der Blätter in derselben Reihenfolge wie hier.
    const keep = new Set(first.map((ws) => ws.name));
    const ordered = names.filter((name) => keep.has(name) && sheets.has(name)).map((name) => [name, sheets.get(name)]);
    sheets.clear();
    for (const [name, read] of ordered) sheets.set(name, read);
    tables = await readTables(context);
    bookNames = await readNames(context, first);
  });
}

// Was das Fenster schickt: neu gelesene Blätter ganz, die übrigen nur mit ihrem Stempel.
function payload() {
  return { book, active, tables, names: bookNames, passt,
           sheets: [...sheets.values()].map((s) => ({ ...(s.sent ? { name: s.name, same: s.stamp } : s),
                                                      ...(hiddenSheets.has(s.name) ? { hidden: true } : {}) })) };
}

function loadPasst() {
  try {
    const stored = Office.context.document.settings.get(PASST);
    passt = stored && typeof stored === "object" && !Array.isArray(stored) ? stored : {};
  } catch (_) {
    passt = {};
  }
}

function savePasst() {
  try {
    Office.context.document.settings.set(PASST, passt);
    Office.context.document.settings.saveAsync(() => {});
  } catch (_) {
    // Ohne Einstellungen gilt „Passt so“ nur, bis das Fenster neu lädt.
  }
}

function passtButton(sheet, hint, li) {
  const button = el("button", "addin-passt", tr("Passt so"));
  button.type = "button";
  button.title = tr("Blendet diesen Hinweis in dieser Mappe aus. Ändert sich an der Stelle etwas, zeigt Gegenprobe ihn wieder.");
  button.addEventListener("click", (e) => {
    e.stopPropagation();
    passt[sheet] = [...(passt[sheet] || []).filter((id) => id !== hint.id), hint.id].slice(-MAX_PASST);
    savePasst();
    li.hidden = true;
    check(true);
  });
  button.addEventListener("keydown", (e) => e.stopPropagation());
  return button;
}

function passtNote(sheet, count) {
  const li = el("li", "addin-passt-note");
  li.append(el("span", "sub small", count === 1 ? tr("1 Hinweis mit „Passt so“ ausgeblendet.") : tr("{0} Hinweise mit „Passt so“ ausgeblendet.", count)));
  const button = el("button", "addin-passt", tr("Wieder zeigen"));
  button.type = "button";
  button.addEventListener("click", () => {
    delete passt[sheet];
    savePasst();
    check(true);
  });
  li.append(button);
  return li;
}

// Antwortet Gegenprobe mit einem Fehler (statt gar nicht), sagt seine Antwort, was los ist: Das zeigt das Fenster
// statt „antwortet nicht“.
class ServerError extends Error {}

async function post(payload, path = "/api/addin/pruefen") {
  const response = await fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-VoiceAgent": "1", ...langHeaders() },
    body: JSON.stringify(payload),
  });
  if (!response.ok) throw await failure(response);
  return response.json();
}

function silent(err) {
  return err instanceof ServerError ? err.message : SILENT;
}

async function failure(response) {
  let text = "";
  try {
    text = (await response.json()).text || "";
  } catch (_) {}
  return new ServerError(text || tr("In Gegenprobe ist dabei ein Fehler passiert, deine Mappe ist unverändert. Versuch es noch einmal."));
}

// Was Excel meldet, wenn es etwas nicht zulässt, in einem Satz mit dem, was man tun kann. Die Codes sind die von
// Office.js (OfficeExtension.Error.code); unbekannte bekommen einen allgemeinen Satz, nie den englischen Text.
function excelProblem(err) {
  const code = (err && err.code) || "";
  const message = String((err && err.message) || "").toLowerCase();
  if (code === "InvalidOperationInCellEditMode" || message.includes("cell edit") || message.includes("bearbeitungsmodus")) {
    return tr("Du schreibst gerade in einer Zelle. Drück Enter oder Esc, dann geht es.");
  }
  if (code === "AccessDenied" || message.includes("protected") || message.includes("geschützt")) {
    return tr("Das Blatt ist geschützt. Unter Überprüfen → Blattschutz aufheben geht es wieder.");
  }
  if (code === "ItemNotFound") return tr("Das Blatt oder die Zelle gibt es so nicht mehr. Prüf nochmals mit „Neu prüfen“ (der Pfeil oben rechts).");
  if (code === "InvalidArgument" || code === "InvalidReference") {
    return tr("Excel nimmt das so nicht an. Prüf nochmals mit „Neu prüfen“ (der Pfeil oben rechts), dann sehe ich den neuen Stand.");
  }
  if (code === "RequestPayloadSizeLimitExceeded" || code === "ResponsePayloadSizeLimitExceeded" ||
      code === "ActivityLimitReached") {
    return tr("Das ist für Excel auf einmal zu viel. Versuch es in einem Moment noch einmal.");
  }
  if (code === "Conflict" || message.includes("merge")) {
    return tr("Verbundene Zellen sind im Weg. Heb den Verbund auf (Start → Verbinden und zentrieren), dann geht es.");
  }
  return tr("Excel hat das gerade nicht zugelassen. Versuch es in einem Moment noch einmal.");
}

async function check(onlyActive = true) {
  if (busy) {
    waiting = true;
    return;
  }
  busy = true;
  again.disabled = true;
  try {
    await checkNow(onlyActive);
  } catch (_) {
    // Was hier noch schiefgeht (ein Hinweis, den das Fenster nicht zeigen kann), darf das Fenster nicht sperren.
    state.textContent = tr("Ich konnte das Ergebnis nicht ganz zeigen. Prüf nochmals mit „Neu prüfen“ (der Pfeil oben rechts).");
  } finally {
    busy = false;
    again.disabled = false;
  }
  if (waiting) {
    waiting = false;
    later();
  }
}

async function checkNow(onlyActive) {
  let data = null;
  try {
    await readBook(onlyActive && sheets.size > 0);
  } catch (err) {
    state.textContent = tr("Excel lässt sich gerade nicht lesen. {0} Sonst versuche ich es bei der nächsten Änderung wieder.", excelProblem(err));
  }
  if (book) {
    const size = sheets.has(active) ? sheets.get(active).size || 0 : 0;
    if (size > BIG_CELLS) state.textContent = tr("Ich prüfe „{0}“ mit {1} Zellen, das dauert einen Moment …", active, cells(size));
    try {
      data = await post(payload());
      if (data && data.state === "again") {
        // Gegenprobe kennt ein Blatt nicht mehr (zum Beispiel nach einem Neustart): alles noch einmal schicken.
        for (const s of sheets.values()) s.sent = false;
        data = await post(payload());
      }
      if (data && data.state !== "again") for (const s of sheets.values()) s.sent = true;
    } catch (err) {
      state.textContent = silent(err);
    }
  }
  if (data && data.version && version && data.version !== version) {
    location.reload();
    return;
  }
  if (data) {
    version = data.version || version;
    render(data);
    document.getElementById("when").textContent = tr("geprüft um {0}", new Date().toLocaleTimeString(LANG === "en" ? "en-GB" : "de-CH", { hour: "2-digit", minute: "2-digit", second: "2-digit" }));
    // Die Werte sind neu: die Herleitung der markierten Zelle auch, und die Mappe in Worten, wenn sie offen ist.
    if (shownPlace) explain(shownPlace, true);
    if (!document.getElementById("page-book").hidden) describe();
  }
}

function later() {
  clearTimeout(timer);
  const size = sheets.has(active) ? sheets.get(active).size || 0 : 0;
  timer = setTimeout(() => check(true), size > BIG_CELLS ? QUIET_BIG_MS : QUIET_MS);
}

// Ein ausgeblendetes Blatt kann Excel nicht zeigen: sagen, wie man es einblendet, statt still nichts zu tun.
function hiddenNote(sheet) {
  if (!hiddenSheets.has(sheet)) return false;
  state.textContent = tr("Das Blatt „{0}“ ist ausgeblendet. Zum Ansehen: Rechtsklick auf einen Blattreiter → Einblenden.", sheet);
  return true;
}

async function select(sheet, cell) {
  if (hiddenNote(sheet)) return;
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(sheet);
      ws.activate();
      // Mehrere Zellen („C5,C14:C16“, eigene Regeln) markiert Excel ab ExcelApi 1.9 zusammen, sonst die erste.
      if (cell.includes(",") && supports("1.9")) ws.getRanges(cell).select();
      else ws.getRange(cell.split(",")[0]).select();
      await context.sync();
    });
  } catch (_) {
    state.textContent = tr("Die Zelle {0} kann ich nicht markieren.", cellLabel(cell));
  }
}

// „C5, C14, …“ statt aller Zellen eines Hinweises, der viele markiert.
function cellLabel(cell) {
  const parts = String(cell || "").split(",");
  return parts.length > 2 ? `${parts[0]}, ${parts[1]}, …` : parts.join(", ");
}

// Die Zelle eines Hinweises wie im Namenfeld von Excel, mit dem Dreieck in der Ecke, das Excel an Zellen mit
// einem Problem zeigt: rot für Fehler, orange zum Prüfen, grau für einen Hinweis.
function cellChip(cell, level) {
  return el("span", `addin-ref-cell ${level || ""}`, cellLabel(cell));
}

// Liegt die Adresse im gelesenen Teil des Blatts? Sonst ist es keine Zelle, sondern etwa eine Fallnummer („F240000“).
function onSheet(sheet, ref) {
  const read = sheets.get(sheet);
  if (!read || !read.values.length) return false;
  const width = Math.max(...read.values.map((r) => r.length));
  return ref.split(":").every((part) => {
    const [, col, row] = /^([A-Z]+)(\d+)$/.exec(part);
    const c = columnNumber(col);
    const r = Number(row);
    return r >= read.row && r < read.row + read.values.length && c >= read.col && c < read.col + width;
  });
}

// Zellen, die im Text vorkommen („zählt nur C32:F32“), lassen sich anklicken und markieren sich dann in Excel.
const REF = /(^|[^A-Za-z0-9$!_.])([A-Z]{1,3}[1-9]\d{0,6}(?::[A-Z]{1,3}[1-9]\d{0,6})?)(?![A-Za-z0-9(_]|\.\d)/g;

// Formeln im Text („=SUMME(A21:E21)“) stehen wie in der Bearbeitungsleiste, nicht als Wörter im Satz.
const INLINE = /(^|[\s(„])(=[A-ZÄÖÜ$(][^\s]*?)(?=[.,;:]?(?:\s|$))/g;

function linked(text, sheet, skip = "") {
  const out = [];
  let at = 0;
  text = String(text);
  for (const m of text.matchAll(INLINE)) {
    const start = m.index + m[1].length;
    // Eine vorgelesene Formel mit Namen („=Umsatz * $B$12“) bleibt Text: Sie hat Leerzeichen.
    if (start < at || /[a-zäöüß]/.test(m[2].replace(/"[^"]*"/g, ""))) continue;
    out.push(...refs(text.slice(at, start), sheet, at === 0 ? skip : ""));
    out.push(el("code", "addin-inline", m[2]));
    at = start + m[2].length;
  }
  out.push(...refs(text.slice(at), sheet, at === 0 ? skip : ""));
  return out;
}

function refs(text, sheet, skip = "") {
  const out = [];
  let at = 0;
  // Was in „…“ steht, ist ein Wert aus der Tabelle („E11.9“, „F240000“), keine Zelle.
  const quoted = [...String(text).matchAll(/„[^“]*“/g)].map((q) => [q.index, q.index + q[0].length]);
  for (const m of String(text).matchAll(REF)) {
    const ref = m[2];
    const start = m.index + m[1].length;
    if (quoted.some(([a, b]) => start > a && start < b)) continue;
    if (start === 0 && ref === skip) {
      // Die Zelle des Hinweises steht schon vorne als Namenfeld.
      at = ref.length;
      continue;
    }
    if (!onSheet(sheet, ref)) continue;
    out.push(text.slice(at, start));
    const span = el("span", "addin-ref", ref);
    clickable(span, tr("{0} in Excel markieren", ref), (e) => {
      if (e) e.stopPropagation();
      select(sheet, ref);
    });
    span.addEventListener("keydown", (e) => e.stopPropagation());
    out.push(span);
    at = start + ref.length;
  }
  out.push(text.slice(at));
  return out;
}

// „Korrigiert: =SUMME(C32:G32), nachgerechnet 418,8013.“ (englisch „Corrected: =SUM(C32:G32), recalculated
// 418.8013.“): steht genau die Korrektur des Knopfs im Tipp, zeigt das Fenster sie einmal wie in der
// Bearbeitungsleiste von Excel, mit dem Ergebnis, statt sie zweimal zu sagen.
const FIXED_START = /^(?:Korrigiert|Corrected|Fixed)(?: in [A-Z]+\d+)?: /;
const FIXED = new RegExp(FIXED_START.source + "(.+?)(?:, (?:nachgerechnet|recalculated(?: this gives| it gives|:)?) (.+?))?" +
  "(?:,? (?:dann bis [A-Z]+\\d+ herunterziehen|then fill (?:it )?down to [A-Z]+\\d+))?\\.(?:\\s+([^]*))?$");

function fixTip(tip, fix) {
  // Ein Text wie „Dr. Schmidt“ hat selbst Punkte: Er steht genau so am Ende des Tipps.
  if (fix && fix.text && FIXED_START.test(tip || "") && tip.endsWith(`: ${fix.german}.`)) {
    return { value: "", rest: "" };
  }
  const m = FIXED.exec(String(tip || ""));
  if (!m || !fix || fix.clear || (m[1] !== shownFormula(fix) && m[1] !== fix.german)) return null;
  return { value: m[2] || "", rest: m[3] || "" };
}

const GROUPS = { fehler: tr("Fehler|Gruppe"), pruefen: tr("Zum Prüfen"), info: tr("Hinweise|Gruppe") };

// Oben die Zahl der Stellen und ein Balken nach Stufe, mit dem, was schon eingesetzt ist (grün).
let shownLevels = [];
let shownOk = false;

function renderLevels() {
  const count = document.getElementById("count");
  const label = document.getElementById("count-label");
  const verdict = document.getElementById("verdict");
  const n = shownLevels.length;
  count.className = `addin-big${!n && shownOk ? " ok" : ""}`;
  count.textContent = n ? String(n) : shownOk ? "✓" : "";
  label.textContent = n ? (n === 1 ? tr("Stelle zum Ansehen") : tr("Stellen zum Ansehen")) : "";
  // Mit Zahl und Legende steht der Satz nur noch für Vorleseprogramme da.
  verdict.classList.toggle("addin-quiet", n > 0);
  const box = document.getElementById("levels");
  const parts = [
    ["fehler", shownLevels.filter((l) => l === "fehler").length, (k) => tr("{0} Fehler", k)],
    ["pruefen", shownLevels.filter((l) => l === "pruefen").length, (k) => tr("{0} zum Prüfen", k)],
    ["info", shownLevels.filter((l) => l === "info").length, (k) => (k === 1 ? tr("1 Hinweis") : tr("{0} Hinweise", k))],
    ["done", inserted.length, (k) => tr("{0} eingesetzt", k)],
  ].filter(([, k]) => k > 0);
  box.hidden = !n;
  if (!n) {
    box.replaceChildren();
    return;
  }
  const track = el("div", "addin-track");
  const legend = el("div", "addin-legend");
  for (const [kind, k, text] of parts) {
    const seg = el("span", `addin-seg ${kind}`);
    seg.style.flex = `${k} 1 0`;
    track.append(seg);
    legend.append(el("span", kind, text(k)));
  }
  box.replaceChildren(track, legend);
}

function render(data) {
  const hints = data.hints || [];
  const verdict = document.getElementById("verdict");
  verdict.textContent = data.text || "";
  verdict.dataset.level = hints.some((h) => h.level === "fehler") ? "fehler"
    : hints.some((h) => h.level === "pruefen") ? "pruefen" : hints.length ? "info" : data.state === "ok" ? "ok" : "";
  shownLevels = hints.map((h) => h.level);
  shownOk = data.state === "ok";
  state.textContent = [data.cut, data.failed].filter(Boolean).join(" ");
  const tab = document.getElementById("tab-hints");
  tab.replaceChildren(tr("Hinweise"), ...(hints.length ? [el("span", "addin-count", String(hints.length))] : []));
  where.replaceChildren(...(data.book ? [el("span", "", data.book), el("span", "addin-sheetname", data.sheet)] : []));
  const items = [];
  let group = "";
  for (const h of hints) {
    if (h.level !== group) {
      group = h.level;
      const count = hints.filter((x) => x.level === group).length;
      const head = el("li", `addin-group ${group}`, GROUPS[group] || tr("Hinweise"));
      head.setAttribute("role", "presentation");
      head.append(el("span", "addin-count", String(count)));
      items.push(head);
    }
    const li = el("li", `addin-hint ${h.level}`);
    const first = cellLabel(h.cell);
    const skip = h.cell && !h.cell.includes(",") ? first : "";
    const fixed = h.fix ? fixTip(h.tip, h.fix) : null;
    if (h.title) {
      // Oben die Zelle und in einem Satz, was falsch ist; darunter warum, und was er gegenüberstellt.
      const head = el("div", "addin-hint-head");
      if (h.cell) head.append(cellChip(h.cell, h.level));
      const name = el("div", "addin-hint-name");
      name.append(el("span", "addin-title", h.title));
      if (h.where) name.append(el("span", "addin-place", h.where));
      head.append(name);
      li.append(head);
      const lead = el("p", "addin-text");
      // Die Zelle vorne im Satz bleibt stehen („D53 in der Sensitivität rechnet …“), sonst fehlt dem Satz der Anfang.
      lead.append(...linked(h.lead || h.text, data.sheet));
      li.append(lead);
      if (h.compare) li.append(compareBox(h.compare, h.fix));
    } else {
      const text = el("p", "addin-text");
      if (h.cell) text.append(cellChip(h.cell, h.level), " ");
      text.append(...linked(h.text, data.sheet, skip));
      li.append(text);
    }
    // Steht die Korrektur schon in der Gegenüberstellung, zeigt der Knopf sie nicht noch einmal.
    const shown = h.compare && h.fix && !h.fix.clear && (h.compare.kind === "formel"
      ? sameFormula(h.compare.want, shownFormula(h.fix))
      : h.fix.text && [`„${h.fix.german}“`, `“${h.fix.german}”`].includes(h.compare.want));
    if (h.tip && !fixed && !h.choices) li.append(el("p", "addin-tip", h.tip));
    if (h.fix) li.append(fixBar(data.sheet, h.fix, fixed, shown));
    if (h.choices) li.append(choiceBox(data.sheet, h.choices));
    if (fixed && fixed.rest) li.append(el("p", "addin-tip", fixed.rest));
    if (h.action) li.append(actionBar(data.sheet, h.action));
    if (h.id) {
      // Neben „Einsetzen“, wenn es eine Korrektur gibt, sonst in einer eigenen Zeile.
      const foot = h.fix && !h.choices ? li.querySelector(".addin-fix-foot") : null;
      if (foot) {
        foot.append(passtButton(data.sheet, h, li));
      } else {
        const row = el("div", "addin-hint-foot");
        row.append(passtButton(data.sheet, h, li));
        li.append(row);
      }
    }
    if (h.cell) clickable(li, tr("{0} in Excel markieren", h.cell.split(",").join(", ")), () => select(data.sheet, h.cell));
    items.push(li);
  }
  if (data.passt) items.push(passtNote(data.sheet, data.passt));
  document.getElementById("hints").replaceChildren(...items);
  allBar(data.sheet, hints.filter((h) => h.fix).length);
  renderDone();
  renderOthers(data.others || []);
  renderDrivers(data.drivers || []);
  const st = data.stats;
  document.getElementById("proof").textContent = st && st.formeln
    ? (st.formeln === 1 ? tr("1 Formel nachgerechnet") : tr("{0} Formeln nachgerechnet", st.formeln)) : "";
  document.getElementById("stats").textContent = st
    ? tr("{0} Formeln, {1} Eingaben, {2} Ergebnisse, Rechenkette bis {3} Stufen tief.", st.formeln, st.eingaben, st.ergebnisse, st.tiefe) +
      (st.nicht_nachrechenbar ? tr(" {0} Formeln kann ich noch nicht nachrechnen, dort zählt der Wert aus Excel.", st.nicht_nachrechenbar) : "") +
      (unread ? (unread === 1
        ? tr(" Die Mappe hat mehr als {0} Blätter: 1 Blatt habe ich nicht gelesen. Wechsle auf ein Blatt, dann prüfe ich es.", MAX_SHEETS)
        : tr(" Die Mappe hat mehr als {0} Blätter: {1} Blätter habe ich nicht gelesen. Wechsle auf ein Blatt, dann prüfe ich es.", MAX_SHEETS, unread)) : "")
    : "";
  document.getElementById("overview").replaceChildren(...(data.overview || []).map((o) => el("li", "", o.text)));
}

// --- Zwei Formeln oder zwei Werte nebeneinander ---

// Teile einer Formel: Bezüge, Funktionsnamen, Zahlen, Texte, Zeichen. Verglichen wird Teil für Teil.
const TOKEN = /\$?[A-Z]{1,3}\$?\d+|[A-ZÄÖÜ][A-ZÄÖÜ0-9.]*(?=\()|\d+(?:,\d+)?|"[^"]*"|„[^“]*“|\s+|[^\s]/g;

function tokens(formula) {
  return String(formula).match(TOKEN) || [];
}

function sameFormula(a, b) {
  return tokens(a).filter((t) => t.trim()).join("") === tokens(b).filter((t) => t.trim()).join("");
}

// Welche Teile beide Formeln gemeinsam haben (längste gemeinsame Folge): Der Rest ist, was sich unterscheidet.
function common(a, b) {
  const n = a.length, m = b.length;
  if (n * m > 40000) return [new Set(), new Set()];
  const table = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      table[i][j] = a[i] === b[j] ? table[i + 1][j + 1] + 1 : Math.max(table[i + 1][j], table[i][j + 1]);
    }
  }
  const keepA = new Set(), keepB = new Set();
  let i = 0, j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      keepA.add(i++);
      keepB.add(j++);
    } else if (table[i + 1][j] >= table[i][j + 1]) {
      i++;
    } else {
      j++;
    }
  }
  return [keepA, keepB];
}

function formulaLine(parts, keep, mark) {
  const code = el("code", "addin-cmp-formula");
  parts.forEach((part, i) => {
    if (keep.has(i) || !part.trim()) code.append(part);
    else code.append(el("mark", mark, part));
  });
  return code;
}

function compareRow(label, value, kind) {
  const row = el("div", `addin-cmp-row ${kind}`);
  row.append(el("span", "addin-cmp-label", label), value);
  return row;
}

function compareBox(compare, fix) {
  const box = el("div", "addin-compare");
  if (compare.kind === "formel") {
    const now = tokens(compare.now), want = tokens(compare.want);
    const [keepNow, keepWant] = common(now, want);
    box.append(compareRow(compare.now_label, formulaLine(now, keepNow, "addin-del"), "now"),
      compareRow(compare.want_label, formulaLine(want, keepWant, "addin-ins"), "want"));
  }
  const values = compare.kind === "wert" ? compare : compare.values;
  if (values) {
    if (compare.kind === "formel") box.append(el("div", "addin-cmp-gap"));
    box.append(compareRow(values.now_label, el("span", "addin-cmp-value", values.now), "now"),
      compareRow(values.want_label, el("span", "addin-cmp-value", values.want), "want"));
  }
  return box;
}

// --- Korrektur einsetzen und rückgängig machen ---

function cellsText(fix) {
  const n = Object.keys(fix.cells).length;
  return fix.range.includes(":") ? tr("{0} bis {1} ({2} Zellen)", ...fix.range.split(":"), n) : fix.range;
}

// Wie addin.sheets_from: die Formel ohne „=“ und Leerzeichen am Rand, "" für eine Zelle ohne Formel.
function formulaOf(f) {
  return typeof f === "string" && f.startsWith("=") ? f.slice(1).trim() : "";
}

// Steht in der Zelle noch, was beim Prüfen dort stand? Eine Formel (ohne „=“), bei Zahlen als Text der Text selbst.
function unchanged(was, formula, value) {
  if (was && typeof was === "object") return formulaOf(formula) === "" && String(value).trim() === String(was.text).trim();
  return formulaOf(formula) === (was || "");
}

function columnNumber(letters) {
  return [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

// Die Zellen einer Korrektur als Blöcke untereinander (C2:C9), damit Excel sie in wenigen Schritten liest und
// schreibt. Formeln und Zahlen kommen in getrennte Blöcke.
function blocksOf(fix) {
  const items = Object.keys(fix.cells).map((cell) => {
    const [, col, row] = /^([A-Z]+)(\d+)$/.exec(cell);
    return { cell, col, n: columnNumber(col), row: Number(row), number: typeof fix.cells[cell] === "number" };
  }).sort((a, b) => a.n - b.n || a.row - b.row);
  const blocks = [];
  for (const item of items) {
    const last = blocks[blocks.length - 1];
    if (last && last.col === item.col && last.end + 1 === item.row && last.number === item.number) {
      last.cells.push(item.cell);
      last.end = item.row;
    } else {
      blocks.push({ col: item.col, start: item.row, end: item.row, number: item.number, cells: [item.cell] });
    }
  }
  return blocks.map((b) => ({ ...b, address: b.cells.length === 1 ? b.cells[0] : `${b.col}${b.start}:${b.col}${b.end}` }));
}

// Die Korrektur wie in der Bearbeitungsleiste von Excel (fx), darunter der Knopf und wohin sie kommt.
function fixBar(sheet, fix, fixed = null, shown = false) {
  const bar = el("div", `addin-fix${shown ? " addin-fix-shown" : ""}`);
  const line = el("div", "addin-fx");
  if (fix.clear) {
    line.append(el("span", "addin-fx-mark", "fx"), el("span", "addin-fx-text", "leer"));
  } else {
    line.append(el("span", "addin-fx-mark", "fx"), el("code", fix.german.startsWith("=") ? "addin-formula" : "addin-fx-text", shownFormula(fix)));
  }
  const foot = el("div", "addin-fix-foot");
  const button = el("button", "addin-primary", tr("Einsetzen"));
  button.type = "button";
  button.title = fix.clear ? tr("Macht {0} leer", cellsText(fix)) : tr("Schreibt die Korrektur in {0}", cellsText(fix));
  const note = el("span", "sub small", fix.clear ? tr("leert {0}", cellsText(fix)) : tr("in {0}", cellsText(fix)));
  if (fixed && fixed.value) note.append(tr(", ergibt "), el("span", "addin-num", fixed.value));
  // Ein Klick auf den Knopf markiert nicht auch die Zelle des Hinweises.
  button.addEventListener("click", (e) => {
    e.stopPropagation();
    insert(sheet, fix, button, note);
  });
  button.addEventListener("keydown", (e) => e.stopPropagation());
  foot.append(button, note);
  if (shown) bar.append(foot);
  else bar.append(line, foot);
  return bar;
}

// Zwei Korrekturen zur Wahl (ein Code und seine Bezeichnung widersprechen sich): Welche stimmt, weiß nur, wer die
// Akte kennt. Jede hat ihren eigenen Knopf; „Alle einsetzen“ nimmt keine davon.
function choiceBox(sheet, choices) {
  const box = el("div", "addin-choices");
  box.append(el("p", "addin-choices-head", "Eins davon einsetzen:"));
  for (const choice of choices) box.append(fixBar(sheet, choice.fix));
  return box;
}

// „Alle einsetzen“: Gegenprobe setzt die Korrekturen zusammen ein, rechnet nach und lässt weg, was sich stört.
// Was es dazu sagt, bleibt stehen, bis wieder etwas eingesetzt oder zurückgenommen wird.
let allSaid = "";

function allBar(sheet, count) {
  const box = document.getElementById("all");
  box.hidden = count < 2 && !allSaid;
  if (count < 2) {
    box.replaceChildren(...(allSaid ? [el("span", "sub small", allSaid)] : []));
    return;
  }
  const button = el("button", "addin-primary", tr("Alle {0} einsetzen", count));
  button.type = "button";
  const note = el("span", "sub small", allSaid || tr("setzt die Korrekturen zusammen ein, nachgerechnet; ein Rückgängig nimmt alle zurück"));
  button.addEventListener("click", () => insertAll(sheet, button, note));
  box.replaceChildren(button, note);
}

async function insertAll(sheet, button, note) {
  button.disabled = true;
  note.textContent = tr("Ich rechne alle zusammen nach …");
  let data = null;
  try {
    data = await post({ sheet }, "/api/addin/alle");
  } catch (err) {
    note.textContent = silent(err);
    button.disabled = false;
    return;
  }
  if (!data || data.state !== "ok") {
    note.textContent = (data && data.text) || tr("Das ging nicht. Setz die Korrekturen einzeln ein.");
    if (data && data.state === "again") later();
    button.disabled = false;
    return;
  }
  const done = await insert(sheet, data.fix, button, note);
  if (done && data.left) {
    // Nach dem neuen Prüfen steht unten, was davon noch nötig ist (oft hat sich ein Teil mit erledigt).
    allSaid = note.textContent + (data.left === 1
      ? tr(" Eine Korrektur habe ich weggelassen, weil sie nicht mit den anderen zusammenpasst. Was noch fehlt, steht unten.")
      : tr(" {0} Korrekturen habe ich weggelassen, weil sie nicht mit den anderen zusammenpassen. Was noch fehlt, steht unten.", data.left));
    note.textContent = allSaid;
  }
}

// Dasselbe Zahlenformat mit zwei Nachkommastellen (#,##0 → #,##0.00), damit 0,4 nicht als 0 dasteht.
function withDecimals(format) {
  const parts = String(format || "").split(";").map((part) => {
    const at = part.replace(/"[^"]*"/g, (m) => " ".repeat(m.length)).lastIndexOf("0");
    return at < 0 ? part : `${part.slice(0, at + 1)}.00${part.slice(at + 1)}`;
  });
  return parts.some((p, i) => p !== String(format || "").split(";")[i]) ? parts.join(";") : "#,##0.00";
}

// Das Zahlenformat, das eine Zelle nach dem Einsetzen bekommt.
function formatAfter(fix, cell, likes, before) {
  if (fix.format && fix.format[cell]) return likes.get(fix.format[cell]).numberFormat[0][0];
  if ((fix.decimals || []).includes(cell)) return withDecimals(before[cell].format);
  return (fix.general || []).includes(cell) ? "General" : before[cell].format;
}

async function insert(sheet, fix, button, note) {
  button.disabled = true;
  allSaid = "";
  let changed = "";
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(sheet);
      const blocks = blocksOf(fix).map((b) => ({ ...b, range: ws.getRange(b.address) }));
      for (const b of blocks) b.range.load("formulas,values,numberFormat");
      // Zahlen statt Text bekommen das Zahlenformat einer echten Zahl derselben Spalte.
      const likes = new Map();
      for (const like of new Set(Object.values(fix.format || {}))) {
        const range = ws.getRange(like);
        range.load("numberFormat");
        likes.set(like, range);
      }
      await context.sync();
      const before = {};
      for (const b of blocks) {
        b.cells.forEach((cell, i) => {
          before[cell] = { formula: b.range.formulas[i][0], value: b.range.values[i][0], format: b.range.numberFormat[i][0] };
        });
      }
      // Seit dem Prüfen geändert: lieber neu prüfen, als etwas zu überschreiben, das gerade getippt wurde.
      changed = Object.keys(fix.cells).find((cell) => !unchanged(fix.was[cell], before[cell].formula, before[cell].value)) || "";
      if (changed) return;
      // Verbundene Zellen: erst den Verbund aufheben, dann steht der Wert in jeder Zeile (Rückgängig verbindet wieder).
      for (const area of fix.unmerge || []) ws.getRange(area).unmerge();
      // Eine Formel, die als Text dastand: erst das Zahlenformat „Standard“, sonst bliebe sie Text.
      for (const cell of fix.general || []) ws.getRange(cell).numberFormat = [["General"]];
      for (const b of blocks) {
        if ((fix.format || fix.decimals) && b.cells.some((cell) => (fix.format || {})[cell] || (fix.decimals || []).includes(cell))) {
          b.range.numberFormat = b.cells.map((cell) => [formatAfter(fix, cell, likes, before)]);
        }
        const put = b.cells.map((cell) => [fix.cells[cell]]);
        if (b.number) b.range.values = put;
        else b.range.formulas = put;
      }
      await context.sync();
      // Excel gibt einer Formel mit Uhrzeiten von selbst ein Uhrzeit-Format (8 Stunden stünden dann als 00:00 da), und
      // das erst, wenn die Formel steht. Darum danach: Die Zelle behält ihr Format, das von vorher, oder das einer
      // Nachbarzelle, wo der Hinweis das sagt (Zahlen, die Text waren, eine Zahl im Uhrzeitformat).
      for (const b of blocks) {
        b.range.numberFormat = b.cells.map((cell) => [formatAfter(fix, cell, likes, before)]);
        // So, wie Excel die Zellen jetzt zeigt: Daran sieht Rückgängig, ob seither jemand etwas geändert hat.
        b.range.load("formulas");
      }
      await context.sync();
      const after = {};
      for (const b of blocks) b.cells.forEach((cell, i) => { after[cell] = b.range.formulas[i][0]; });
      inserted.unshift({ sheet, fix, before, after });
      inserted.splice(MAX_UNDO);
    });
  } catch (err) {
    note.textContent = tr("Einsetzen ging nicht. {0} Du kannst die Korrektur oben auch selbst eintippen.", excelProblem(err));
    button.disabled = false;
    return false;
  }
  if (changed) {
    note.textContent = tr("In {0} steht inzwischen etwas anderes. Ich prüfe neu, dann kannst du es noch einmal versuchen.", changed);
    later();
    return false;
  }
  note.textContent = fix.clear ? tr("Geleert: {0}.", cellsText(fix)) : tr("Eingesetzt in {0}.", cellsText(fix));
  renderDone();
  // Die Änderung löst in Excel ohnehin ein neues Prüfen aus; so auch, wenn Excel sie nicht meldet.
  later();
  return true;
}

async function undo(entry, button, note) {
  button.disabled = true;
  allSaid = "";
  let changed = "";
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(entry.sheet);
      const blocks = blocksOf(entry.fix).map((b) => ({ ...b, range: ws.getRange(b.address) }));
      for (const b of blocks) b.range.load("formulas");
      await context.sync();
      changed = blocks.flatMap((b) => b.cells.filter((cell, i) => String(b.range.formulas[i][0]) !== String(entry.after[cell])))[0] || "";
      if (changed) return;
      for (const b of blocks) {
        const was = b.cells.map((cell) => entry.before[cell]);
        if (b.number) {
          // Es war Text: als Text zurück. Erst im Textformat schreiben, dann das alte Format; Excel lässt den Text so.
          b.range.numberFormat = was.map(() => ["@"]);
          b.range.values = was.map((w) => [w.value]);
          b.range.numberFormat = was.map((w) => [w.format]);
        } else {
          b.range.formulas = was.map((w) => [w.formula]);
        }
      }
      for (const area of entry.fix.unmerge || []) ws.getRange(area).merge(false);
      await context.sync();
      // Erst wenn die alte Formel steht, das alte Format: Sonst gäbe Excel ihr wieder ein eigenes.
      for (const b of blocks) if (!b.number) b.range.numberFormat = b.cells.map((cell) => [entry.before[cell].format]);
      await context.sync();
    });
  } catch (err) {
    note.textContent = tr("Rückgängig ging nicht. {0}", excelProblem(err));
    button.disabled = false;
    return;
  }
  if (changed) {
    note.textContent = tr("In {0} hat sich seit dem Einsetzen etwas geändert, darum lasse ich es so.", changed);
    return;
  }
  inserted.splice(inserted.indexOf(entry), 1);
  renderDone();
  state.textContent = tr("Rückgängig gemacht: In {0} steht wieder, was vorher dort war.", cellsText(entry.fix));
  later();
}

// Nur die letzte Korrektur steht offen da; die früheren liegen eingeklappt darunter, sonst füllen sie das Fenster.
let doneOpen = false;

function renderDone() {
  const box = document.getElementById("done");
  box.hidden = !inserted.length;
  renderLevels();
  const rows = inserted.map((entry) => {
    const row = el("div", "addin-done-row");
    const elsewhere = entry.sheet !== active ? tr(" auf Blatt {0}", entry.sheet) : "";
    const button = el("button", "", tr("Rückgängig"));
    button.type = "button";
    if (entry.action) {
      const note = el("span", "small", `${entry.said}${elsewhere}`);
      button.addEventListener("click", () => undoAction(entry, button, note));
      row.append(note, button);
      return row;
    }
    const note = el("span", "small", entry.fix.clear ? tr("Geleert: {0}{1}", cellsText(entry.fix), elsewhere)
      : tr("Eingesetzt in {0}{1}: ", cellsText(entry.fix), elsewhere));
    if (!entry.fix.clear) note.append(el("span", entry.fix.german.startsWith("=") ? "addin-formula" : "", shownFormula(entry.fix)));
    // Bei „Alle einsetzen“: welche Korrekturen es waren.
    if (entry.fix.parts) note.title = entry.fix.parts.join("\n");
    button.addEventListener("click", () => undo(entry, button, note));
    row.append(note, button);
    return row;
  });
  if (rows.length < 2) {
    box.replaceChildren(...rows);
    return;
  }
  const more = document.createElement("details");
  more.className = "addin-done-more";
  more.open = doneOpen;
  more.addEventListener("toggle", () => { doneOpen = more.open; });
  const n = rows.length - 1;
  more.append(el("summary", "small", n === 1 ? tr("1 frühere Korrektur") : tr("{0} frühere Korrekturen", n)), ...rows.slice(1));
  box.replaceChildren(rows[0], more);
}

// --- Pivot-Tabellen und Diagramme: aktualisieren, Summe statt Anzahl, Bereiche einer Reihe setzen ---

// Ein Bereich, wie Excel ihn nennt, ohne „=“, $ und Anführungszeichen: So sieht man, ob er sich seither geändert hat.
function plainRef(text) {
  return String(text || "").replace(/^=/, "").replace(/[$']/g, "").trim().toLowerCase();
}

const ACTIONS = {
  refresh: { label: tr("Aktualisieren"), note: (a) => tr("aktualisiert die Pivot-Tabelle „{0}“ aus ihren Daten", a.pivot) },
  sum: { label: tr("Summe nehmen"), note: (a) => tr("fasst „{0}“ als Summe zusammen", a.field) },
  filter: { label: tr("Einsetzen"), note: (a) => tr("stellt den Berichtsfilter {0} auf „{1}“", a.field, a.value) },
  source: {
    label: tr("Einsetzen"),
    note: (a) => (a.table
      ? tr("erweitert die Tabelle {0} auf {1} und baut die Pivot-Tabelle „{2}“ damit neu auf, mit denselben Feldern", a.table, a.source, a.pivot)
      : tr("baut die Pivot-Tabelle „{0}“ mit den Daten aus {1} neu auf, mit denselben Feldern", a.pivot, a.source)),
  },
  name: {
    label: tr("Einsetzen"),
    note: (a) => tr("lässt den Namen {0} auf {1} zeigen; die Formeln damit bleiben, wie sie sind", a.name, a.to.replace(/^=/, "").replace(/\$/g, "")),
  },
  table: {
    label: tr("Einsetzen"),
    note: (a) => tr("erweitert die Tabelle {0} auf {1} und aktualisiert die Pivot-Tabelle „{2}“", a.table, a.source, a.pivot),
  },
  series: {
    label: tr("Einsetzen"),
    note: (a) => tr("setzt im Diagramm „{0}“ {1}", a.chart, a.series.map((c) => c.values + (c.categories ? tr(" mit {0}", c.categories) : "")).join(", ")),
  },
};

function actionBar(sheet, action) {
  const kind = ACTIONS[action.kind];
  const bar = el("div", "addin-fix addin-act");
  if (action.kind === "rule") {
    // Ein Verstoß gegen eine eigene Regel: die Regel gleich ändern können, wenn der Wert doch stimmt.
    const button = el("button", "", tr("Regel ändern"));
    button.type = "button";
    button.addEventListener("click", (e) => {
      e.stopPropagation();
      openRule(action.spalte, action.blatt || "");
    });
    button.addEventListener("keydown", (e) => e.stopPropagation());
    bar.append(button);
    return bar;
  }
  if (!kind) return bar;
  const button = el("button", "", kind.label);
  button.type = "button";
  const note = el("span", "sub small", kind.note(action));
  button.addEventListener("click", (e) => {
    e.stopPropagation();
    act(sheet, action, button, note);
  });
  button.addEventListener("keydown", (e) => e.stopPropagation());
  bar.append(button, note);
  return bar;
}

function dimensions(item, change) {
  return {
    values: item.getDimensionDataSourceString(Excel.ChartSeriesDimension.values),
    categories: change.categories ? item.getDimensionDataSourceString(Excel.ChartSeriesDimension.categories) : null,
  };
}

// Welche Einträge eines Felds die Pivot-Tabelle zeigt; alle: Es ist nichts ausgeblendet.
async function shownItems(context, field) {
  field.items.load("items/name,items/visible");
  await context.sync();
  const shown = field.items.items.filter((i) => i.visible).map((i) => i.name);
  return { shown, all: shown.length === field.items.items.length };
}

function sameItems(a, b) {
  const norm = (list) => list.map((x) => String(x).trim().toLowerCase()).sort().join("\u0001");
  return norm(a) === norm(b);
}

function filterField(ws, action) {
  return ws.pivotTables.getItem(action.pivot).filterHierarchies.getItem(action.field).fields.getItem(action.field);
}

// Excel kennt keinen Befehl, die Quelle einer Pivot-Tabelle zu ändern. Darum baut Einsetzen sie an derselben Stelle
// mit denselben Zeilen, Spalten, Werten und Berichtsfiltern neu auf. Gibt false zurück, wenn ihre Quelle inzwischen
// nicht mehr eine aus „expect“ ist. table: die Tabelle wird auf address gebracht und ist die neue Quelle.
async function rebuildPivot(context, ws, name, sheet, address, expect, table) {
  const pt = ws.pivotTables.getItem(name);
  const groups = { rows: pt.rowHierarchies, columns: pt.columnHierarchies, filters: pt.filterHierarchies };
  for (const group of Object.values(groups)) group.load("items/name");
  pt.dataHierarchies.load("items/name,items/summarizeBy");
  const layout = pt.layout;
  layout.load("layoutType,showRowGrandTotals,showColumnGrandTotals");
  const place = layout.getRange().getCell(0, 0);
  place.load("address");
  const source = supports("1.15") ? pt.getDataSourceString() : null;
  await context.sync();
  if (source && !expect.map(plainRef).includes(plainRef(source.value).split("!").pop())) return false;
  const was = expect[0];
  for (const group of Object.values(groups)) for (const h of group.items) h.fields.load("items/name");
  for (const d of pt.dataHierarchies.items) d.field.load("name");
  await context.sync();
  const names = {};
  for (const [key, group] of Object.entries(groups)) {
    names[key] = group.items.map((h) => (h.fields.items[0] ? h.fields.items[0].name : h.name));
  }
  const values = pt.dataHierarchies.items.map((d) => ({ field: d.field.name, name: d.name, by: d.summarizeBy }));
  // Was die Berichtsfilter gerade zeigen, kommt mit.
  const chosen = {};
  for (const h of pt.filterHierarchies.items) {
    const field = h.fields.items[0];
    if (!field) continue;
    const items = await shownItems(context, field);
    if (!items.all) chosen[field.name] = items.shown;
  }
  const settings = {
    layoutType: layout.layoutType,
    showRowGrandTotals: layout.showRowGrandTotals,
    showColumnGrandTotals: layout.showColumnGrandTotals,
  };
  const target = place.address.split("!").pop();
  const build = async (from, asTable) => {
    const data = asTable ? context.workbook.tables.getItem(asTable) : context.workbook.worksheets.getItem(sheet).getRange(from);
    const fresh = ws.pivotTables.add(name, data, ws.getRange(target));
    for (const n of names.rows) fresh.rowHierarchies.add(fresh.hierarchies.getItem(n));
    for (const n of names.columns) fresh.columnHierarchies.add(fresh.hierarchies.getItem(n));
    for (const n of names.filters) fresh.filterHierarchies.add(fresh.hierarchies.getItem(n));
    for (const v of values) {
      const d = fresh.dataHierarchies.add(fresh.hierarchies.getItem(v.field));
      d.summarizeBy = v.by;
      d.name = v.name;
    }
    fresh.layout.set(settings);
    await context.sync();
    if (supports("1.12")) {
      for (const [n, shown] of Object.entries(chosen)) {
        fresh.filterHierarchies.getItem(n).fields.getItem(n).applyFilter({ manualFilter: { selectedItems: shown } });
      }
      await context.sync();
    }
  };
  if (table) {
    const t = context.workbook.tables.getItem(table);
    const now = t.getRange();
    now.load("address");
    await context.sync();
    // Die Tabelle auf den neuen Bereich (beim Rückgängig wieder auf den alten).
    if (plainRef(now.address).split("!").pop() !== plainRef(address)) {
      t.resize(context.workbook.worksheets.getItem(sheet).getRange(address));
      await context.sync();
    }
  }
  pt.delete();
  await context.sync();
  try {
    await build(address, table);
  } catch (err) {
    // Ging der Neuaufbau schief: die Pivot-Tabelle mit der alten Quelle wieder herstellen, dann den Fehler melden.
    try {
      const half = ws.pivotTables.getItemOrNullObject(name);
      await context.sync();
      if (!half.isNullObject) {
        half.delete();
        await context.sync();
      }
      await build(was, null);
    } catch (_) {}
    throw err;
  }
  return true;
}

async function act(sheet, action, button, note) {
  button.disabled = true;
  allSaid = "";
  let said = "";
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(sheet);
      if (action.kind === "refresh") {
        ws.pivotTables.getItem(action.pivot).refresh();
        await context.sync();
        said = tr("Die Pivot-Tabelle „{0}“ ist aktualisiert.", action.pivot);
      } else if (action.kind === "sum") {
        const field = ws.pivotTables.getItem(action.pivot).dataHierarchies.getItem(action.field);
        field.load("summarizeBy");
        await context.sync();
        if (field.summarizeBy !== "Count") {
          said = tr("„{0}“ zählt inzwischen nicht mehr. Ich prüfe neu.", action.field);
          return;
        }
        field.summarizeBy = "Sum";
        await context.sync();
        field.load("name");
        await context.sync();
        inserted.unshift({ sheet, action, name: field.name, said: tr("Summe statt Anzahl in „{0}“", field.name) });
        inserted.splice(MAX_UNDO);
        said = tr("„{0}“ summiert jetzt.", field.name);
      } else if (action.kind === "filter") {
        if (!supports("1.12")) {
          said = tr("Diese Excel-Version kann den Filter nicht von hier aus stellen. Der Tipp oben sagt, wie es von Hand geht.");
          return;
        }
        const field = filterField(ws, action);
        const now = await shownItems(context, field);
        if (action.was.length ? !sameItems(now.shown, action.was) : !now.all) {
          said = tr("Der Filter {0} hat sich inzwischen geändert. Ich prüfe neu.", action.field);
          return;
        }
        field.applyFilter({ manualFilter: { selectedItems: [action.value] } });
        await context.sync();
        inserted.unshift({ sheet, action, said: tr("Berichtsfilter {0} auf „{1}“", action.field, action.value) });
        inserted.splice(MAX_UNDO);
        said = tr("Die Pivot-Tabelle „{0}“ zeigt jetzt nur {1} = „{2}“.", action.pivot, action.field, action.value);
      } else if (action.kind === "source") {
        if (action.table && !supports("1.13")) {
          said = tr("Diese Excel-Version kann die Tabelle nicht von hier aus erweitern. Der Tipp oben sagt, wie es von Hand geht.");
          return;
        }
        if (!(await rebuildPivot(context, ws, action.pivot, action.sheet, action.source, [action.was], action.table))) {
          said = tr("Die Quelle der Pivot-Tabelle „{0}“ hat sich inzwischen geändert. Ich prüfe neu.", action.pivot);
          return;
        }
        inserted.unshift({ sheet, action, said: tr("Pivot-Tabelle „{0}“ mit den Daten aus {1}", action.pivot, action.source) });
        inserted.splice(MAX_UNDO);
        said = tr("Die Pivot-Tabelle „{0}“ nimmt jetzt die Daten aus {1}.", action.pivot, action.source);
      } else if (action.kind === "name") {
        const item = context.workbook.names.getItem(action.name);
        item.load("formula");
        await context.sync();
        if (plainRef(item.formula) !== plainRef(action.was)) {
          said = tr("Der Name {0} hat sich inzwischen geändert. Ich prüfe neu.", action.name);
          return;
        }
        item.formula = action.to;
        await context.sync();
        inserted.unshift({ sheet, action, said: tr("Name {0} auf {1}", action.name, action.to.replace(/^=/, "")) });
        inserted.splice(MAX_UNDO);
        said = tr("Der Name {0} zeigt jetzt auf {1}.", action.name, action.to.replace(/^=/, ""));
      } else if (action.kind === "table") {
        if (!supports("1.13")) {
          said = tr("Diese Excel-Version kann die Tabelle nicht von hier aus erweitern. Der Tipp oben sagt, wie es von Hand geht.");
          return;
        }
        const table = context.workbook.tables.getItem(action.table);
        const now = table.getRange();
        now.load("address");
        await context.sync();
        if (plainRef(now.address).split("!").pop() !== plainRef(action.was)) {
          said = tr("Die Tabelle {0} hat sich inzwischen geändert. Ich prüfe neu.", action.table);
          return;
        }
        table.resize(context.workbook.worksheets.getItem(action.sheet).getRange(action.source));
        ws.pivotTables.getItem(action.pivot).refresh();
        await context.sync();
        inserted.unshift({ sheet, action, said: tr("Tabelle {0} auf {1}", action.table, action.source) });
        inserted.splice(MAX_UNDO);
        said = tr("Die Tabelle {0} geht jetzt bis {1}, die Pivot-Tabelle „{2}“ ist aktualisiert.", action.table, action.source, action.pivot);
      } else if (action.kind === "series") {
        const chart = ws.charts.getItem(action.chart);
        const items = action.series.map((change) => {
          const item = chart.series.getItemAt(change.index);
          return { change, item, now: dimensions(item, change) };
        });
        await context.sync();
        // Seit dem Prüfen geändert: lieber neu prüfen, als etwas zu überschreiben.
        const moved = items.some(({ change, now }) => plainRef(now.values.value) !== plainRef(change.was.values) ||
          (now.categories && plainRef(now.categories.value) !== plainRef(change.was.categories)));
        if (moved) {
          said = tr("Am Diagramm „{0}“ hat sich inzwischen etwas geändert. Ich prüfe neu, dann kannst du es noch einmal versuchen.", action.chart);
          return;
        }
        for (const { change, item } of items) {
          const data = context.workbook.worksheets.getItem(change.sheet);
          item.setValues(data.getRange(change.values));
          if (change.categories) item.setXAxisValues(data.getRange(change.categories));
        }
        await context.sync();
        // So, wie Excel die Bereiche jetzt nennt: Daran sieht Rückgängig, ob seither jemand etwas geändert hat.
        const after = items.map(({ change, item }) => dimensions(item, change));
        await context.sync();
        inserted.unshift({
          sheet, action,
          after: after.map((a) => ({ values: a.values.value, categories: a.categories ? a.categories.value : "" })),
          said: tr("Diagramm „{0}“: {1}", action.chart, action.series.map((c) => c.values).join(", ")),
        });
        inserted.splice(MAX_UNDO);
        said = tr("Eingesetzt im Diagramm „{0}“.", action.chart);
      }
    });
  } catch (err) {
    note.textContent = tr("Das ging nicht. {0} Der Tipp oben sagt, wie es in Excel von Hand geht.", excelProblem(err));
    button.disabled = false;
    return;
  }
  note.textContent = said;
  renderDone();
  later();
}

async function undoAction(entry, button, note) {
  button.disabled = true;
  allSaid = "";
  let changed = false;
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(entry.sheet);
      if (entry.action.kind === "sum") {
        const field = ws.pivotTables.getItem(entry.action.pivot).dataHierarchies.getItem(entry.name);
        field.load("summarizeBy");
        await context.sync();
        changed = field.summarizeBy !== "Sum";
        if (changed) return;
        field.summarizeBy = "Count";
        await context.sync();
        return;
      }
      if (entry.action.kind === "filter") {
        const field = filterField(ws, entry.action);
        const now = await shownItems(context, field);
        changed = !sameItems(now.shown, [entry.action.value]);
        if (changed) return;
        if (entry.action.was.length) field.applyFilter({ manualFilter: { selectedItems: entry.action.was } });
        else field.clearAllFilters();
        await context.sync();
        return;
      }
      if (entry.action.kind === "source") {
        const a = entry.action;
        changed = !(await rebuildPivot(context, ws, a.pivot, a.sheet, a.was, [a.source, a.table || ""], null));
        // Hatte die Pivot-Tabelle vorher eine feste Quelle über der Tabelle: die Tabelle wieder so lang wie vorher.
        if (!changed && a.table) {
          context.workbook.tables.getItem(a.table).resize(context.workbook.worksheets.getItem(a.sheet).getRange(a.was));
          await context.sync();
        }
        return;
      }
      if (entry.action.kind === "name") {
        const item = context.workbook.names.getItem(entry.action.name);
        item.load("formula");
        await context.sync();
        changed = plainRef(item.formula) !== plainRef(entry.action.to);
        if (changed) return;
        item.formula = entry.action.was;
        await context.sync();
        return;
      }
      if (entry.action.kind === "table") {
        const a = entry.action;
        const table = context.workbook.tables.getItem(a.table);
        const now = table.getRange();
        now.load("address");
        await context.sync();
        changed = plainRef(now.address).split("!").pop() !== plainRef(a.source);
        if (changed) return;
        table.resize(context.workbook.worksheets.getItem(a.sheet).getRange(a.was));
        ws.pivotTables.getItem(a.pivot).refresh();
        await context.sync();
        return;
      }
      const chart = ws.charts.getItem(entry.action.chart);
      const items = entry.action.series.map((change) => {
        const item = chart.series.getItemAt(change.index);
        return { change, item, now: dimensions(item, change) };
      });
      await context.sync();
      changed = items.some(({ now }, i) => plainRef(now.values.value) !== plainRef(entry.after[i].values) ||
        (now.categories && plainRef(now.categories.value) !== plainRef(entry.after[i].categories)));
      if (changed) return;
      for (const { change, item } of items) {
        const data = context.workbook.worksheets.getItem(change.sheet);
        item.setValues(data.getRange(change.undo.values));
        if (change.categories && change.undo.categories) item.setXAxisValues(data.getRange(change.undo.categories));
      }
      await context.sync();
    });
  } catch (err) {
    note.textContent = tr("Rückgängig ging nicht. {0}", excelProblem(err));
    button.disabled = false;
    return;
  }
  if (changed) {
    note.textContent = tr("Seit dem Einsetzen hat sich dort etwas geändert, darum lasse ich es so.");
    return;
  }
  inserted.splice(inserted.indexOf(entry), 1);
  renderDone();
  state.textContent = tr("Rückgängig gemacht: {0}.", entry.said);
  later();
}

// Mit Maus und Tastatur (Enter, Leertaste) bedienbar.
function clickable(node, title, go) {
  node.tabIndex = 0;
  node.title = title;
  node.addEventListener("click", go);
  node.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      go();
    }
  });
}

async function show(sheet) {
  if (hiddenNote(sheet)) return;
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItem(sheet).activate();
      await context.sync();
    });
  } catch (_) {
    state.textContent = tr("Das Blatt {0} kann ich nicht öffnen.", sheet);
  }
}

function renderOthers(others) {
  const shown = others.filter((o) => o.fehler === null || o.fehler + o.pruefen > 0);
  document.getElementById("others-box").hidden = !shown.length;
  document.getElementById("others").replaceChildren(...shown.map((o) => {
    const li = el("li", "addin-other");
    const parts = o.fehler === null ? [tr("wird noch geprüft")] : [o.fehler ? tr("{0} Fehler", o.fehler) : "", o.pruefen ? tr("{0} zum Prüfen", o.pruefen) : ""];
    li.append(el("span", "", o.hidden ? `${o.name} (ausgeblendet)` : o.name), el("span", `sub small${o.fehler ? " addin-bad" : ""}`, parts.filter(Boolean).join(", ")));
    clickable(li, tr("Blatt {0} öffnen", o.name), () => show(o.name));
    return li;
  }));
  // Große Blätter prüft Gegenprobe im Hintergrund: ab und zu nachfragen, ohne neu zu lesen.
  clearTimeout(othersTimer);
  if (others.some((o) => o.fehler === null)) othersTimer = setTimeout(askOthers, OTHERS_MS);
}

async function askOthers() {
  if (busy) {
    othersTimer = setTimeout(askOthers, OTHERS_MS);
    return;
  }
  try {
    const data = await post({}, "/api/addin/andere");
    if (data.book === book && data.sheet === active) renderOthers(data.others || []);
  } catch (_) {
    // Antwortet er gerade nicht, kommt es mit der nächsten Prüfung.
  }
}

function percent(value) {
  const text = Math.abs(value).toLocaleString(LANG === "en" ? "en-US" : "de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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
    const card = el("div", "addin-driver");
    card.append(el("strong", "", `${output} = ${g.value}`));
    const list = el("ul", "small");
    for (const d of g.items) list.append(el("li", "", tr("{0} ({1}): 1 % mehr ergibt {2}", d.input, d.input_value, percent(d.elasticity))));
    card.append(list);
    return card;
  }));
}

// --- Herleitung der markierten Zelle ---

function showTab(name) {
  for (const tab of ["hints", "why", "book", "file"]) {
    document.getElementById(`tab-${tab}`).setAttribute("aria-selected", String(tab === name));
    document.getElementById(`page-${tab}`).hidden = tab !== name;
  }
}

async function selectedPlace() {
  let address = "";
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    address = range.address;
  });
  const at = address.lastIndexOf("!");
  let sheet = address.slice(0, at);
  if (sheet.startsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, cell: address.slice(at + 1).split(":")[0].replace(/\$/g, "") };
}

function onSelection() {
  clearTimeout(selectTimer);
  selectTimer = setTimeout(async () => {
    let place;
    try {
      place = await selectedPlace();
    } catch (_) {
      return;
    }
    if (`${place.sheet}!${place.cell}` === ownSelect) {
      ownSelect = "";
      return;
    }
    explain(place, false);
  }, 250);
}

async function explain(place, again) {
  let data;
  try {
    data = await post(place, "/api/addin/herleitung");
  } catch (_) {
    return;
  }
  if (data.state === "again") {
    // Gegenprobe kennt die Mappe noch nicht (gerade gestartet): einmal prüfen, dann noch einmal.
    if (!again) {
      shownPlace = place;
      check(false);
    }
    return;
  }
  shownPlace = data.state === "ok" ? place : null;
  renderWhy(data);
}

function nameOf(label) {
  const match = /\((.*)\)$/.exec(label || "");
  return match ? match[1] : "";
}

function hasProblem(node) {
  return node.check === "anders" || (node.children || []).some(hasProblem);
}

function verdict(node) {
  if (node.kind !== "formel") return node.kind === "leer" ? tr("Die Zelle ist leer.") : tr("Eingabe: {0}", node.value);
  if (node.check === "ok") return tr("Excel zeigt {0}. Nachgerechnet stimmt es.", node.value);
  if (node.check === "anders") return tr("Excel zeigt {0}, nachgerechnet ergibt die Formel aber {1}.", node.value, node.ours);
  return tr("Excel zeigt {0}. Diese Formel kann ich noch nicht nachrechnen.", node.value);
}

function mark(sheet, cell) {
  ownSelect = `${sheet}!${cell.split(":")[0]}`;
  select(sheet, cell);
}

function row(node, home) {
  const line = el("div", `addin-row${node.check === "anders" ? " bad" : ""}`);
  const where = node.sheet !== home ? `${node.sheet}!${node.cell}` : node.cell;
  line.append(el("span", "addin-cell", where));
  const name = nameOf(node.label);
  if (name) line.append(el("span", "addin-name", name));
  line.append(el("span", "addin-value", node.value || ""));
  if (node.check === "ok") line.append(el("span", "addin-good", "✓"));
  if (node.check === "anders") line.append(el("span", "addin-bad", `≠ ${node.ours}`));
  if (node.kind !== "fremd") clickable(line, tr("{0} in Excel markieren", where), () => mark(node.sheet, node.cell));
  return line;
}

function tree(items, home, depth) {
  const list = el("ul", "addin-tree");
  for (const node of items) {
    const item = el("li");
    const kids = node.children && node.children.length ? tree(node.children, home, depth + 1) : null;
    const head = el("div", "addin-node");
    if (kids) {
      // Offen: die erste Stufe, der Weg zu einer Zahl, die nicht stimmt, und eine Tabelle mit ihrer Fundstelle.
      const open = depth < 1 || hasProblem(node) || (node.kind === "bereich" && node.children.length <= 2);
      const toggle = el("button", "addin-toggle", open ? "▾" : "▸");
      toggle.type = "button";
      toggle.setAttribute("aria-expanded", String(open));
      toggle.setAttribute("aria-label", tr("{0} auf- oder zuklappen", node.cell));
      kids.hidden = !open;
      toggle.addEventListener("click", () => {
        kids.hidden = !kids.hidden;
        toggle.textContent = kids.hidden ? "▸" : "▾";
        toggle.setAttribute("aria-expanded", String(!kids.hidden));
      });
      head.append(toggle);
    } else {
      head.append(el("span", "addin-toggle"));
    }
    head.append(row(node, home));
    item.append(head);
    const notes = [];
    if (node.note) notes.push(node.note);
    if (node.kind === "formel") notes.push(node.words || node.formula);
    if (node.again) notes.push(tr("kommt weiter oben schon vor"));
    if (node.more) notes.push(tr("und {0} weitere", node.more));
    for (const text of notes) item.append(el("div", "addin-note sub small", text));
    if (kids) item.append(kids);
    list.append(item);
  }
  return list;
}

function renderWhy(data) {
  const box = document.getElementById("why");
  const empty = document.getElementById("why-empty");
  const tab = document.getElementById("tab-why");
  if (data.state !== "ok") {
    box.replaceChildren();
    empty.hidden = false;
    tab.textContent = tr("Herleitung");
    return;
  }
  const t = data.tree;
  empty.hidden = true;
  tab.textContent = tr("Herleitung {0}", data.cell);
  const parts = [el("h2", "addin-why-title", t.label)];
  if (t.kind === "formel") {
    parts.push(el("p", "addin-formula", t.formula));
    if (t.words && t.words !== t.formula) parts.push(el("p", "sub small", t.words));
    if (t.rows) parts.push(el("p", "addin-rows", t.rows));
  }
  parts.push(el("p", t.check === "anders" ? "addin-bad" : "", verdict(t)));
  const deeper = (data.problems || []).filter((c) => c !== data.cell);
  if (deeper.length) parts.push(el("p", "addin-bad small", tr("Weiter unten rechnet {0} anders als Excel zeigt.", deeper.slice(0, 4).join(", "))));
  if (t.fixed && t.fixed.length) parts.push(el("p", "sub small", tr("Feste Zahlen in der Formel: {0}", t.fixed.join(", "))));
  if (t.children && t.children.length) {
    parts.push(el("p", "addin-why-head", tr("Entsteht aus")));
    parts.push(tree(t.children, data.sheet, 0));
  }
  if (t.uses && t.uses.length) {
    parts.push(el("p", "addin-why-head", tr("Fließt in")));
    const list = el("ul", "addin-tree");
    for (const use of t.uses) {
      const item = el("li");
      const line = el("div", "addin-row");
      line.append(el("span", "addin-cell", use.sheet !== data.sheet ? `${use.sheet}!${use.cell}` : use.cell));
      const name = nameOf(use.label);
      if (name) line.append(el("span", "addin-name", name));
      clickable(line, tr("{0} in Excel markieren", use.cell), () => select(use.sheet, use.cell));
      item.append(line);
      list.append(item);
    }
    if (t.uses_more) list.append(el("li", "sub small", tr("und weitere")));
    parts.push(list);
  } else if (t.kind !== "leer") {
    parts.push(el("p", "sub small", tr("Keine Formel liest diese Zelle.")));
  }
  if (data.cut) parts.push(el("p", "sub small", tr("Der Baum ist sehr groß, ich zeige die ersten 80 Zellen.")));
  box.replaceChildren(...parts);
}

async function listen() {
  // Nach jeder Änderung und beim Blattwechsel neu prüfen, sobald es kurz ruhig ist.
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.onChanged.add(async () => later());
      context.workbook.worksheets.onActivated.add(async () => later());
      await context.sync();
    });
  } catch (_) {
    // Ältere Excel-Versionen kennen diese Ereignisse nicht: dann alle paar Sekunden.
    setInterval(() => check(true), 5000);
  }
  // Nur ein anderes Zahlenformat (Start → Standard) meldet Excel getrennt, ab ExcelApi 1.9. Kennt Excel das nicht,
  // bleibt es beim Prüfen nach jeder anderen Änderung.
  if (!supports("1.9")) return;
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.onFormatChanged.add(async () => later());
      await context.sync();
    });
  } catch (_) {}
}

// --- Eigene Regeln je Spalte ---

let rules = [];
// Die Regel im Formular: ihre Spalte, das Blatt, auf dem die Spalte steht, und wie sie vor dem Ändern hieß.
let editing = null;
const ruleForm = document.getElementById("rule-form");
const RELATION_WORDS = {
  datum: { ge: tr("nicht vor"), le: tr("nicht nach") },
  zahl: { ge: tr("nicht kleiner als"), le: tr("nicht größer als") },
};

function rulesSay(text) {
  document.getElementById("rules-state").textContent = text;
}

async function loadRules() {
  try {
    const data = await post({}, "/api/addin/regeln");
    rules = data.regeln || [];
    renderRules();
  } catch (_) {
    // Antwortet Gegenprobe nicht, sagt es das Prüfen schon.
  }
}

function smallButton(label, go) {
  const button = el("button", "", label);
  button.type = "button";
  button.addEventListener("click", go);
  return button;
}

function renderRules() {
  document.getElementById("rules-count").textContent = rules.length ? String(rules.length) : "";
  document.getElementById("rules").replaceChildren(...rules.map((r) => {
    const li = el("li", "addin-rule");
    const text = el("div", "addin-rule-text");
    const head = el("div");
    head.append(el("strong", "", r.spalte));
    if (r.blatt) head.append(el("span", "sub small", tr(" nur auf „{0}“", r.blatt)));
    text.append(head, el("div", "small", r.text));
    const bar = el("div", "addin-rule-bar");
    const remove = smallButton(tr("Löschen"), async () => {
      // Kein Dialog im Seitenfenster: Der zweite Klick löscht.
      if (remove.dataset.sure !== "1") {
        remove.dataset.sure = "1";
        remove.textContent = tr("Wirklich?");
        setTimeout(() => {
          remove.dataset.sure = "";
          remove.textContent = tr("Löschen");
        }, 4000);
        return;
      }
      await removeRule(r);
    });
    bar.append(smallButton(tr("Ändern"), () => openRule(r.spalte, r.blatt)), remove);
    li.append(text, bar);
    return li;
  }));
}

async function removeRule(rule) {
  let data;
  try {
    data = await post({ spalte: rule.spalte, blatt: rule.blatt }, "/api/addin/regeln/loeschen");
  } catch (err) {
    rulesSay(silent(err));
    return;
  }
  if (data.state !== "ok") {
    rulesSay(data.text);
    return;
  }
  rules = data.regeln;
  renderRules();
  if (editing && editing.before && editing.before.spalte === rule.spalte && editing.before.blatt === rule.blatt) closeRule();
  rulesSay(tr("Die Regel für „{0}“ ist gelöscht. Ich prüfe neu.", rule.spalte));
  check(true);
}

// column leer: die Spalte der markierten Zelle. Sonst die Spalte dieses Namens, um ihre Regel zu ändern.
async function openRule(column, onlySheet, retried = false) {
  const box = document.getElementById("rules-box");
  box.open = true;
  let place = { sheet: active, cell: "" };
  if (!column) {
    try {
      place = await selectedPlace();
    } catch (_) {
      rulesSay(tr("Markiere in Excel eine einzelne Zelle der Spalte, dann noch einmal."));
      return;
    }
  }
  let data;
  try {
    data = await post({ ...place, spalte: column || "", ...(column ? { blatt: onlySheet || "" } : {}) }, "/api/addin/regeln/spalte");
  } catch (err) {
    rulesSay(silent(err));
    return;
  }
  if (data.state === "again") {
    // Gegenprobe kennt die Mappe noch nicht: einmal prüfen, dann noch einmal.
    if (!retried) {
      await check(false);
      openRule(column, onlySheet, true);
    }
    return;
  }
  if (data.state !== "ok") {
    rulesSay(data.text);
    return;
  }
  rulesSay("");
  fillRule(data);
  ruleForm.scrollIntoView({ block: "nearest" });
}

function field(id) {
  return document.getElementById(id);
}

function fillRule(data) {
  const r = data.regel || {};
  editing = { spalte: data.spalte, blatt: data.blatt, before: data.regel ? { spalte: r.spalte, blatt: r.blatt } : null };
  field("rule-column").textContent = data.regel ? tr("Regel für „{0}“ ändern", data.spalte) : tr("Neue Regel für „{0}“", data.spalte);
  field("rule-now").textContent = data.jetzt ? tr("Jetzt in der Spalte: {0}", data.jetzt) : "";
  field("rule-required").checked = !!r.pflicht;
  field("rule-kind").value = data.regel ? r.art || "" : data.art || "";
  field("rule-min").value = r.min === null || r.min === undefined ? "" : String(r.min).replace(".", ",");
  field("rule-max").value = r.max === null || r.max === undefined ? "" : String(r.max).replace(".", ",");
  field("rule-after").value = r.nicht_vor || "";
  field("rule-before").value = r.nicht_nach || "";
  field("rule-future").checked = !!r.keine_zukunft;
  const values = r.werte || [];
  field("rule-values").value = values.join(values.some((v) => v.includes(",")) ? "; " : ", ");
  field("rule-unique").checked = !!r.einmal;
  field("rule-pattern").value = r.muster || "";
  const other = field("rule-other");
  const columns = [...(data.spalten || [])];
  if (r.spalte2 && !columns.includes(r.spalte2)) columns.push(r.spalte2);
  other.replaceChildren(Object.assign(el("option", "", tr("keine")), { value: "" }), ...columns.map((c) => el("option", "", c)));
  other.options[0].value = "";
  other.value = r.spalte2 || "";
  field("rule-relation").value = r.beziehung || "ge";
  field("rule-only").checked = !!r.blatt;
  field("rule-only-text").textContent = data.blatt ? tr("nur auf dem Blatt „{0}“", data.blatt) : tr("nur auf diesem Blatt");
  field("rule-only").disabled = !data.blatt;
  field("rule-error").textContent = "";
  showKind();
  ruleForm.hidden = false;
}

// Was zur gewählten Art passt: von und bis für Zahlen, Grenzen für Daten, Wörter für den Vergleich.
function showKind() {
  const kind = field("rule-kind").value;
  field("rule-numbers").hidden = !["", "zahl", "ganz"].includes(kind);
  field("rule-dates").hidden = !["", "datum"].includes(kind);
  const words = RELATION_WORDS[kind === "datum" ? "datum" : "zahl"];
  const relation = field("rule-relation");
  relation.options[0].textContent = kind === "" ? tr("nicht vor / nicht kleiner als") : words.ge;
  relation.options[1].textContent = kind === "" ? tr("nicht nach / nicht größer als") : words.le;
}

function collectRule() {
  const numbers = !field("rule-numbers").hidden;
  const dates = !field("rule-dates").hidden;
  return {
    spalte: editing.spalte,
    blatt: field("rule-only").checked ? editing.blatt : "",
    pflicht: field("rule-required").checked,
    art: field("rule-kind").value,
    min: numbers ? field("rule-min").value.trim() : "",
    max: numbers ? field("rule-max").value.trim() : "",
    nicht_vor: dates ? field("rule-after").value : "",
    nicht_nach: dates ? field("rule-before").value : "",
    keine_zukunft: dates && field("rule-future").checked,
    werte: field("rule-values").value,
    einmal: field("rule-unique").checked,
    muster: field("rule-pattern").value.trim(),
    spalte2: field("rule-other").value,
    beziehung: field("rule-relation").value,
  };
}

async function saveRule(event) {
  event.preventDefault();
  if (!editing) return;
  const rule = collectRule();
  const button = field("rule-save");
  button.disabled = true;
  let data;
  try {
    data = await post({ regel: rule, vorher: editing.before }, "/api/addin/regeln/speichern");
  } catch (err) {
    data = { state: "error", text: silent(err) };
  }
  button.disabled = false;
  if (data.state !== "ok") {
    field("rule-error").textContent = data.text;
    return;
  }
  rules = data.regeln;
  renderRules();
  closeRule();
  rulesSay(tr("Die Regel für „{0}“ ist gespeichert. Ich prüfe neu.", rule.spalte));
  check(true);
}

function closeRule() {
  ruleForm.hidden = true;
  editing = null;
}

field("rule-new").addEventListener("click", () => openRule("", ""));
field("rule-kind").addEventListener("change", showKind);
field("rule-cancel").addEventListener("click", closeRule);
ruleForm.addEventListener("submit", saveRule);
document.getElementById("rules-box").addEventListener("toggle", (e) => {
  if (e.target.open) loadRules();
});

// --- Mappe in Worten ---

let describing = false;

async function describe(retried = false) {
  if (describing) return;
  describing = true;
  const note = document.getElementById("book-state");
  let data = null;
  try {
    data = await post({}, "/api/addin/mappe");
  } catch (err) {
    note.textContent = silent(err);
  }
  describing = false;
  if (!data) return;
  if (data.state === "again") {
    // Gegenprobe kennt die Mappe noch nicht: Nach dem Prüfen kommt describe() von selbst.
    if (!retried) check(false);
    return;
  }
  renderBook(data);
}

// Eine Zelle oder ein Bereich zum Anklicken: markiert ihn in Excel (auf seinem Blatt).
function place(sheet, cell, home) {
  const chip = el("span", "addin-chip", sheet !== home ? `${sheet}!${cell}` : cell);
  clickable(chip, tr("{0} in Excel markieren", cell), (e) => {
    if (e) e.stopPropagation();
    select(sheet, cell);
  });
  chip.addEventListener("keydown", (e) => e.stopPropagation());
  return chip;
}

function checkMark(check) {
  if (check === "ok") return el("span", "addin-good", "✓");
  if (check === "anders") return el("span", "addin-bad", "≠ nachgerechnet anders");
  return el("span", "sub small", tr("nicht nachgerechnet"));
}

function renderResult(r) {
  const li = el("li", `addin-hint${r.check === "anders" || r.upstream ? " fehler" : ""}`);
  const head = el("div", "addin-result-head");
  head.append(place(r.sheet, r.cell, ""));
  if (r.name) head.append(el("span", "", r.name));
  if (r.value) head.append(el("span", "addin-value", r.value));
  head.append(checkMark(r.check));
  li.append(head);
  for (const line of r.text) li.append(el("p", /^(?:Unterwegs|On the way)/.test(line) ? "addin-bad small" : "sub small", line));
  clickable(li, tr("{0} in Excel markieren", r.cell), () => select(r.sheet, r.cell));
  return li;
}

function renderSheet(s, open) {
  const box = el("details", "addin-box");
  box.open = open;
  const summary = el("summary");
  summary.append(el("span", "", s.name));
  const counts = [s.formulas === 1 ? tr("1 Formel") : tr("{0} Formeln", s.formulas)];
  if (s.hints && s.hints.fehler) counts.push(tr("{0} Fehler", s.hints.fehler));
  if (s.hints && s.hints.pruefen) counts.push(tr("{0} zum Prüfen", s.hints.pruefen));
  summary.append(el("span", `sub small${s.hints && s.hints.fehler ? " addin-bad" : ""}`, counts.join(", ")));
  box.append(summary);
  if (s.table) {
    const t = s.table;
    let text = tr("Tabelle {0} mit {1} Zeilen.", t.area, t.rows);
    if (t.inputs.length) text += ` Eingegeben: ${t.inputs.join(", ")}.`;
    if (t.computed.length) text += ` Gerechnet: ${t.computed.join(", ")}.`;
    box.append(el("p", "small", text));
  }
  for (const r of s.reads) {
    const how = r.functions.length ? tr(" mit {0}", r.functions.join(", ")) : "";
    box.append(el("p", "small", tr("Liest aus „{0}“{1}: {2}.", r.sheet, how, r.columns.join(", "))));
  }
  if (s.read_by.length) box.append(el("p", "small", tr("Daraus lesen {0}.", s.read_by.map((n) => tr("„{0}“", n)).join(", "))));
  for (const c of s.columns) {
    const part = el("div", "addin-part");
    const head = el("div", "addin-result-head");
    head.append(el("strong", "", c.name), place(s.name, c.area, s.name), checkMark(c.check));
    part.append(head, el("p", "addin-formula", c.formula));
    if (c.words && c.words !== c.formula) part.append(el("p", "small", c.words));
    for (const line of c.text) part.append(el("p", `sub small${/^(?:Nachgerechnet anders|Recalculated differently)/.test(line) ? " addin-bad" : ""}`, line));
    box.append(part);
  }
  for (const f of s.figures) {
    const part = el("div", "addin-part");
    const head = el("div", "addin-result-head");
    head.append(place(s.name, f.cell, s.name));
    if (f.name) head.append(el("strong", "", f.name));
    if (f.value) head.append(el("span", "addin-value", f.value));
    head.append(checkMark(f.check));
    part.append(head, el("p", "addin-formula", f.formula));
    if (f.words) part.append(el("p", "small", f.words));
    if (f.wrong) part.append(el("p", "addin-bad small", tr("Nachgerechnet: {0}.", f.wrong.join(", "))));
    if (f.today) part.append(el("p", "sub small", tr("Hängt vom heutigen Datum ab (HEUTE()).")));
    box.append(part);
  }
  if (s.figures_more) box.append(el("p", "sub small", tr("Und {0} weitere Formeln.", s.figures_more)));
  if (s.parts.length) {
    const list = el("ul", "small");
    for (const p of s.parts) list.append(el("li", "", p.text));
    box.append(list);
  }
  if (!s.table && !s.columns.length && !s.figures.length && !s.parts.length) box.append(el("p", "sub small", tr("Nur Text oder Zahlen, keine Formeln.")));
  return box;
}

function renderBook(data) {
  const note = document.getElementById("book-state");
  const box = document.getElementById("book");
  if (data.state !== "ok") {
    note.textContent = data.text || tr("Die Mappe kann ich gerade nicht beschreiben.");
    box.replaceChildren();
    return;
  }
  note.textContent = "";
  const parts = data.summary.map((line) => el("p", "", line));
  if (data.results.length) {
    parts.push(el("h3", "addin-book-head", tr("Ergebnisse")));
    const list = el("ul", "addin-hints");
    list.append(...data.results.map(renderResult));
    parts.push(list);
  }
  parts.push(el("h3", "addin-book-head", tr("Die Blätter")));
  const single = data.sheets.length === 1;
  parts.push(...data.sheets.map((s) => renderSheet(s, single)));
  box.replaceChildren(...parts);
}

// --- Prüfbericht ---

// file: über die hineingezogene Datei statt über die offene Mappe.
async function report(withData, file = false, retried = false) {
  const note = document.getElementById(file ? "file-report-state" : "report-state");
  note.textContent = tr("Ich schreibe den Bericht …");
  let data;
  try {
    data = await post({ daten: withData, datei: file }, "/api/addin/bericht");
  } catch (_) {
    note.textContent = tr("Den Bericht konnte ich nicht schreiben. Läuft Gegenprobe noch?");
    return;
  }
  if (data.state !== "ok") {
    if (!retried && !file) {
      await check(false);
      report(withData, file, true);
    } else {
      note.textContent = file ? tr("Erst muss ich die Datei prüfen: Prüfen, dann noch einmal.") : tr("Erst muss ich die Mappe prüfen: Neu prüfen, dann noch einmal.");
    }
    return;
  }
  const url = new URL(data.url, location.href).href;
  try {
    Office.context.ui.openBrowserWindow(url);
  } catch (_) {
    window.open(url, "_blank");
  }
  note.textContent = tr("Der Bericht ist im Browser aufgegangen.");
}

// --- Eine Datei: mit der offenen Mappe vergleichen oder prüfen, ohne sie zu öffnen ---

const KINDS = { ergebnis: tr("Ergebnis"), folge: tr("Dazwischen"), ersetzt: tr("Feste Zahl"), formel: tr("Formel"), eingabe: tr("Eingabe"), zeilen: tr("Zeilen"), spalten: tr("Spalten") };
let lastFile = null;
const fileState = document.getElementById("file-state");

const OTHER_FILES = [
  [/\.xls$/i, tr("Alte .xls-Dateien kann ich nicht lesen. Öffne sie in Excel und speichere sie als .xlsx.")],
  [/\.ods$/i, tr("Das ist eine Tabelle aus LibreOffice oder OpenOffice (.ods). Speichere sie dort als Excel-Datei (.xlsx), dann prüfe ich sie.")],
  [/\.numbers$/i, tr("Das ist eine Numbers-Datei. Exportiere sie in Numbers als Excel-Datei (Ablage → Exportieren → Excel), dann prüfe ich sie.")],
];

function takeFile(file) {
  const text = /\.(csv|tsv|txt)$/i.test(file.name);
  if (!text && !/\.(xlsx|xlsm)$/i.test(file.name)) {
    const other = OTHER_FILES.find(([pattern]) => pattern.test(file.name));
    fileState.textContent = other ? other[1] : tr("Ich prüfe Excel-Mappen (.xlsx) und CSV-Dateien.");
    return;
  }
  lastFile = file;
  document.getElementById("file-name").textContent = `Datei: ${file.name}`;
  document.getElementById("file-bar").hidden = false;
  document.getElementById("file-compare").hidden = text;
  // Eine CSV hat keine Formeln, ist also keine frühere Version der Mappe: gleich prüfen.
  if (text) return checkFile(file);
  // Zuerst vergleichen; ist es keine Version der offenen Mappe, gleich prüfen.
  compareFile(file);
}

async function send(path, file) {
  const response = await fetch(`${path}?name=${encodeURIComponent(file.name)}`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream", "X-VoiceAgent": "1", ...langHeaders() },
    body: file,
  });
  if (!response.ok) throw await failure(response);
  return response.json();
}

async function compareFile(file, retried = false) {
  fileState.textContent = tr("Ich vergleiche {0} mit der offenen Mappe …", file.name);
  document.getElementById("file-report").hidden = true;
  let data;
  try {
    data = await send("/api/addin/vergleich", file);
  } catch (err) {
    fileState.textContent = silent(err);
    return;
  }
  if (data.state === "again" && !retried) {
    // Gegenprobe kennt die offene Mappe noch nicht: einmal ganz lesen, dann vergleichen.
    await check(false);
    return compareFile(file, true);
  }
  if (data.state === "ok" && data.related === false) {
    return checkFile(file, tr("Die Datei hat kein Blatt wie die offene Mappe, darum prüfe ich sie für sich."));
  }
  if (data.state !== "ok") {
    fileState.textContent = data.text || tr("Vergleichen geht gerade nicht. Neu prüfen, dann noch einmal.");
    document.getElementById("file-result").replaceChildren();
    return;
  }
  renderDiff(data);
}

async function checkFile(file, lead = "") {
  fileState.textContent = tr("Ich prüfe {0} …", file.name);
  let data;
  try {
    data = await send("/api/addin/datei", file);
  } catch (err) {
    fileState.textContent = silent(err);
    return;
  }
  if (data.state !== "ok") {
    fileState.textContent = data.text || tr("Prüfen geht gerade nicht.");
    document.getElementById("file-result").replaceChildren();
    return;
  }
  renderFile(data, lead);
}

function chips(sheet, cells) {
  const box = el("div", "addin-cells");
  for (const cell of cells) {
    const chip = el("span", "addin-chip", cell);
    clickable(chip, tr("{0} in Excel markieren", cell), (e) => {
      if (e) e.stopPropagation();
      select(sheet, cell);
    });
    // Enter auf einer Zelle markiert nur sie, nicht auch die Zelle der ganzen Änderung.
    chip.addEventListener("keydown", (e) => e.stopPropagation());
    box.append(chip);
  }
  return box;
}

function card(level, badge, cell, text) {
  const li = el("li", `addin-hint ${level}`);
  const head = el("div", "addin-hint-head");
  if (cell) head.append(cellChip(cell, level));
  head.append(el("span", `addin-kind ${level}`, badge));
  li.append(head, el("p", "addin-text", text));
  return li;
}

function renderDiff(data) {
  fileState.textContent = data.text;
  document.getElementById("file-name").textContent = `Vorher: ${data.before} · Jetzt: ${data.now}`;
  const parts = [];
  const same = [];
  for (const sheet of data.sheets) {
    if (sheet.state === "gleich" && !sheet.old) {
      same.push(sheet.name);
      continue;
    }
    const title = sheet.old ? `${sheet.name} (vorher ${sheet.old})` : sheet.name;
    parts.push(el("h3", "addin-sheet", title));
    if (sheet.state === "neu") parts.push(el("p", "sub small", tr("Dieses Blatt ist neu.")));
    if (sheet.state === "weg") parts.push(el("p", "addin-bad small", tr("Dieses Blatt gibt es in der offenen Mappe nicht mehr.")));
    if (sheet.state === "gleich") parts.push(el("p", "sub small", tr("Sonst unverändert.")));
    const list = el("ul", "addin-hints");
    for (const item of sheet.items) {
      const li = card(item.level, KINDS[item.kind] || tr("Änderung"), item.cell, item.text);
      if (item.why) li.append(el("p", "small", item.why));
      if (item.tip) li.append(el("p", "sub small", item.tip));
      const others = (item.cells || []).filter((c) => c !== item.cell);
      if (others.length && sheet.state !== "weg") li.append(chips(sheet.name, others));
      if (item.cell && sheet.state !== "weg") clickable(li, tr("{0} in Excel markieren", item.cell), () => select(sheet.name, item.cell));
      list.append(li);
    }
    if (sheet.items.length) parts.push(list);
    if (sheet.more) parts.push(el("p", "sub small", tr("Und {0} weitere Änderungen auf diesem Blatt.", sheet.more)));
    if (sheet.cut) parts.push(el("p", "sub small", tr("Das Blatt ist sehr lang: Ich vergleiche die ersten 20 000 Zeilen.")));
  }
  if (same.length) parts.push(el("p", "sub small", tr("Unverändert: {0}.", same.join(", "))));
  document.getElementById("file-result").replaceChildren(...parts);
}

function renderFile(data, lead) {
  fileState.textContent = lead ? `${lead} ${data.text}` : data.text;
  document.getElementById("file-name").textContent = tr("Geprüft: {0} (nicht in Excel geöffnet)", data.book);
  const parts = [];
  for (const sheet of data.sheets) {
    parts.push(el("h3", "addin-sheet", sheet.hidden ? `${sheet.name} (ausgeblendet)` : sheet.name));
    if (sheet.failed) parts.push(el("p", "sub small", sheet.failed));
    if (!sheet.hints.length) {
      parts.push(el("p", "sub small", sheet.text));
      continue;
    }
    const list = el("ul", "addin-hints");
    for (const h of sheet.hints) {
      const li = card(h.level, LEVELS[h.level] || tr("Hinweis"), h.cell, h.text);
      if (h.tip) li.append(el("p", "sub small", h.tip));
      list.append(li);
    }
    parts.push(list);
  }
  document.getElementById("file-result").replaceChildren(...parts);
  document.getElementById("file-report").hidden = !data.sheets.length;
  document.getElementById("file-report-state").textContent = "";
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
  if (file) takeFile(file);
});
document.getElementById("file").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) takeFile(file);
  e.target.value = "";
});
document.getElementById("file-compare").addEventListener("click", async () => {
  if (!lastFile) return;
  await check(false);
  compareFile(lastFile);
});
document.getElementById("file-check").addEventListener("click", () => lastFile && checkFile(lastFile));
document.getElementById("file-plain").addEventListener("click", () => report(false, true));
document.getElementById("file-full").addEventListener("click", () => report(true, true));
// Eine Datei, die neben dem Feld landet, öffnet sonst das Fenster selbst.
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => e.preventDefault());

// Die Übungsmappe: Herunterladen kann das Fenster in Excel nicht, darum öffnet Gegenprobe eine Kopie selbst.
async function openPractice(event) {
  event.preventDefault();
  const link = event.currentTarget.href;
  const note = document.getElementById("practice-state");
  note.textContent = tr("Ich öffne die Übungsmappe …");
  let data = null;
  try {
    data = await post({}, "/api/addin/uebung");
  } catch (_) {
    data = null;
  }
  if (data && data.state === "ok") {
    note.textContent = tr("„{0}“ ist in Excel aufgegangen. Klick dort auf Start → Prüfer.", data.name);
  } else if (data && data.state === "closed") {
    note.textContent = tr("Öffnen ging nicht, die Mappe liegt hier: {0}. In Excel über Datei → Öffnen.", data.path);
  } else {
    // Dann lädt der Browser sie herunter.
    try {
      Office.context.ui.openBrowserWindow(link);
      note.textContent = tr("Die Übungsmappe kommt über den Browser, in Downloads.");
    } catch (_) {
      note.textContent = (data && data.text) || tr("Die Übungsmappe konnte ich nicht öffnen. Läuft Gegenprobe noch?");
    }
  }
}

// --- Über Gegenprobe: Grenzen und Datenschutz, beim ersten Öffnen und über „i“ ---
const ABOUT_SEEN = "gegenprobe-hinweis-1";
const about = document.getElementById("about");
let aboutFrom = null;

function aboutSeen() {
  try {
    return window.localStorage.getItem(ABOUT_SEEN) === "ja";
  } catch (_) {
    return false;
  }
}

function showAbout(first) {
  aboutFrom = document.activeElement;
  document.getElementById("about-title").textContent = first ? tr("Bevor du loslegst") : tr("Über Gegenprobe");
  document.getElementById("about-ok").textContent = first ? tr("Verstanden") : tr("Schließen");
  const version = document.getElementById("about-version");
  if (version.textContent.includes("__VERSION__")) version.textContent = "";
  about.hidden = false;
  document.getElementById("about-ok").focus();
}

function closeAbout() {
  about.hidden = true;
  try {
    window.localStorage.setItem(ABOUT_SEEN, "ja");
  } catch (_) {
    // Ohne Speicher im Fenster kommt der Hinweis beim nächsten Öffnen wieder.
  }
  if (aboutFrom && aboutFrom.focus) aboutFrom.focus();
}

for (const button of document.querySelectorAll("[data-lang]")) {
  button.setAttribute("aria-pressed", String(button.dataset.lang === LANG));
  button.addEventListener("click", () => {
    if (button.dataset.lang === LANG) return;
    closeAbout();
    chooseLang(button.dataset.lang);
  });
}
document.getElementById("about-open").addEventListener("click", () => showAbout(false));
document.getElementById("about-link").addEventListener("click", (event) => {
  event.preventDefault();
  showAbout(false);
});
document.getElementById("about-ok").addEventListener("click", closeAbout);
about.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeAbout();
});
if (!aboutSeen()) showAbout(true);

again.addEventListener("click", () => check(false));
document.getElementById("practice").addEventListener("click", openPractice);
document.getElementById("report-plain").addEventListener("click", () => report(false));
document.getElementById("report-full").addEventListener("click", () => report(true));
document.getElementById("tab-hints").addEventListener("click", () => showTab("hints"));
document.getElementById("tab-why").addEventListener("click", () => showTab("why"));
document.getElementById("tab-book").addEventListener("click", () => {
  showTab("book");
  if (!document.getElementById("book").childElementCount) document.getElementById("book-state").textContent = tr("Ich lese die ganze Mappe …");
  // Alle Blätter frisch lesen; danach beschreibt check() die Mappe.
  check(false);
});
document.getElementById("tab-file").addEventListener("click", () => showTab("file"));

if (window.Office) {
  Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) {
      state.textContent = tr("Dieses Fenster gehört in Excel: Start, dann Prüfer.");
      again.disabled = true;
      return;
    }
    loadPasst();
    check(false);
    loadRules();
    listen();
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelection);
    } catch (_) {
      // Ohne Auswahl-Ereignis gibt es keine Herleitung, der Rest geht trotzdem.
    }
  });
} else {
  state.textContent = tr("office.js von Microsoft lädt nicht. Ist der Mac mit dem Internet verbunden?");
  again.disabled = true;
}
