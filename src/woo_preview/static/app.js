"use strict";

const KEY = new URLSearchParams(location.search).get("k") || "";
const PAGE = 40;
const STEPS = ["step-file", "step-progress", "step-error", "step-preview", "step-done"];
const LABELS = {
  "Title": "Nazwa",
  "Content": "Opis",
  "Short Description": "Krótki opis",
  "Image URL": "Zdjęcia",
  "Kategorie produktów": "Kategorie",
  "Product Tags": "Tagi",
  "Status": "Status",
  "Shipping Class": "Klasa wysyłkowa",
  "Wysyłka w ciągu": "Wysyłka w ciągu",
  "Backorders": "Zamówienia bez stanu",
  "Weight": "Waga",
  "Length": "Długość",
  "Width": "Szerokość",
  "Height": "Wysokość",
  "_alkan_compatible_product_skus": "Produkty kompatybilne",
  "_alkan_compatible_product_ids": "Produkty kompatybilne",
};

const $ = (id) => document.getElementById(id);
let state = null;
let products = [];
let productsVersion = -1;
let shown = PAGE;
let mode = "both";
let filter = "";
let poller = null;
let dismissedError = "";
let profilesFilled = false;

function esc(value) {
  return String(value ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function plural(n, one, few, many) {
  if (n === 1) return one;
  if ([2, 3, 4].includes(n % 10) && ![12, 13, 14].includes(n % 100)) return few;
  return many;
}

function label(column) {
  if (LABELS[column]) return LABELS[column];
  return column.startsWith("Atrybut Produktu: ") ? column.slice("Atrybut Produktu: ".length) : column;
}

async function api(path, options = {}) {
  const headers = Object.assign({ "X-Key": KEY }, options.headers || {});
  const response = await fetch(path, Object.assign({}, options, { headers }));
  let data = {};
  try { data = await response.json(); } catch (e) { /* pusta odpowiedz */ }
  if (!response.ok) throw new Error(data.error || `Błąd ${response.status}`);
  return data;
}

function showStep(id) {
  for (const step of STEPS) $(step).classList.toggle("hidden", step !== id);
}

function confirmDialog(title, bodyHtml, yesLabel, danger) {
  return new Promise((resolve) => {
    $("modal-title").textContent = title;
    $("modal-body").innerHTML = bodyHtml;
    const yes = $("modal-yes");
    yes.textContent = yesLabel;
    yes.className = danger ? "primary prod" : "primary";
    $("modal").classList.remove("hidden");
    const done = (answer) => {
      $("modal").classList.add("hidden");
      yes.onclick = null;
      $("modal-no").onclick = null;
      resolve(answer);
    };
    yes.onclick = () => done(true);
    $("modal-no").onclick = () => done(false);
  });
}

function fail(message) {
  $("error-text").textContent = message;
  $("error-back").onclick = () => (state ? render() : showStep("step-file"));
  showStep("step-error");
}

// --- stan aplikacji -----------------------------------------------------------------------

async function refresh() {
  try {
    state = await api("/api/state");
  } catch (error) {
    stopPolling();
    fail(error.message);
    return;
  }
  const session = state.session;
  if (session && session.version !== productsVersion && state.job.state !== "running") {
    const data = await api("/api/products");
    products = data.products;
    productsVersion = data.version;
    shown = PAGE;
  }
  render();
}

function startPolling() {
  if (!poller) poller = setInterval(refresh, 1000);
}

function stopPolling() {
  if (poller) clearInterval(poller);
  poller = null;
}

function render() {
  fillProfiles();
  const job = state.job;
  const session = state.session;
  if (job.state === "running") {
    renderProgress(job, session);
    showStep("step-progress");
    startPolling();
    return;
  }
  stopPolling();
  const errorKey = `${job.kind}|${job.message}|${job.log.length}`;
  if (job.state === "error" && dismissedError !== errorKey) {
    $("error-text").textContent = job.message;
    $("error-back").onclick = () => { dismissedError = errorKey; render(); };
    showStep("step-error");
    return;
  }
  if (!session) {
    showStep("step-file");
    return;
  }
  if (session.applied) {
    renderDone(session);
    showStep("step-done");
    return;
  }
  renderPreview(session);
  showStep("step-preview");
}

function renderProgress(job, session) {
  const titles = {
    preview: "Sprawdzam plik w sklepie…",
    apply: "Zapisuję zmiany w sklepie…",
    revert: "Cofam zapis…",
  };
  $("progress-title").textContent = titles[job.kind] || "Pracuję…";
  const last = job.log.length ? job.log[job.log.length - 1] : "Łączę się ze sklepem…";
  let line = esc(last);
  if (session && session.production && (job.kind === "apply" || job.kind === "revert")) {
    line += '<br><strong class="note prod">Potwierdź w okienku Windows „AlkanInator”. Jeśli go nie widać, sprawdź pasek zadań.</strong>';
  }
  $("progress-line").innerHTML = line;
  $("log").textContent = job.log.join("\n");
}

// --- krok 1: plik ------------------------------------------------------------------------

function fillProfiles() {
  if (profilesFilled || !state) return;
  profilesFilled = true;
  const select = $("profile");
  select.innerHTML = "";
  for (const profile of state.profiles) {
    const option = document.createElement("option");
    option.value = profile.name;
    option.textContent = `${profile.host || profile.name}${profile.production ? " - PRODUKCJA" : " - sklep testowy"} (${profile.name})`;
    option.selected = profile.default;
    select.appendChild(option);
  }
  updateFileStep();
}

function selectedProfile() {
  return (state?.profiles || []).find((p) => p.name === $("profile").value);
}

function updateFileStep() {
  const profile = selectedProfile();
  const note = $("profile-note");
  if (!state || !state.profiles.length) {
    note.textContent = "Brak konfiguracji sklepu (src/woo_api/config.json) - poproś o jej uzupełnienie.";
    note.className = "note prod";
  } else if (profile && profile.production) {
    note.textContent = `Uwaga: to sklep PRODUKCYJNY (${profile.host}). Przed zapisem pojawi się okienko Windows do potwierdzenia.`;
    note.className = "note prod";
  } else {
    note.textContent = profile ? `Sklep testowy ${profile.host} - zmiany nie trafią do klientów.` : "";
    note.className = "note";
  }
  $("preview").disabled = !(profile && $("file").files.length);
}

async function startPreview() {
  const file = $("file").files[0];
  if (!file) return;
  $("preview").disabled = true;
  try {
    await api("/api/preview", {
      method: "POST",
      headers: {
        "Content-Type": "application/octet-stream",
        "X-Profile": encodeURIComponent($("profile").value),
        "X-File-Name": encodeURIComponent(file.name),
        "X-All-Rows": $("all-rows").checked ? "1" : "0",
      },
      body: await file.arrayBuffer(),
    });
    await refresh();
  } catch (error) {
    fail(error.message);
  } finally {
    updateFileStep();
  }
}

// --- krok 2: podglad ---------------------------------------------------------------------

function stat(value, text, cls) {
  return `<div class="stat ${cls || ""}"><b>${esc(value)}</b><span>${esc(text)}</span></div>`;
}

function renderPreview(session) {
  const summary = session.summary || {};
  const errors = summary.errors || 0;
  const warnings = summary.warnings || 0;
  const n = session.to_apply;
  const badge = session.production ? '<span class="badge prod">PRODUKCJA</span>' : '<span class="badge local">sklep testowy</span>';
  $("preview-head").innerHTML = `Plik <strong>${esc(session.file)}</strong> → sklep <strong>${esc(session.host)}</strong> ${badge}`;
  $("stats").innerHTML =
    stat(summary.rows || 0, "wierszy sprawdzonych w sklepie") +
    (summary.skipped_unchanged ? stat(summary.skipped_unchanged, "pominiętych - bez zmian od pobrania") : "") +
    stat(n, plural(n, "produkt do zapisu", "produkty do zapisu", "produktów do zapisu"), n ? "good" : "") +
    stat(summary.changes || 0, "zmienionych pól") +
    stat(errors, plural(errors, "produkt z błędem", "produkty z błędami", "produktów z błędami"), errors ? "bad" : "") +
    stat(warnings, plural(warnings, "ostrzeżenie", "ostrzeżenia", "ostrzeżeń"), warnings ? "warn" : "");

  const blockers = [];
  if (errors) blockers.push(`Plik ma błędy w ${errors} ${plural(errors, "produkcie", "produktach", "produktach")} (na liście na czerwono, na początku). Popraw plik i wczytaj go ponownie - do tego czasu zapis jest wstrzymany.`);
  if (!session.writes_enabled) blockers.push("Zapis przez API jest wyłączony w sklepie (WooCommerce > AlkanInator API).");
  if (session.max_products && n > session.max_products) blockers.push(`Za dużo produktów naraz: ${n}, a sklep przyjmuje najwyżej ${session.max_products} w jednym zapisie. Podziel plik na mniejsze części.`);
  const infos = [];
  if (!n && !errors) infos.push("Sklep ma już wszystkie wartości z tego pliku - nie ma nic do zapisania.");
  if (warnings) infos.push("Ostrzeżenia nie blokują zapisu, ale przejrzyj je na liście (np. produkt zmieniony w sklepie po pobraniu pliku - plik może cofnąć cudze zmiany).");
  $("blockers").innerHTML = blockers.map((t) => `<div class="blocker">${esc(t)}</div>`).join("") + infos.map((t) => `<div class="info">${esc(t)}</div>`).join("");

  const extra = [];
  if ((session.new_terms || []).length) {
    const terms = session.new_terms.map((t) => `${esc(t.taxonomy)}: <strong>${esc(t.name)}</strong>`).join(", ");
    extra.push(`<div class="info">Wartości, których sklep jeszcze nie zna: ${terms}. Nowe wartości i tagi wymagają zgody - zapytaj, zanim je dodasz.</div>`);
  }
  const groups = session.variant_groups || {};
  let related = groups.live
    ? `Warianty na kartach: ${esc(groups.count || 0)} grup WPC prosto z ${esc(groups.source)}, z wariantami ze sklepu.`
    : `Warianty na kartach: grupy WPC z ${esc(groups.source || "zrzutu")} (${esc(groups.count || 0)} grup) - ten sklep ma starszą wtyczkę bridge, więc inne warianty widać tylko wtedy, gdy są w tym pliku.`;
  if (session.related_fetched) related += ` Dociągnięto ze sklepu ${session.related_fetched} ${plural(session.related_fetched, "produkt", "produkty", "produktów")} spoza pliku (warianty i powiązania, tylko odczyt).`;
  extra.push(`<p class="hint">${related}</p>`);
  if ((session.ignored_columns || []).length) {
    extra.push(`<div class="info">Kolumny pominięte (sklep ich nie zna): ${esc(session.ignored_columns.join(", "))}</div>`);
  }
  $("extra").innerHTML = extra.join("");

  renderColumns(session);
  renderProducts();
  renderApprove(session, blockers.length > 0 || errors > 0);
}

function renderColumns(session) {
  const box = $("columns-box");
  const all = session.all_columns || [];
  box.classList.toggle("hidden", all.length < 2);
  $("columns").innerHTML = all
    .map((column) => `<label class="chip" title="${esc(column)}"><input type="checkbox" value="${esc(column)}" ${session.columns.includes(column) ? "checked" : ""}> ${esc(label(column))}</label>`)
    .join("");
  const update = () => {
    const chosen = [...$("columns").querySelectorAll("input:checked")].map((i) => i.value);
    const same = chosen.length === session.columns.length && chosen.every((c) => session.columns.includes(c));
    $("recalc").disabled = same || chosen.length === 0;
  };
  $("columns").onchange = update;
  update();
}

function visibleProducts() {
  const needle = filter.trim().toLowerCase();
  if (!needle) return products;
  return products.filter((p) => `${p.title} ${p.sku} ${p.id}`.toLowerCase().includes(needle));
}

function statusBadge(product) {
  if (product.errors.length) return '<span class="status error">Błąd</span>';
  if (product.status === "changed") return '<span class="status changed">Do zapisu</span>';
  return '<span class="status unchanged">Bez zmian</span>';
}

function renderProducts() {
  const list = visibleProducts();
  const html = list.slice(0, shown).map((p) => {
    const tags = p.columns.map((c) => `<span class="tag" title="${esc(c)}">${esc(label(c))}</span>`).join("");
    return `<details class="product ${p.errors.length ? "err" : ""}" data-i="${p.i}">
      <summary>${statusBadge(p)}<span class="ptitle"><strong>${esc(p.title)}</strong><span class="sku">${esc(p.sku)}${p.id ? " · id " + p.id : ""}</span></span><span>${tags}</span></summary>
      <div class="pbody"></div>
    </details>`;
  }).join("");
  $("products").innerHTML = html || '<p class="hint">Brak produktów do pokazania.</p>';
  $("more").classList.toggle("hidden", list.length <= shown);
  $("more").querySelector("button").textContent = `Pokaż kolejne (pokazano ${Math.min(shown, list.length)} z ${list.length})`;
  for (const details of $("products").querySelectorAll("details")) {
    details.addEventListener("toggle", () => { if (details.open) fillProduct(details); });
  }
}

function fillProduct(details) {
  const product = products[Number(details.dataset.i)];
  const body = details.querySelector(".pbody");
  const messages = product.errors.map((e) => `<li class="e">${esc(e)}</li>`).concat(product.warnings.map((w) => `<li class="w">${esc(w)}</li>`));
  const rows = product.changes.map((c) => `<tr><th title="${esc(c.column)}">${esc(label(c.column))}</th><td>${c.html}</td></tr>`).join("");
  body.innerHTML =
    (messages.length ? `<ul class="msgs">${messages.join("")}</ul>` : "") +
    (rows ? `<table class="changes"><tbody>${rows}</tbody></table>` : "") +
    (product.changes.length ? '<div class="cards"></div>' : "");
  fillCards(details);
}

function fillCards(details) {
  const cards = details.querySelector(".cards");
  if (!cards) return;
  const i = details.dataset.i;
  const sides = mode === "both" ? ["before", "after"] : [mode];
  cards.className = sides.length === 1 ? "cards single" : "cards";
  cards.innerHTML = sides.map((side) => `<div class="card-col ${side}"><h4>${side === "before" ? "Przed zmianą (teraz w sklepie)" : "Po zmianie"}</h4>
      <iframe sandbox="allow-same-origin" loading="lazy" title="${side}" src="/card?k=${encodeURIComponent(KEY)}&i=${i}&side=${side}"></iframe></div>`).join("");
  for (const frame of cards.querySelectorAll("iframe")) {
    frame.addEventListener("load", () => {
      try {
        frame.style.height = `${frame.contentDocument.documentElement.scrollHeight + 16}px`;
      } catch (e) { /* zostaje domyslna wysokosc */ }
    });
  }
}

function renderApprove(session, blocked) {
  const n = session.to_apply;
  const button = $("apply");
  button.className = session.production ? "primary big prod" : "primary big";
  button.textContent = n ? `Zatwierdź i zapisz ${n} ${plural(n, "produkt", "produkty", "produktów")}` : "Zatwierdź i zapisz";
  button.disabled = blocked || !n;
  const where = session.production ? `w sklepie PRODUKCYJNYM <strong>${esc(session.host)}</strong>` : `w sklepie testowym <strong>${esc(session.host)}</strong>`;
  let note = n ? `Zapis zmieni ${n} ${plural(n, "produkt", "produkty", "produktów")} ${where}. Zapisze się dokładnie to, co widzisz powyżej.` : "Nie ma nic do zapisania.";
  if (blocked) note = "Zapis jest wstrzymany - zobacz komunikaty na górze.";
  $("approve-note").innerHTML = note;
}

async function applyChanges() {
  const session = state.session;
  const n = session.to_apply;
  const fields = session.columns.map(label).join(", ");
  let body = `<p>W sklepie <strong>${esc(session.host)}</strong> zmieni się <strong>${n} ${plural(n, "produkt", "produkty", "produktów")}</strong>.</p><p>Pola: ${esc(fields)}</p>`;
  if (session.production) {
    body += '<p class="note prod">To sklep PRODUKCYJNY - zmiany od razu zobaczą klienci. Za chwilę pojawi się okienko Windows: kliknij w nim „Tak”, żeby zapisać.</p>';
  }
  body += "<p>Zapis można potem cofnąć przyciskiem „Cofnij ten zapis”.</p>";
  if (!(await confirmDialog("Zapisać zmiany?", body, "Tak, zapisz", session.production))) return;
  try {
    await api("/api/apply", { method: "POST" });
    await refresh();
  } catch (error) {
    fail(error.message);
  }
}

async function recalc(columns) {
  try {
    await api("/api/recalc", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ columns }) });
    await refresh();
  } catch (error) {
    fail(error.message);
  }
}

// --- krok 3: wynik -----------------------------------------------------------------------

function renderDone(session) {
  const applied = session.applied.summary || {};
  const n = applied.applied || 0;
  const reverted = session.reverted;
  $("done-title").textContent = reverted ? "Zapis cofnięty" : "Zapisano";
  let html = reverted
    ? `<p class="lead">Przywrócono poprzednie wartości w sklepie <strong>${esc(session.host)}</strong> (produktów: ${esc((reverted.summary || {}).applied ?? (reverted.summary || {}).changed ?? "")}).</p>`
    : `<p class="lead">Zapisano ${n} ${plural(n, "produkt", "produkty", "produktów")} w sklepie <strong>${esc(session.host)}</strong>.</p>`;
  html += `<p>Identyfikator zapisu: <code>${esc(session.changeset)}</code> - po nim zapis można znaleźć i cofnąć także w WooCommerce > AlkanInator API.</p>`;
  const errors = (reverted ? reverted.errors : session.applied.errors) || [];
  if (errors.length) {
    html += `<div class="blocker">Nie wszystko się udało:<ul>${errors.map((e) => `<li>${esc(e)}</li>`).join("")}</ul></div>`;
  }
  const reports = session.reports.concat(reverted && reverted.report ? [reverted.report] : []);
  html += `<details class="log"><summary>Raporty (Excel)</summary><pre>${esc(reports.join("\n"))}</pre></details>`;
  $("done").innerHTML = html;
  $("revert").classList.toggle("hidden", Boolean(reverted));
}

async function revertChanges() {
  const session = state.session;
  let body = `<p>Przywrócić poprzednie wartości w sklepie <strong>${esc(session.host)}</strong> (zapis <code>${esc(session.changeset)}</code>)?</p><p>Pola zmienione w sklepie po tym zapisie zostaną bez zmian.</p>`;
  if (session.production) body += '<p class="note prod">Sklep PRODUKCYJNY - pojawi się okienko Windows do potwierdzenia.</p>';
  if (!(await confirmDialog("Cofnąć zapis?", body, "Tak, cofnij", true))) return;
  try {
    await api("/api/revert", { method: "POST" });
    await refresh();
  } catch (error) {
    fail(error.message);
  }
}

async function reset() {
  try {
    await api("/api/reset", { method: "POST" });
    $("file").value = "";
    productsVersion = -1;
    products = [];
    await refresh();
  } catch (error) {
    fail(error.message);
  }
}

// --- zdarzenia ---------------------------------------------------------------------------

$("profile").addEventListener("change", updateFileStep);
$("file").addEventListener("change", updateFileStep);
$("preview").addEventListener("click", startPreview);
$("apply").addEventListener("click", applyChanges);
$("revert").addEventListener("click", revertChanges);
$("again").addEventListener("click", reset);
$("back").addEventListener("click", reset);
$("refresh").addEventListener("click", () => recalc(state.session.columns));
$("recalc").addEventListener("click", () => recalc([...$("columns").querySelectorAll("input:checked")].map((i) => i.value)));
$("search").addEventListener("input", (event) => { filter = event.target.value; shown = PAGE; renderProducts(); });
$("more").querySelector("button").addEventListener("click", () => { shown += PAGE; renderProducts(); });
$("mode").addEventListener("click", (event) => {
  const button = event.target.closest("button");
  if (!button) return;
  mode = button.dataset.mode;
  for (const b of $("mode").querySelectorAll("button")) b.classList.toggle("on", b === button);
  for (const details of $("products").querySelectorAll("details[open]")) fillCards(details);
});
$("quit").addEventListener("click", async () => {
  if (!(await confirmDialog("Zamknąć aplikację?", "<p>Aplikację otworzysz ponownie ikoną „Wgraj zmiany do sklepu”.</p>", "Zamknij", false))) return;
  try {
    await api("/api/shutdown", { method: "POST" });
    stopPolling();
    $("bye").classList.remove("hidden");
  } catch (error) {
    fail(error.message);
  }
});

setInterval(() => api("/api/ping", { method: "POST" }).catch(() => {}), 30000);
refresh();
