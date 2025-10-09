import { BudgetCore } from "/js/core/BudgetCore.js";

const STORAGE_KEY = "sameieportalen.budget.v1";
const core = new BudgetCore({ currency: "NOK", locale: "nb-NO" });

// --- Persistens (last inn)
try {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (raw) {
    const restored = BudgetCore.fromJSON(JSON.parse(raw));
    restored.items.forEach(i => core.addItem(i));
  }
} catch { /* ignorer korrupt lagring */ }

const $ = s => document.querySelector(s);
const list = $("#items");
const totalEl = $("#total");
const form = $("#add-item");

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

function save() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(core.toJSON()));
}

function render() {
  if (list) {
    list.innerHTML = core.items.map(i => `
      <li data-id="${i.id}">
        <span class="label">${escapeHtml(i.label)}</span>
        <span class="amount">${core.format(i.amount)}</span>
        <button class="remove" type="button" aria-label="Fjern">✕</button>
      </li>
    `).join("");
  }
  if (totalEl) totalEl.textContent = core.format(core.total());
  save();
}

// Legg til
form?.addEventListener("submit", (e) => {
  e.preventDefault();
  const label = form.label?.value?.trim();
  const amount = Number(form.amount?.value);
  const category = form.category?.value || "general";
  core.addItem({ label, amount, category });
  form.reset();
  render();
});

// Fjern (event delegation)
list?.addEventListener("click", (e) => {
  const btn = e.target.closest("button.remove");
  if (!btn) return;
  const id = btn.closest("li")?.dataset.id;
  if (id) {
    core.removeItem(id);
    render();
  }
});

// Første render
render();
