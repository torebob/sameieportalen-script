// Kjerne: ren logikk uten DOM. Lett å teste og gjenbruke.
export class BudgetCore {
  #items = []; // { id, label, amount, category }
  #currency; #locale;

  constructor({ currency = "NOK", locale = "nb-NO", initial = [] } = {}) {
    this.#currency = currency;
    this.#locale = locale;
    initial.forEach(i => this.addItem(i));
  }

  get items() { return [...this.#items]; }

  addItem(item) {
    const { label, amount, category = "general", id } = item ?? {};
    if (!label || typeof amount !== "number" || Number.isNaN(amount)) {
      throw new Error("Ugyldig item: label og numerisk amount kreves");
    }
    const newId = id ?? (crypto.randomUUID?.() || String(Date.now() + Math.random()));
    const normalized = { id: newId, label: String(label), amount: +amount, category: String(category) };
    this.#items.push(normalized);
    return newId;
  }

  removeItem(id) {
    const before = this.#items.length;
    this.#items = this.#items.filter(i => i.id !== id);
    return before !== this.#items.length;
  }

  updateItem(id, patch) {
    const idx = this.#items.findIndex(i => i.id === id);
    if (idx === -1) return false;
    const next = { ...this.#items[idx], ...patch };
    if (!next.label || typeof next.amount !== "number" || Number.isNaN(next.amount)) {
      throw new Error("Ugyldig oppdatering");
    }
    this.#items[idx] = next;
    return true;
  }

  clear() { this.#items = []; }

  total(opts = {}) {
    if (opts.byCategory) {
      return this.#items.reduce((acc, i) => {
        acc[i.category] = (acc[i.category] || 0) + i.amount;
        return acc;
      }, {});
    }
    return this.#items.reduce((sum, i) => sum + i.amount, 0);
  }

  format(n) {
    return new Intl.NumberFormat(this.#locale, {
      style: "currency",
      currency: this.#currency
    }).format(n);
  }

  toJSON() {
    return {
      meta: { currency: this.#currency, locale: this.#locale, version: 1 },
      items: this.#items
    };
  }

  static fromJSON(json) {
    const { meta, items } = json ?? {};
    return new BudgetCore({
      currency: meta?.currency ?? "NOK",
      locale: meta?.locale ?? "nb-NO",
      initial: Array.isArray(items) ? items : []
    });
  }
}
