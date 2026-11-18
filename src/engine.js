/* A19 · independent calculation kernel. No Ideal Potolok imports or remote storage. */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.A19Engine = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const amount = (v) => {
    if (v === null || v === undefined || v === '') return 0;
    const n = Number(String(v).replace(',', '.'));
    return Number.isFinite(n) ? n : 0;
  };
  const money = (v) => Math.round((amount(v) + Number.EPSILON) * 100) / 100;
  const nonNegative = (v) => Math.max(0, amount(v));
  const qtyRound = (v) => Math.round(nonNegative(v) * 10000) / 10000;
  const quantity = (line, geometry, count) => {
    const width = nonNegative(geometry.widthMm);
    const height = nonNegative(geometry.heightMm);
    const area = width * height / 1000000;
    const perimeter = 2 * (width + height) / 1000;
    const factor = nonNegative(line.factor ?? 1);
    const driver = line.driver || 'manual';
    const each = line.qtyOverride !== null && line.qtyOverride !== undefined && line.qtyOverride !== ''
      ? nonNegative(line.qtyOverride)
      : driver === 'area' ? area * factor
      : driver === 'perimeter' ? perimeter * factor
      : nonNegative(line.qty ?? 1);
    return { each: qtyRound(each), total: qtyRound(each * Math.max(1, Math.floor(nonNegative(count) || 1))), area: qtyRound(area), perimeter: qtyRound(perimeter) };
  };
  function calcObject(object, catalog) {
    const count = Math.max(1, Math.floor(nonNegative(object.count) || 1));
    const lookup = Object.fromEntries((catalog || []).map(item => [item.id, item]));
    const lines = (object.lines || []).map((line) => {
      const item = lookup[line.itemId] || {};
      const fix = line.fix || null;
      const q = quantity({
        ...line,
        driver: line.driver ?? item.driver ?? 'manual',
        factor: line.factor ?? item.factor ?? 1,
      }, object, count);
      const unitCost = nonNegative(line.costOverride ?? fix?.unitCost ?? item.unitCost ?? line.unitCost ?? 0);
      const unitPrice = nonNegative(line.priceOverride ?? fix?.unitPrice ?? item.unitPrice ?? line.unitPrice ?? 0);
      const category = fix?.category ?? item.category ?? line.category ?? 'material';
      return {
        ...q, itemId: line.itemId || null, category,
        name: fix?.name ?? item.name ?? line.name ?? 'Позиция',
        unit: fix?.unit ?? item.unit ?? line.unit ?? 'шт',
        unitCost, unitPrice,
        cost: money(q.total * unitCost),
        sale: money(q.total * unitPrice),
        incomplete: unitPrice === 0 || unitCost === 0,
      };
    });
    return {
      id: object.id,
      title: object.title,
      count,
      area: qtyRound(nonNegative(object.widthMm) * nonNegative(object.heightMm) / 1000000),
      lines,
      sale: money(lines.reduce((s, x) => s + x.sale, 0)),
      material: money(lines.filter(x => x.category === 'material').reduce((s, x) => s + x.cost, 0)),
      labor: money(lines.filter(x => x.category === 'work').reduce((s, x) => s + x.cost, 0)),
      cost: money(lines.reduce((s, x) => s + x.cost, 0)),
    };
  }
  function calcOrder(order, catalog) {
    const objects = (order.objects || []).map(x => calcObject(x, catalog));
    const sale = money(objects.reduce((s,x) => s + x.sale, 0));
    const material = money(objects.reduce((s,x) => s + x.material, 0));
    const labor = money(objects.reduce((s,x) => s + x.labor, 0));
    const expense = money(nonNegative(order.otherExpense));
    const cost = money(objects.reduce((s,x) => s + x.cost, 0) + expense);
    const discountPct = Math.min(100, nonNegative(order.discountPct));
    const discount = money(sale * discountPct / 100);
    const total = money(sale - discount);
    return { objects, sale, material, labor, expense, cost, discountPct, discount, total,
      profit: money(total - cost), margin: total ? money((total - cost) / total * 100) : 0,
      incomplete: objects.some(o => o.lines.some(l => l.incomplete)) || !objects.some(o => o.lines.length) };
  }
  function freezeOrder(order, catalog) {
    const lookup = Object.fromEntries((catalog || []).map(x => [x.id, x]));
    (order.objects || []).forEach(obj => (obj.lines || []).forEach(line => {
      if (line.fix) return;
      const item = lookup[line.itemId] || line;
      line.fix = {
        name: item.name || 'Позиция', category: item.category || 'material', unit: item.unit || 'шт',
        unitCost: nonNegative(line.costOverride ?? item.unitCost),
        unitPrice: nonNegative(line.priceOverride ?? item.unitPrice),
      };
    }));
  }
  return { amount, money, quantity, calcObject, calcOrder, freezeOrder };
});
