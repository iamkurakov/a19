const test = require('node:test');
const assert = require('node:assert/strict');
const E = require('../src/engine.js');
const catalog = [
  {id:'M1', name:'ПВХ', category:'material', unit:'м²', driver:'area', factor:1, unitCost:200, unitPrice:400},
  {id:'W1', name:'Монтаж', category:'work', unit:'шт', driver:'manual', unitCost:100, unitPrice:240},
];
const order = () => ({ discountPct:10, otherExpense:50, objects:[{id:'O1',title:'Вывеска',widthMm:2000,heightMm:1000,count:2,lines:[{itemId:'M1'},{itemId:'W1',qty:1}]}] });
test('geometry works in meters and quantities are multiplied by object count',()=>{
  assert.deepEqual(E.quantity({driver:'area',factor:1},{widthMm:2000,heightMm:1000},2),{each:2,total:4,area:2,perimeter:6});
});
test('calculates materials, works, discount, expense and profit independently',()=>{
  const x=E.calcOrder(order(),catalog);
  assert.equal(x.material,800);assert.equal(x.labor,200);assert.equal(x.cost,1050);
  assert.equal(x.sale,2080);assert.equal(x.discount,208);assert.equal(x.total,1872);assert.equal(x.profit,822);
});
test('line-level price and cost overrides do not modify catalog',()=>{
  const o=order();o.objects[0].lines[0].priceOverride=500;o.objects[0].lines[0].costOverride=220;
  const x=E.calcOrder(o,catalog);assert.equal(x.objects[0].lines[0].sale,2000);
  assert.equal(x.objects[0].lines[0].cost,880);assert.equal(catalog[0].unitCost,200);
});
test('manual quantity override beats automatic geometry',()=>{
  const o=order();o.objects[0].lines[0].qtyOverride=3;
  assert.equal(E.calcOrder(o,catalog).objects[0].lines[0].total,6);
});
test('freezing preserves historical prices after catalog changes',()=>{
  const o=order();E.freezeOrder(o,catalog);
  const first=E.calcOrder(o,catalog).total;
  const changed=structuredClone(catalog);changed[0].unitPrice=1900;changed[1].unitCost=600;
  assert.equal(E.calcOrder(o,changed).total,first);
});
test('incomplete estimate is visibly marked',()=>{
  const o=order();const list=structuredClone(catalog);list[0].unitCost=0;
  assert.equal(E.calcOrder(o,list).incomplete,true);
});
