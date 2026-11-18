const test=require('node:test');
const assert=require('node:assert/strict');
const {Model,ExcelError,normalizeTables}=require('../src/excel.js');
function sample(){
 return {format:'a19-excel-model-v1',sheets:{
  'Объём1':{A1:{v:4},A2:{v:6},A3:{f:'SUM(A1:A2)',v:10},B1:{f:'IF(A1>1,A1*3,0)',v:12},B2:{f:'ROUNDUP(A1/3,0)',v:2},B3:{f:'IF(ISNA(VLOOKUP("x",\'справочник\'!A1:C2,3,0)),0,10)',v:0},C1:{f:'VLOOKUP("Q",\'справочник\'!A1:C2,3,0)',v:42},C2:{f:'INDEX(A1:A2,2)',v:6},C3:{f:'SUBTOTAL(109,A1:A2)',v:10}},
  'справочник':{A1:{v:'Q'},B1:{v:'misc'},C1:{v:42},A2:{v:'R'},B2:{v:'data'},C2:{v:88}}
 },sheetOrder:['Объём1','справочник'],names:{},localNames:{},external:{},externalNames:{},tables:{},overrides:{}};
}
test('IF / SUM / ROUNDUP / INDEX / SUBTOTAL / VLOOKUP',()=>{
 const m=new Model(sample());assert.equal(m.get('Объём1','A3'),10);assert.equal(m.get('Объём1','B1'),12);
 assert.equal(m.get('Объём1','B2'),2);assert.equal(m.get('Объём1','B3'),0);
 assert.equal(m.get('Объём1','C1'),42);assert.equal(m.get('Объём1','C2'),6);
 assert.equal(m.get('Объём1','C3'),10);
});
test('editable overrides are propagated and restored without altering formulas',()=>{
 const m=new Model(sample());m.edit('Объём1','A1',8);assert.equal(m.get('Объём1','A3'),14);
 assert.equal(m.get('Объём1','B1'),24);assert.equal(m.serialize().sheets['Объём1'].B1.f,'IF(A1>1,A1*3,0)');
 m.restore('Объём1','A1');assert.equal(m.get('Объём1','A3'),10);
});
test('circular references and missing external sources are explicit errors',()=>{
 const d=sample();d.sheets['Объём1'].D1={f:'D2+1'};d.sheets['Объём1'].D2={f:'D1+1'};d.sheets['Объём1'].D3={f:'[3]!missing_price'};
 const m=new Model(d);assert.ok(m.get('Объём1','D1') instanceof ExcelError);assert.equal(m.get('Объём1','D3').code,'#N/A');
});
test('structured table current row, body totals and total row translate to cell references',()=>{
 const d=sample();d.tables={Table5:{sheet:'Объём1',ref:'A3:C8',columns:['Материал','Итого, руб.','Себестоимость материалов '],totals:1}};
 assert.equal(normalizeTables('Table5[[#This Row],[Материал]]',d,'Объём1','A6'),"'Объём1'!A6");
 assert.equal(normalizeTables('Table5[[#Totals],[Итого, руб.]]',d,'Объём1','A6'),"'Объём1'!B8");
 assert.equal(normalizeTables('Table5[Итого, руб.]',d,'Объём1','A6'),"'Объём1'!B4:B7");
});
