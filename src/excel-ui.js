/* A19 private workbook editing surface: single user, local IndexedDB only. */
(() => {
  'use strict';
  const EX=window.A19Excel;
  let model=null,activeSheet='Объём1',tab='materials',loading=true,loadedLabel='',activeOrderId=null,pendingOrderId=null;
  let persistError='';
  const esc=x=>String(x??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
  const $=s=>document.querySelector(s);
  const fmt=x=>x instanceof EX.ExcelError?x.code:x===null||x===undefined?'—':typeof x==='number'?new Intl.NumberFormat('ru-RU',{maximumFractionDigits:2}).format(x):String(x);
  const rub=x=>x instanceof EX.ExcelError?x.code:typeof x==='number'?fmt(x)+' ₽':fmt(x);
  const numeric=x=>typeof x==='number'&&Number.isFinite(x);
  const key=s=>String(s||'').replace(/\$/g,'');
  const titles={
    'Объём1':'Объёмные элементы 1', 'Объём2':'Объёмные элементы 2', 'Объём3':'Объёмные элементы 3', 'Объём4':'Объёмные элементы 4',
    'Металлокаркас':'Металлокаркас', 'материалы':'Прайс материалов', 'услуги':'Прайс работ', 'КП':'Коммерческое предложение',
    'КП материалы и услуги накладные':'КП: полная экономика', 'КП материалы и услуги':'КП: детализация',
    'В производство':'Производственное задание', 'В производство1':'Производственное задание 2',
    'Сдельная зп':'Сдельная оплата', 'Работники':'Работники', 'Акт брака':'Акт брака', 'эскиз ':'Эскиз'
  };
  const sections={
    'Объём1':{mat:[24,66],lab:[70,108],stats:{sale:'L111',profit:'D121',materials:'I67',works:'I109',hours:'D122'}},
    'Объём2':{mat:[16,48],lab:[52,81],stats:{sale:'L84',profit:'D94',materials:'I49',works:'I82',hours:'D95'}},
    'Объём3':{mat:[16,48],lab:[52,81],stats:{sale:'L84',profit:'D94',materials:'I49',works:'I82',hours:'D95'}},
    'Объём4':{mat:[16,50],lab:[54,83],stats:{sale:'L86',profit:'D96',materials:'I51',works:'I84',hours:'D97'}},
    'Металлокаркас':{mat:[14,18],lab:[22,26],stats:{sale:'L29',profit:'D39',materials:'I19',works:'I27',hours:'D40'}}
  };
  const fields={
    'Объём1':[{key:'J10',name:'Ширина, мм'},{key:'J11',name:'Высота, мм'},{key:'C13',name:'Светодиоды, шт.'},{key:'C14',name:'Объёмные элементы, шт.'},{key:'C15',name:'Периметр, м'}],
    'Объём2':[{key:'C8',name:'Светодиоды, шт.'},{key:'C9',name:'Элементы, шт.'},{key:'C10',name:'Периметр, м'}],
    'Объём3':[{key:'C8',name:'Светодиоды, шт.'},{key:'C9',name:'Элементы, шт.'},{key:'C10',name:'Периметр, м'}],
    'Объём4':[{key:'C8',name:'Светодиоды, шт.'},{key:'C9',name:'Элементы, шт.'},{key:'C10',name:'Периметр, м'}],
    'Металлокаркас':[]
  };
  async function db(){return new Promise((resolve,reject)=>{const req=indexedDB.open('a19-excel-workbook',1);req.onupgradeneeded=()=>req.result.createObjectStore('private');req.onerror=()=>reject(req.error);req.onsuccess=()=>resolve(req.result);});}
  async function store(op,value){const idb=await db();return new Promise((resolve,reject)=>{const tx=idb.transaction('private',op==='read'?'readonly':'readwrite');const req=op==='read'?tx.objectStore('private').get('workbook'):tx.objectStore('private').put(value,'workbook');req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);tx.oncomplete=()=>idb.close();});}
  async function start(){try{const d=await store('read');if(d?.format==='a19-excel-model-v1')model=new EX.Model(d);}catch(e){persistError='Не удалось получить доступ к IndexedDB: '+e.message;}loading=false; if($('.nav-item.active')?.dataset.view==='excel')renderInto();}
  async function save(){try{await store('write',model.serialize());if(activeOrderId&&window.A19OrderBridge)window.A19OrderBridge.updateExcel(activeOrderId,model.overrides,snapshot());persistError='';}catch(e){persistError='Изменения не удалось сохранить: '+e.message;renderInto();}}
  const modelFile=async file=>{const raw=await file.text();const data=JSON.parse(raw);if(data.format!=='a19-excel-model-v1'||!data.sheets?.['Объём1']||!data.tables)throw new Error('Это не экспорт исходного калькулятора А19');return new EX.Model(data);};
  const card=(n,label)=>`<div class="excel-stat"><span class="mini-label">${esc(label)}</span><strong>${esc(rub(n))}</strong></div>`;
  function cell(sheet,ref,editable=false,style=''){
    const entry=model.data.sheets[sheet]?.[ref],isFormula=!!entry?.f,hasOver=Object.prototype.hasOwnProperty.call(model.overrides,sheet+'!'+ref);
    const val=model.get(sheet,ref);
    const v=val instanceof EX.ExcelError?val.code:val;
    if(!editable)return `<span class="${val instanceof EX.ExcelError?'excel-error':''}">${esc(fmt(v))}</span>`;
    if(!hasOver&&isFormula)return `<span class="excel-auto" title="Расчёт по формуле">${esc(fmt(v))}</span><button class="mini-icon" title="Установить вручную" data-excel-action="override" data-cell="${ref}" data-sheet="${esc(sheet)}">✎</button>`;
    const set=hasOver?model.overrides[sheet+'!'+ref]:v;
    return `<div class="excel-input-wrap"><input data-excel-value="${ref}" data-excel-sheet="${esc(sheet)}" value="${esc(set??'')}" aria-label="${esc(ref)}" ${style}/>${hasOver?`<button class="mini-icon" title="Вернуть формулу" data-excel-action="restore" data-cell="${ref}" data-sheet="${esc(sheet)}">↺</button>`:''}</div>`;
  }
  function lineTable(sheet,kind){const conf=sections[sheet];const range=kind==='materials'?conf.mat:conf.lab;
    const entries=[];
    for(let r=range[0];r<=range[1];r++){
      const section=model.get(sheet,'B'+r),name=model.get(sheet,'C'+r);
      if(typeof section!=='string'&&typeof name!=='string')continue;
      if(!name||!String(name).trim())continue;
      const cols=kind==='materials'?[['D','Цена продажи'],['G','Расход'],['H','Включить'],['I','Сумма'],['K','Себестоимость']]:[['D','Стоимость'],['G','Норма'],['H','Количество'],['I','Сумма'],['J','Время, мин'],['K','Себестоимость']];
      entries.push(`<tr><td class="sheet-ref">${r}</td><td><span class="muted">${esc(section)}</span><strong class="excel-name">${esc(name)}</strong></td>${cols.map(([c])=>`<td>${cell(sheet,c+r,true)}</td>`).join('')}</tr>`);
    }
    const cols=kind==='materials'?['Цена продажи','Расход','Коэффициент','Сумма','Себестоимость']:['Стоимость','Норма','Количество','Сумма','Время, мин','Себестоимость'];
    return `<div class="excel-table-wrap"><table class="excel-table"><thead><tr><th>Стр.</th><th>Позиция из Excel</th>${cols.map(x=>`<th>${esc(x)}</th>`).join('')}</tr></thead><tbody>${entries.join('')}</tbody></table></div>`;
  }
  function genericTable(sheet){const cells=model.data.sheets[sheet]||{},rows=new Map();
    for(const [ref,e] of Object.entries(cells)){const r=Number(ref.match(/\d+/)?.[0]);const col=ref.match(/^[A-Z]+/)?.[0];if(!r||!col)continue;if(!rows.has(r))rows.set(r,[]);rows.get(r).push({ref,col,e});}
    const list=[...rows].sort((a,b)=>a[0]-b[0]);
    return `<div class="excel-table-wrap"><table class="excel-table"><thead><tr><th>Стр.</th><th>Значения и формулы</th></tr></thead><tbody>${list.map(([r,arr])=>`<tr><td class="sheet-ref">${r}</td><td><div class="excel-cell-list">${arr.map(({ref,col,e})=>`<div class="excel-cell"><label>${col}${e.f?' · fx':''}</label>${cell(sheet,ref,true)}</div>`).join('')}</div></td></tr>`).join('')}</tbody></table></div>`;
  }
  function snapshot(){
    const value=model.get('КП','N29');const fallback=model.get('Объём1','L111');
    return {total:typeof value==='number'?value:typeof fallback==='number'?fallback:0,checkedAt:new Date().toISOString(),errors:model.details.size};
  }
  function openOrderId(id){
    pendingOrderId=id;
    if(!model)return;
    const o=window.A19OrderBridge?.getExcel(id);if(!o)return;
    model.overrides={...(o.excelOverrides||{})};model.clear();activeOrderId=id;pendingOrderId=null;activeSheet='Объём1';tab='materials';
  }
  function chooseNew(){activeOrderId=null;pendingOrderId=null;if(model){model.overrides={};model.clear();save().then(renderInto);}}
  function saveAsOrder(){
    if(!model)return;
    const customer=prompt('Заказчик для нового расчёта Excel (можно оставить пустым):','');
    if(customer===null)return;
    activeOrderId=window.A19OrderBridge?.newExcel(customer,model.overrides,snapshot())||null;
    renderInto();
  }
  function page(){
    if(loading)return `<div class="page-heading"><h1>Загрузка расчётной модели</h1></div>`;
    if(!model)return `<div class="page-heading"><div><div class="mini-label">А19 / ПРОИЗВОДСТВЕННЫЙ РАСЧЁТ</div><h1>Калькулятор Excel</h1><p>Исходные формулы и прайс загружаются только в этот браузер.</p></div></div>
    <section class="card excel-hero"><div class="mini-label">ЛОКАЛЬНОЕ РАБОЧЕЕ МЕСТО · 1 МЕНЕДЖЕР</div><h2>Полная модель исходной книги</h2><p>16 листов, 4 115 формул, связанные цены материалов и производственные операции. Для работы загрузите приватный файл <code>A19_excel_model_PRIVATE.json</code>. Он не размещается в публичном GitHub и не передаётся на сервер.</p><button class="primary" data-excel-action="import">Загрузить расчётную модель</button><p class="excel-footnote">Данные хранятся на этом устройстве. Резервную копию следует сохранять отдельно. При смене устройства файл импортируется повторно.</p></section>${persistError?`<p class="excel-error">${esc(persistError)}</p>`:''}`;
    const orders=window.A19OrderBridge?.excelOrders()||[];
    const orderActions=`<section class="card excel-section"><div class="section-title"><h2>Расчёты менеджера</h2><small>${orders.length} сохранено в этом браузере</small></div><div class="excel-order-bar"><select id="excel-order-select"><option value="">Текущий черновик без заказа</option>${orders.map(o=>`<option value="${esc(o.id)}" ${activeOrderId===o.id?'selected':''}>№ ${o.number} · ${esc(o.customer||'Без имени')} · ${esc(rub(o.total))}</option>`).join('')}</select><button class="secondary" data-excel-action="new">Новый расчёт</button><button class="primary" data-excel-action="save-order">${activeOrderId?'Сохранено в заказе':'Сохранить как заказ'}</button></div></section>`;
    const audit=model.audit();const manualCount=Object.keys(model.overrides).length;const names=model.data.sheetOrder||Object.keys(model.data.sheets);
    const config=sections[activeSheet],stats=config?.stats;
    const statsUI=stats?`<div class="excel-stats">${card(model.get(activeSheet,stats.sale),'Итого по Excel')}${card(model.get(activeSheet,stats.materials),'Материалы / продажа')}${card(model.get(activeSheet,stats.works),'Работы / продажа')}${card(model.get(activeSheet,stats.profit),'Плановая прибыль')}</div>`:'';
    const inputFields=(fields[activeSheet]||[]).map(({key:ref,name})=>`<div class="field"><label>${esc(name)} · ${ref}</label>${cell(activeSheet,ref,true)}</div>`).join('');
    const tabs=config?`<div class="excel-tabs"><button data-excel-action="tab" data-tab="materials" class="${tab==='materials'?'current':''}">Материалы</button><button data-excel-action="tab" data-tab="services" class="${tab==='services'?'current':''}">Производственные операции</button><button data-excel-action="tab" data-tab="all" class="${tab==='all'?'current':''}">Все строки</button></div>`:'';
    const core=config?(tab==='materials'?lineTable(activeSheet,'materials'):tab==='services'?lineTable(activeSheet,'services'):genericTable(activeSheet)):genericTable(activeSheet);
    return `<div class="page-heading"><div><div class="mini-label">A19 / ФОРМУЛЬНАЯ МОДЕЛЬ · ЛОКАЛЬНО</div><h1>Калькулятор Excel</h1><p>Формулы воспроизводятся из исходной книги, ручные правки не затираются при пересчёте.</p></div><div class="toolbar"><button class="secondary" data-excel-action="export">Резервная копия модели</button><button class="secondary" data-excel-action="import">Загрузить другую книгу</button></div></div>
    <div class="excel-audit"><strong>${audit.ok} из ${audit.formulas} формул вычислены</strong><span>${audit.error} ошибок / отсутствующих значений · ${manualCount?manualCount+" ручных корректировок (сравнение с исходным кэшем не применяется)":audit.diff+" числовых расхождений с исходным кэшем"}</span></div>
    ${persistError?`<div class="warning">${esc(persistError)}</div>`:''}
    <section class="card excel-section"><div class="section-title"><h2>Раздел расчёта</h2><small>${esc(activeSheet)}</small></div><div class="field"><label>Лист исходного Excel</label><select id="excel-sheet">${names.map(n=>`<option value="${esc(n)}" ${n===activeSheet?'selected':''}>${esc(titles[n]||n)}</option>`).join('')}</select></div></section>
    ${orderActions}
    ${statsUI}
    ${inputFields?`<section class="card excel-section"><div class="section-title"><h2>Параметры конструкции</h2><small>Данные из исходного листа</small></div><div class="excel-params">${inputFields}</div></section>`:''}
    <section class="card excel-section"><div class="section-title"><h2>${esc(titles[activeSheet]||activeSheet)}</h2><small>Суммы пересчитываются при изменении полей</small></div>${tabs}${core}</section>
    <p class="excel-footnote">Числа и правила взяты из приложенного Excel. Формулы с отсутствующими внешними значениями показывают ошибки, а не предполагаемые цены. Публичный сайт не хранит исходный прайс, заказы или токены.</p>`;
  }
  function renderInto(){if($('#app'))$('#app').innerHTML=page();}
  function download(){const blob=new Blob([JSON.stringify(model.serialize())],{type:'application/json;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download='A19_excel_model_backup_'+new Date().toISOString().slice(0,10)+'.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);}
  document.addEventListener('click',async e=>{const b=e.target.closest('[data-excel-action]');if(!b)return;
    const act=b.dataset.excelAction;
    if(act==='import'){$('#excel-model-input').click();return;}
    if(act==='export'&&model){download();return;}
    if(act==='save-order'){if(activeOrderId){await save();renderInto();}else saveAsOrder();return;}
    if(act==='new'){if(confirm('Начать новый расчёт? Текущий несохранённый черновик заменится исходными данными Excel.'))chooseNew();return;}
    if(act==='tab'){tab=b.dataset.tab;renderInto();return;}
    if(act==='override'&&model){const sheet=b.dataset.sheet,ref=b.dataset.cell;const val=model.get(sheet,ref);model.edit(sheet,ref,val instanceof EX.ExcelError?'':val);await save();renderInto();return;}
    if(act==='restore'&&model){model.restore(b.dataset.sheet,b.dataset.cell);await save();renderInto();}
  });
  document.addEventListener('change',async e=>{
    if(e.target.id==='excel-model-input'){const file=e.target.files?.[0];e.target.value='';if(!file)return;
      try{const next=await modelFile(file);if(model&&!confirm('Заменить загруженную модель? Сделайте резервную копию текущего расчёта.'))return;model=next;activeOrderId=null;if(pendingOrderId)openOrderId(pendingOrderId);await save();activeSheet='Объём1';tab='materials';renderInto();}
      catch(err){alert('Не удалось импортировать модель: '+err.message);}return;
    }
    if(!model)return;
    if(e.target.id==='excel-order-select'){if(e.target.value){openOrderId(e.target.value);renderInto();}else chooseNew();return;}
    if(e.target.id==='excel-sheet'){activeSheet=e.target.value;tab='materials';renderInto();return;}
    if(e.target.matches('[data-excel-value]')){
      const sheet=e.target.dataset.excelSheet,ref=key(e.target.dataset.excelValue);let val=e.target.value.trim();
      if(val!==''&&/^[+\-]?(?:\d+(?:[.,]\d*)?|[.,]\d+)$/.test(val))val=Number(val.replace(',','.'));
      model.edit(sheet,ref,val);await save();renderInto();
    }
  });
  window.A19ExcelUI={render:page,renderInto,openOrderId};start();
})();
