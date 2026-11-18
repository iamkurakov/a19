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
  async function store(op,value,storageKey='workbook'){const idb=await db();return new Promise((resolve,reject)=>{const tx=idb.transaction('private',op==='read'?'readonly':'readwrite');const req=op==='read'?tx.objectStore('private').get(storageKey):tx.objectStore('private').put(value,storageKey);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);tx.oncomplete=()=>idb.close();});}
  async function start(){try{const d=await store('read');if(d?.format==='a19-excel-model-v1')model=new EX.Model(d);const session=await store('read',null,'session');if(model&&session?.activeOrderId)openOrderId(session.activeOrderId);}catch(e){persistError='Не удалось получить доступ к IndexedDB: '+e.message;}loading=false;renderInto();}
  async function save(){try{await store('write',model.serialize());if(activeOrderId&&window.A19OrderBridge)window.A19OrderBridge.updateExcel(activeOrderId,model.overrides,snapshot());persistError='';}catch(e){persistError='Изменения не удалось сохранить: '+e.message;renderInto();}}
  const workspaceData=()=>model?model.serialize():null;
  async function ensurePrivateModel(data){
    if(model)return false;
    if(!data||data.format!=='a19-excel-model-v1'||!data.sheets?.['Объём1'])throw new Error('Некорректная облачная модель Excel');
    model=new EX.Model(data);
    if(pendingOrderId)openOrderId(pendingOrderId);
    await save();loading=false;renderInto();return true;
  }
  function confirmedExternal(){
    if(!model)return [];
    const result=[];
    for(const [book,sheets] of Object.entries(model.data.external||{}))for(const [sheet,cells] of Object.entries(sheets||{}))for(const [cell,x] of Object.entries(cells||{}))if(x?.manuallyConfirmed===true)result.push({book,sheet,cell,value:x.v});
    return result;
  }
  async function applyConfirmedExternal(values){
    if(!model||!Array.isArray(values))return;
    let modified=false;
    for(const item of values){
      if(!item||!/^\d+$/.test(String(item.book))||!/^[A-Z]+\d+$/.test(item.cell)||typeof item.sheet!=='string')continue;
      if(model.data.external?.[item.book]?.[item.sheet]?.[item.cell]?.manuallyConfirmed)continue;
      model.setExternal(item.book,item.sheet,item.cell,item.value);modified=true;
    }
    if(modified){await save();renderInto();}
  }
  async function restoreWorkspace(data){
    if(data!==null && (data.format!=='a19-excel-model-v1'||!data.sheets?.['Объём1']))throw new Error('Неверная модель Excel');
    if(data){const next=new EX.Model(data);await store('write',next.serialize());model=next;}
    activeOrderId=null;await store('write',{activeOrderId:null},'session');
    persistError='';activeSheet='Объём1';tab='materials';renderInto();
  }
  const modelFile=async file=>{const raw=await file.text();const data=JSON.parse(raw);if(data.format!=='a19-excel-model-v1'||!data.sheets?.['Объём1']||!data.tables)throw new Error('Это не экспорт исходного калькулятора А19');return new EX.Model(data);};
  const card=(n,label)=>`<div class="excel-stat"><span class="mini-label">${esc(label)}</span><strong>${esc(rub(n))}</strong><small class="metric-disclaimer">Предварительно · для сверки</small></div>`;
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
    const estimate=typeof value==='number'?value:typeof fallback==='number'?fallback:null;
    const audit=model.detailedAudit();
    // Never mark the client-facing total as verified while source formulas or references have errors.
    return {total:null,estimate,verified:false,usage:'testing-only',checkedAt:new Date().toISOString(),errors:audit.summary.error,sourceErrors:audit.errors.filter(e=>e.originalError).length,missingExternal:audit.errors.filter(e=>!e.originalError).length};
  }
  function openOrderId(id){
    pendingOrderId=id;
    if(!model)return;
    const o=window.A19OrderBridge?.getExcel(id);if(!o)return;
    model.overrides={...(o.excelOverrides||{})};model.clear();activeOrderId=id;pendingOrderId=null;activeSheet='Объём1';tab='materials';store('write',{activeOrderId:id},'session').catch(()=>{});
  }
  function chooseNew(){activeOrderId=null;pendingOrderId=null;store('write',{activeOrderId:null},'session').catch(()=>{});if(model){model.overrides={};model.clear();save().then(renderInto);}}
  function saveAsOrder(){
    if(!model)return;
    const customer=prompt('Заказчик для нового расчёта Excel (можно оставить пустым):','');
    if(customer===null)return;
    activeOrderId=window.A19OrderBridge?.newExcel(customer,model.overrides,snapshot())||null;
    if(activeOrderId)store('write',{activeOrderId},'session').catch(()=>{});renderInto();
  }
  function diagnosticView(){
    const report=model.detailedAudit(),all=report.errors;
    const original=all.filter(x=>x.originalError),missing=all.filter(x=>!x.originalError&&/No external cache for/.test(x.details));
    const ext=new Map();
    for(const item of missing){
      const m=item.details.match(/No external cache for (\d+):(.+)!([A-Z]+\d+)/);
      if(m){const id=m[1]+'|'+m[2]+'|'+m[3];if(!ext.has(id))ext.set(id,{book:m[1],sheet:m[2],cell:m[3]});}
    }
    const groups=[['#REF!','Повреждённые ссылки'],['#DIV/0!','Деление на ноль'],['#VALUE!','Ошибки значений'],['#N/A','Отсутствующие данные']];
    const counts=groups.map(([code,title])=>`<div class="diag-count"><strong>${all.filter(x=>x.code===code).length}</strong><span>${title}</span></div>`).join('');
    const inputs=[...ext.values()].sort((a,b)=>a.sheet.localeCompare(b.sheet,'ru')||a.cell.localeCompare(b.cell,'ru',{numeric:true})).map(x=>{
      const source=x.sheet==='Стоимость трудочасов'?'Оплата труда / '+x.cell:'Прайс материалов / '+x.cell;
      return `<div class="diag-input"><div><strong>${esc(source)}</strong><small>Внешняя книга № ${esc(x.book)} · ${esc(x.sheet)}!${esc(x.cell)}</small></div><input data-excel-external-book="${esc(x.book)}" data-excel-external-sheet="${esc(x.sheet)}" data-excel-external-cell="${esc(x.cell)}" aria-label="Значение для ${esc(source)}" placeholder="Укажите подтверждённое значение" autocomplete="off"/><button class="secondary" data-excel-action="set-external" data-excel-book="${esc(x.book)}" data-excel-external-sheet="${esc(x.sheet)}" data-excel-cell="${esc(x.cell)}">Применить</button></div>`;}).join('');
    const items=all.map(x=>{
      const origin=x.originalError?'Ошибка исходного Excel':'Нет значения во внешнем кэше';
      const cause=x.details||([['#REF!','Повреждена ссылка в оригинале'],['#DIV/0!','Зависимые расчёты при нулевом количестве'],['#VALUE!','Ошибка исходных формул / значений']].find(([code])=>code===x.code)?.[1]||'');
      return `<tr><td><strong>${esc(x.sheet)}!${esc(x.cell)}</strong></td><td class="excel-error">${esc(x.code)}</td><td>${esc(origin)}</td><td>${esc(cause)}</td><td><button class="secondary diag-goto" data-excel-action="goto" data-excel-sheet="${esc(x.sheet)}" data-excel-cell="${esc(x.cell)}">Открыть</button></td></tr>`;
    }).join('');
    return `<div class="page-heading"><div><div class="mini-label">А19 / КОНТРОЛЬ РАСЧЁТА</div><h1>Диагностика формул</h1><p>Показаны все ошибки из импортированной Excel-модели. Проверка не заменяет согласование себестоимости.</p></div><button class="secondary" data-excel-action="export-diagnostic">Скачать список ошибок</button></div>
    <div class="diag-counts">${counts}</div><div class="warning diag-warning"><strong>${report.summary.error} проблемных формул.</strong> ${original.length} — из исходной книги, ${missing.length} — без внешних данных. Незаполненные листы 2–4 порождают часть взаимосвязанных ошибок.</div>
    <section class="card"><div class="section-title"><h2>Недостающие исходные данные</h2><small>${ext.size} уникальных ячеек внешней книги</small></div><p class="excel-footnote">Подставлять нули вместо неизвестных ставок нельзя. Поля ниже позволяют внести подтверждённые значения из исходного прайса. Они сохраняются в браузере и автоматически отправляются в приватный справочник при подключённом облаке.</p>${inputs||'<p>Все доступные значения внешних ссылок заполнены.</p>'}</section>
    <section class="card"><div class="section-title"><h2>Список ошибок</h2><small>${all.length} формул</small></div><div class="excel-table-wrap"><table class="excel-table diag-table"><thead><tr><th>Лист и ячейка</th><th>Код</th><th>Источник</th><th>Причина</th><th></th></tr></thead><tbody>${items}</tbody></table></div></section>`;
  }
  function diagnosticDownload(){
    const errors=model.detailedAudit().errors;const rows=[['Лист','Ячейка','Ошибка','Причина','Исходный Excel','Формула'],...errors.map(x=>[x.sheet,x.cell,x.code,x.details||'',x.originalError?'Да':'Нет',x.formula])];
    const csv=rows.map(row=>row.map(x=>'"'+String(x??'').replace(/"/g,'""')+'"').join(';')).join('\r\n');
    const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'});const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download='A19_errors_'+new Date().toISOString().slice(0,10)+'.csv';a.click();setTimeout(()=>URL.revokeObjectURL(u),1500);
  }
  function printDraft(){
    const v=model.get('КП','N29');
    if(!(typeof v==='number'&&Number.isFinite(v))){alert('Итоговая сумма КП содержит ошибку. Сначала проверьте формулы.');return;}
    const o=activeOrderId?window.A19OrderBridge?.getExcel(activeOrderId):null;
    const fieldsList=fields['Объём1'].map(x=>`<tr><td>${esc(x.name)}</td><td>${esc(fmt(model.get('Объём1',x.key)))}</td></tr>`).join('');
    $('#print-sheet').innerHTML=`<img src="assets/logo.svg" alt="A19"><h1>ТЕСТОВЫЙ РАСЧЁТ № ${esc(o?.number||'—')}</h1><p>Клиент: ${esc(o?.customer?.name||'Не указан')}<br>Дата: ${new Date().toLocaleDateString('ru-RU')}</p><h2>Параметры конструкции</h2><table>${fieldsList}</table><div class="print-footer">Расчётная стоимость: ${esc(rub(v))}</div><p style="margin-top:32px;font-weight:700">НЕ ДЛЯ ОПЛАТЫ И НЕ ДЛЯ КЛИЕНТА. Расчёт в режиме тестирования. Исходная модель содержит нерешённые ошибки и недостающие справочные данные.</p>`;
    window.print();
  }
  function page(){
    if(loading)return `<div class="page-heading"><h1>Загрузка расчётной модели</h1></div>`;
    if(!model)return `<div class="page-heading"><div><div class="mini-label">А19 / ПРОИЗВОДСТВЕННЫЙ РАСЧЁТ</div><h1>Калькулятор Excel</h1><p>Исходные формулы и прайс загружаются только в этот браузер.</p></div></div>
    <section class="card excel-hero"><div class="mini-label">ОДИН МЕНЕДЖЕР · ПРИВАТНОЕ ХРАНИЛИЩЕ А19</div><h2>Подключение рабочего калькулятора</h2><p>Модель уже находится в закрытом репозитории <code>iamkurakov/a19_data</code>. При первом подключении GitHub система сама загружает цены и все 4 115 формул. Повторно скачивать Excel или JSON не нужно.</p><div class="cloud-entry-buttons"><button class="primary" data-cloud-action="open">Подключить приватную базу</button><button class="secondary" data-excel-action="import">Загрузить JSON вручную</button></div><p class="excel-footnote">Для облачного подключения используется отдельный токен GitHub с доступом только к <code>a19_data</code>. Вход в А19 не требуется. Без подключения данные доступны только на этом устройстве.</p></section><section class="card excel-how"><h2>Как сейчас работает А19</h2><p>После подключения вводите размеры конструкции, выбираете материалы и работы, при необходимости меняете цифры вручную. Система пересчитывает результат, а сохранённые заказы отправляет в закрытый репозиторий. Это тестовый расчёт: окончательная цена требует проверки ошибок исходного Excel.</p><h3>Что ещё уточняем в формулах</h3><p>54 ошибки пришли из самой книги (в основном незаполненные блоки 2–4), ещё 36 связаны с отсутствующими ставками внешнего справочника. Нули вместо неизвестных цен не подставляем.</p></section>${persistError?`<p class="excel-error">${esc(persistError)}</p>`:''}`;
    if(activeSheet==='@diag')return diagnosticView();
    const orders=window.A19OrderBridge?.excelOrders()||[];
    const orderActions=`<section class="card excel-section"><div class="section-title"><h2>Расчёты менеджера</h2><small>${orders.length} заказов · автоматическая отправка при подключённом облаке</small></div><div class="excel-order-bar"><select id="excel-order-select"><option value="">Текущий черновик без заказа</option>${orders.map(o=>`<option value="${esc(o.id)}" ${activeOrderId===o.id?'selected':''}>№ ${o.number} · ${esc(o.customer||'Без имени')} · ${esc(rub(o.estimate))} · тестовый</option>`).join('')}</select><button class="secondary" data-excel-action="new">Новый расчёт</button><button class="primary" data-excel-action="save-order">${activeOrderId?'Сохранить изменения':'Сохранить тестовый расчёт'}</button></div></section>`;
    const report=model.detailedAudit(),audit=report.summary;const manualCount=Object.keys(model.overrides).length;const names=model.data.sheetOrder||Object.keys(model.data.sheets);
    const sourceCount=report.errors.filter(e=>e.originalError).length;
    const externalCount=report.errors.length-sourceCount;
    const testAdvice=`<section class="test-panel" role="status" aria-live="polite"><div class="test-panel-main"><div class="test-panel-icon">!</div><div><div class="test-panel-caption">РЕЖИМ ПРОВЕРКИ · ЦЕНА НЕ УТВЕРЖДЕНА</div><h2>Расчёты доступны для тестирования</h2><p>Меняйте размеры, состав и цены, сохраняйте отдельные заказы и сравнивайте результат с Excel. Сейчас найдено <strong>${audit.error}</strong> проблемных формул: ${sourceCount} из исходной книги и ${externalCount} без внешних данных. Итоги и прибыль показываются <strong>только предварительно</strong>, не для выставления окончательного КП.</p></div></div><div class="test-panel-actions"><button class="secondary" data-excel-action="diagnostics">Разобрать ошибки</button><span class="test-panel-state">${manualCount} ручных корректировок</span></div></section>`;
    const config=sections[activeSheet],stats=config?.stats;
    const statsUI=stats?`<div class="excel-stats">${card(model.get(activeSheet,stats.sale),'Итого по Excel')}${card(model.get(activeSheet,stats.materials),'Материалы / продажа')}${card(model.get(activeSheet,stats.works),'Работы / продажа')}${card(model.get(activeSheet,stats.profit),'Плановая прибыль')}</div>`:'';
    const inputFields=(fields[activeSheet]||[]).map(({key:ref,name})=>`<div class="field"><label>${esc(name)} · ${ref}</label>${cell(activeSheet,ref,true)}</div>`).join('');
    const tabs=config?`<div class="excel-tabs"><button data-excel-action="tab" data-tab="materials" class="${tab==='materials'?'current':''}">Материалы</button><button data-excel-action="tab" data-tab="services" class="${tab==='services'?'current':''}">Производственные операции</button><button data-excel-action="tab" data-tab="all" class="${tab==='all'?'current':''}">Все строки</button></div>`:'';
    const core=config?(tab==='materials'?lineTable(activeSheet,'materials'):tab==='services'?lineTable(activeSheet,'services'):genericTable(activeSheet)):genericTable(activeSheet);
    return `<div class="page-heading"><div><div class="mini-label">A19 / ФОРМУЛЬНАЯ МОДЕЛЬ · ЛОКАЛЬНО</div><h1>Расчёт конструкции</h1><p>Материалы, производственные операции, металлокаркас и экономика в одном рабочем пространстве.</p></div><div class="toolbar"><button class="secondary" data-excel-action="print">Распечатать тестовый расчёт</button><button class="secondary" data-excel-action="diagnostics">Ошибки расчёта</button><button class="secondary" data-excel-action="export">Резервная копия модели</button><button class="secondary" data-excel-action="import">Загрузить другую книгу</button></div></div>
    <div class="excel-audit"><strong>${audit.ok} / ${audit.formulas} формул вычислены</strong><span>${audit.error} требуют проверки · ${manualCount?manualCount+" ручных правок":audit.diff+" расхождений с исходным кэшем"}</span></div>${testAdvice}
    <section class="card excel-how"><div class="section-title"><h2>Как пользоваться расчётом</h2><small>Краткая инструкция</small></div><p>1. Выберите тип элемента и задайте параметры. 2. Проверьте материалы и работы — стоимость меняется автоматически. 3. Любое значение можно изменить вручную. 4. Сохраните расчёт как заказ. Когда облако подключено, заказ автоматически отправляется в закрытый репозиторий.</p><p><strong>Что с ошибками:</strong> ${sourceCount} исходных ошибок Excel и ${externalCount} недостающих внешних значений. Многие ошибки относятся к пустым группам элементов 2–4. Расчёт можно тестировать, но цена и прибыль не утверждены. <button class="inline-link" data-excel-action="diagnostics">Посмотреть причины и адреса ячеек</button></p></section>
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
    if(act==='diagnostics'&&model){activeSheet='@diag';renderInto();return;}
    if(act==='export-diagnostic'&&model){diagnosticDownload();return;}
    if(act==='goto'&&model){activeSheet=b.dataset.excelSheet;tab='all';renderInto();return;}
    if(act==='print'&&model){printDraft();return;}
    if(act==='set-external'&&model){
      const book=b.dataset.excelBook,sheet=b.dataset.excelExternalSheet,cell=b.dataset.excelCell;
      const input=[...document.querySelectorAll('[data-excel-external-cell]')].find(x=>x.dataset.excelExternalBook===book&&x.dataset.excelExternalSheet===sheet&&x.dataset.excelExternalCell===cell);
      if(!input)return;const raw=input.value.trim();if(!raw){alert('Введите подтверждённое значение из исходного справочника.');return;}
      const isNumber=/^[BJ]/.test(cell);let value=raw;
      if(isNumber){value=Number(raw.replace(',','.'));if(!Number.isFinite(value)||value<0){alert('Для цены или ставки требуется неотрицательное число.');return;}}
      if(!confirm('Подтвердить значение «'+raw+'» для '+sheet+'!'+cell+'? Оно будет внесено в локальный справочник А19.'))return;
      model.setExternal(book,sheet,cell,value);await save();window.A19Cloud?.markExternal();renderInto();return;
    }
    if(act==='export'&&model){download();return;}
    if(act==='save-order'){if(activeOrderId){await save();renderInto();}else saveAsOrder();return;}
    if(act==='new'){if(confirm('Начать новый расчёт? Текущий несохранённый черновик заменится исходными данными Excel.'))chooseNew();return;}
    if(act==='tab'){tab=b.dataset.tab;renderInto();return;}
    if(act==='override'&&model){const sheet=b.dataset.sheet,ref=b.dataset.cell;const val=model.get(sheet,ref);model.edit(sheet,ref,val instanceof EX.ExcelError?'':val);await save();renderInto();return;}
    if(act==='restore'&&model){model.restore(b.dataset.sheet,b.dataset.cell);await save();renderInto();}
  });
  document.addEventListener('change',async e=>{
    if(e.target.id==='excel-model-input'){const file=e.target.files?.[0];e.target.value='';if(!file)return;
      try{const next=await modelFile(file);if(model&&!confirm('Заменить загруженную модель? Сделайте резервную копию текущего расчёта.'))return;model=next;activeOrderId=null;await store('write',{activeOrderId:null},'session');if(pendingOrderId)openOrderId(pendingOrderId);await save();activeSheet='Объём1';tab='materials';renderInto();}
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
  window.A19ExcelUI={render:page,renderInto,openOrderId,chooseNew,workspaceData,restoreWorkspace,ensurePrivateModel,hasModel:()=>!!model,confirmedExternal,applyConfirmedExternal,ready:null,onBridgeReady:()=>{if(pendingOrderId)openOrderId(pendingOrderId);},setSheet:(sheet)=>{activeSheet=sheet;tab='materials';},renderSheet:(sheet)=>{if(activeSheet!==sheet){activeSheet=sheet;tab='all';}return page();},diagnostics:()=>{activeSheet='@diag';return page();}};window.A19ExcelUI.ready=start();
})();
