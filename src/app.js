/* A19 browser-only prototype. Do not connect this module to Ideal Potolok or production secrets. */
(() => {
  'use strict';
  const E = window.A19Engine;
  const STORAGE = 'a19.prototype.v1';
  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const rub = (n) => new Intl.NumberFormat('ru-RU',{maximumFractionDigits:2}).format(E.money(n))+' ₽';
  const num = (n) => new Intl.NumberFormat('ru-RU',{maximumFractionDigits:4}).format(E.amount(n));
  const uid = () => crypto?.randomUUID?.() ?? 'a19-'+Date.now()+'-'+Math.random().toString(36).slice(2);
  const clone = (x) => JSON.parse(JSON.stringify(x));
  const blank = () => ({schema:1,orders:[],catalog:[]});
  const storage = () => {try { const s=JSON.parse(localStorage.getItem(STORAGE)); return s&&s.schema===1&&Array.isArray(s.orders)&&Array.isArray(s.catalog)?s:blank(); }catch(e){return blank();}};
  let db = storage();
  let view = 'orders';
  let selected = null;
  const persist = () => { try {localStorage.setItem(STORAGE, JSON.stringify(db));} catch(e){alert('Не удалось сохранить данные в браузере. Скачайте резервную копию и освободите место.');} };
  const save = () => {persist();render();};
  const byId = id => db.orders.find(x=>x.id===id);
  const objById = (order,id) => order?.objects.find(x=>x.id===id);
  const itemById = id => db.catalog.find(x=>x.id===id);
  const statusNames = {draft:'Черновик',sent:'КП отправлено',approved:'Согласовано',production:'В производстве',done:'Завершено',declined:'Отказ'};
  const makeObject = (n) => ({id:uid(),title:'Рекламный объект '+n,widthMm:'',heightMm:'',count:1,lines:[],note:''});
  function makeOrder(){
    const sequence = Math.max(0,...db.orders.map(x=>Number(x.number)||0))+1;
    return {id:uid(),number:sequence,created:new Date().toISOString(),updated:new Date().toISOString(),status:'draft',customer:{name:'',phone:'',address:''},manager:'',discountPct:0,otherExpense:0,note:'',objects:[makeObject(1)]};
  }
  const date = value => value ? new Date(value).toLocaleDateString('ru-RU') : '—';
  const field = (label,value,path,scope,ref='',extra='') => `<div class="field"><label>${esc(label)}</label><input data-edit="${esc(path)}" data-scope="${scope}" data-ref="${esc(ref)}" value="${esc(value ?? '')}" ${extra}></div>`;
  const selectField = (label,current,opts,path,scope,ref='') => `<div class="field"><label>${esc(label)}</label><select data-edit="${esc(path)}" data-scope="${scope}" data-ref="${esc(ref)}">${opts.map(([v,n])=>`<option value="${esc(v)}" ${String(v)===String(current)?'selected':''}>${esc(n)}</option>`).join('')}</select></div>`;
  const btn = (text,action,css='secondary',extra='') => `<button class="${css}" data-action="${action}" ${extra}>${text}</button>`;
  const statuses = Object.entries(statusNames);
  function header(title,sub,action='') { return `<div class="page-heading"><div><div class="mini-label">A19 / РАБОЧАЯ СИСТЕМА</div><h1>${esc(title)}</h1><p>${esc(sub)}</p></div>${action}</div>`; }
  function ordersView(){
    const total = db.orders.reduce((s,o)=>s+E.calcOrder(o,db.catalog).total,0);
    const active = db.orders.filter(o=>!['done','declined'].includes(o.status)).length;
    const done = db.orders.filter(o=>o.status==='done').length;
    return `${header('Заказы','Коммерческие расчёты и рекламные конструкции.',btn('＋ Новый расчёт','new-order','primary'))}
      <div class="stats"><div class="stat"><div class="mini-label">ВСЕГО РАСЧЁТОВ</div><div class="metric">${db.orders.length}</div></div><div class="stat"><div class="mini-label">В РАБОТЕ</div><div class="metric">${active}</div></div><div class="stat"><div class="mini-label">СУММА РАСЧЁТОВ / ВСЕ СТАТУСЫ</div><div class="metric">${rub(total)}</div></div></div>
      <section class="card"><div class="section-title"><h2>Реестр заказов</h2><small>Завершено: ${done}</small></div>
      ${db.orders.length ? `<div class="orders-list">${[...db.orders].reverse().map(o=>{
        const calc=E.calcOrder(o,db.catalog);return `<button class="order-row" data-action="open-order" data-id="${esc(o.id)}"><span class="order-id">№ ${o.number}<small>${date(o.created)}</small></span><span><strong>${esc(o.customer.name||'Клиент не указан')}</strong><small>${o.objects.length} объект(ов)</small></span><span class="hide-mid dim">${esc(o.manager||'Менеджер не указан')}</span><span><span class="tag">${esc(statusNames[o.status]||o.status)}</span><small>${rub(calc.total)}</small></span><span class="muted">Открыть →</span></button>`;}).join('')}</div>` : `<div class="empty"><strong>Пока нет расчётов</strong>Создайте первый заказ и добавьте материалы или работы из справочника.</div>`}</section>`;
  }
  function lineRow(order,obj,line,index,calculated){
    const item=itemById(line.itemId);
    const c=calculated.lines[index];
    const auto=['area','perimeter'].includes(line.driver||item?.driver);
    const label=auto ? `Авто: ${num(c.each)} ${esc(c.unit)} / объект` : `База: ${num(line.qty??1)} ${esc(c.unit)} / объект`;
    const k=`data-object="${esc(obj.id)}" data-line="${esc(index)}"`;
    return `<div class="line">
      <div><div class="line-name">${esc(c.name)}<small>${esc(c.category==='work'?'Работа':'Материал')} · ${label}</small></div>${line.fix?'<span class="badge-ok">Зафиксировано</span>':''}</div>
      <div class="field"><label>Кол-во / объект</label><input class="inline-input" data-scope="line" data-edit="qtyOverride" ${k} inputmode="decimal" placeholder="${esc(num(c.each))}" value="${esc(line.qtyOverride??'')}"></div>
      <div class="field"><label>Себестоим. / ед.</label><input class="inline-input" data-scope="line" data-edit="costOverride" ${k} inputmode="decimal" placeholder="${esc(num(line.fix?.unitCost??item?.unitCost??0))}" value="${esc(line.costOverride??'')}"></div>
      <div class="field"><label>Цена / ед.</label><input class="inline-input" data-scope="line" data-edit="priceOverride" ${k} inputmode="decimal" placeholder="${esc(num(line.fix?.unitPrice??item?.unitPrice??0))}" value="${esc(line.priceOverride??'')}"></div>
      <div class="field"><label>Итого кол-во</label><div class="line-total">${num(c.total)}</div></div>
      <div class="field"><label>Продажа</label><div class="line-total">${rub(c.sale)}</div></div>
      <button class="plain-close" title="Удалить позицию" data-action="delete-line" ${k}>×</button>
    </div>`;
  }
  function objectCard(order,obj,index,calc){
    const choices=db.catalog.map(i=>`<option value="${esc(i.id)}">${esc(i.category==='work'?'Работа':'Материал')} · ${esc(i.name)}</option>`).join('');
    const control = db.catalog.length ? `<div class="picker"><select id="pick-${esc(obj.id)}" aria-label="Позиция справочника">${choices}</select><button class="tiny-button" data-action="add-line" data-object="${esc(obj.id)}">Добавить</button></div>` : `<span class="hint">Сначала добавьте материал или работу в справочник</span><button class="tiny-button" data-view="catalog">Справочник</button>`;
    return `<div class="object-card"><div class="object-head"><div><div class="mini-label">ОБЪЕКТ / ${String(index+1).padStart(2,'0')}</div><h3>${esc(obj.title||'Рекламная конструкция')}</h3><p>Геометрия задаётся в миллиметрах. Нормы расхода применяются только к позициям с режимом «площадь» или «периметр».</p></div><div class="button-group">${btn('Удалить объект','delete-object','tiny-button danger',`data-object="${esc(obj.id)}"`)}</div></div>
      <div class="object-fields">${field('Наименование',obj.title,'title','object',obj.id)}${field('Ширина, мм',obj.widthMm,'widthMm','object',obj.id,'inputmode="decimal"')}${field('Высота, мм',obj.heightMm,'heightMm','object',obj.id,'inputmode="decimal"')}${field('Количество конструкций',obj.count,'count','object',obj.id,'type="number" min="1" step="1"')}</div>
      <div class="lines">${calc.lines.length?obj.lines.map((l,i)=>lineRow(order,obj,l,i,calc)).join(''):'<div class="empty" style="margin-top:16px"><strong>Нет позиций</strong>Добавьте материалы и операции для этого объекта.</div>'}</div>
      <div class="line-footer">${control}<span class="hint">Площадь одной конструкции: ${num(calc.area)} м²</span></div>
      <div class="object-totals"><span>Себестоимость <strong>${rub(calc.cost)}</strong></span><span>Продажа <strong>${rub(calc.sale)}</strong></span></div></div>`;
  }
  function orderView(o){
    const c=E.calcOrder(o,db.catalog);
    const printed = c.incomplete ? '<span class="badge-warn">Не заполнены некоторые цены или состав расчёта</span>' : '<span class="badge-ok">Все позиции имеют цены</span>';
    return `<div class="page-heading"><div><button class="back" data-action="back">← Вернуться к заказам</button><h1>Расчёт № ${o.number}</h1><p>Создан ${date(o.created)} · ${esc(o.customer.name||'Клиент не указан')}</p></div><div class="toolbar">${btn('Копировать','copy-order')}${btn('КП для печати','print','primary',c.incomplete?'disabled title="Сначала заполните состав и цены"':'')}</div></div>
      <section class="card"><div class="section-title"><h2>Карточка заказа</h2>${printed}</div><div class="form-grid three">
      ${field('Клиент',o.customer.name,'customer.name','order')}${field('Телефон',o.customer.phone,'customer.phone','order')}${field('Адрес монтажа',o.customer.address,'customer.address','order')}
      ${field('Менеджер',o.manager,'manager','order')}${selectField('Статус',o.status,statuses,'status','order')}${field('Доп. расходы / заказ, ₽',o.otherExpense,'otherExpense','order','','inputmode="decimal"')}
      ${field('Скидка, %',o.discountPct,'discountPct','order','','inputmode="decimal"')}${field('Примечание',o.note,'note','order')}
      </div></section>
      <section class="card"><div class="section-title"><h2>Состав заказа</h2>${btn('＋ Добавить объект','add-object','secondary')}</div>${o.objects.map((obj,i)=>objectCard(o,obj,i,c.objects[i])).join('')}</section>
      <section class="card"><div class="section-title"><h2>Экономика заказа</h2><small>Скидка применяется к продаже по всему заказу</small></div><div class="financial"><div><span class="mini-label">ПРОДАЖА ПО ПРАЙСУ</span><strong>${rub(c.sale)}</strong></div><div><span class="mini-label">СЕБЕСТОИМОСТЬ</span><strong>${rub(c.cost)}</strong></div><div><span class="mini-label">ИТОГО СО СКИДКОЙ</span><strong>${rub(c.total)}</strong></div><div><span class="mini-label">ПЛАНОВАЯ ПРИБЫЛЬ</span><strong>${rub(c.profit)}</strong></div></div>
      <div class="summary-note">Материалы: ${rub(c.material)} · Работы: ${rub(c.labor)} · Прочие расходы: ${rub(c.expense)} · Скидка: ${rub(c.discount)} · Маржинальность: ${num(c.margin)} %. Налоги, накладные коэффициенты и технологические нормы исходной таблицы пока не перенесены.</div></section>
      <div class="toolbar">${btn('Удалить заказ','delete-order','secondary danger')}</div>`;
  }
  function catalogView(){
    const row=(i)=>`<tr><td class="dim">${esc(i.id.slice(0,8))}</td><td><input class="name" data-scope="catalog" data-ref="${esc(i.id)}" data-edit="name" value="${esc(i.name)}"></td>
      <td><select data-scope="catalog" data-ref="${esc(i.id)}" data-edit="category"><option value="material" ${i.category==='material'?'selected':''}>Материал</option><option value="work" ${i.category==='work'?'selected':''}>Работа</option></select></td>
      <td><input data-scope="catalog" data-ref="${esc(i.id)}" data-edit="unit" value="${esc(i.unit)}"></td>
      <td><select data-scope="catalog" data-ref="${esc(i.id)}" data-edit="driver">${[['manual','Вручную'],['area','Площадь'],['perimeter','Периметр']].map(([v,n])=>`<option value="${v}" ${i.driver===v?'selected':''}>${n}</option>`).join('')}</select></td>
      <td><input data-scope="catalog" data-ref="${esc(i.id)}" data-edit="factor" type="number" min="0" step="any" value="${esc(i.factor??1)}"></td>
      <td><input data-scope="catalog" data-ref="${esc(i.id)}" data-edit="unitCost" type="number" min="0" step="any" value="${esc(i.unitCost)}"></td>
      <td><input data-scope="catalog" data-ref="${esc(i.id)}" data-edit="unitPrice" type="number" min="0" step="any" value="${esc(i.unitPrice)}"></td>
      <td><button class="plain-close" title="Удалить из справочника" data-action="delete-item" data-id="${esc(i.id)}">×</button></td></tr>`;
    return `${header('Справочник','Единые позиции материалов и производственных операций. Пока хранятся локально.')}
      <section class="card"><div class="section-title"><h2>Добавить позицию</h2></div><div class="form-grid three">
      <div class="field"><label>Наименование</label><input id="new-name" placeholder="Название материала или операции"></div>
      <div class="field"><label>Тип</label><select id="new-category"><option value="material">Материал</option><option value="work">Работа</option></select></div>
      <div class="field"><label>Единица</label><input id="new-unit" value="шт" placeholder="шт, м², м.п."></div>
      <div class="field"><label>Количество по</label><select id="new-driver"><option value="manual">Вручную</option><option value="area">Площади, м²</option><option value="perimeter">Периметру, м</option></select></div>
      <div class="field"><label>Себестоимость за единицу</label><input id="new-cost" type="number" min="0" step="any" value="0"></div>
      <div class="field"><label>Продажа за единицу</label><input id="new-price" type="number" min="0" step="any" value="0"></div>
      </div><div style="margin-top:18px">${btn('Добавить в справочник','add-item','primary')}</div></section>
      <section class="card"><div class="section-title"><h2>Позиции</h2><small>${db.catalog.length} поз.</small></div>${db.catalog.length?`<div class="table-wrap"><table class="catalog-table"><thead><tr><th>ID</th><th>Позиция</th><th>Тип</th><th>Ед.</th><th>Расход</th><th>Коэф.</th><th>Себест.</th><th>Продажа</th><th></th></tr></thead><tbody>${db.catalog.map(row).join('')}</tbody></table></div>`:'<div class="empty"><strong>Справочник пока пустой</strong>Цены из Excel не импортированы. Добавляйте только подтверждённые позиции.</div>'}
      <p class="hint">Для режима «Площадь» количество = ширина × высота в м² × коэффициент × количество объектов. Для «Периметра» — длина периметра в метрах × коэффициент. Сложные производственные нормы Excel будут подключены отдельно.</p></section>`;
  }
  function aboutView(){return `${header('О системе','Границы первой версии и правила безопасности.')}
      <section class="card about-text"><h2>А19 — отдельный продукт</h2><p>Этот прототип не использует код, хранилище, учётные данные и рабочие заказы «Идеал Потолок». Исходный продукт и его репозитории остаются без изменений.</p><p><strong>Уже работает:</strong> создание и копирование заказов, несколько рекламных объектов в заказе, ручное редактирование, независимый каталог, простые геометрические нормы расхода, подсчёт стоимости, скидок и прибыли, фиксация цен завершённых заказов, скачивание резервной копии и печать предварительного КП.</p><p><strong>Не готово к производству:</strong> подключение Excel-модели и внешних справочников, специализированные формулы, облачная база, права сотрудников, аудит изменений, полноценные производственные документы, проверка соответствия корпоративному сайту.</p><p>Сохранение данных сейчас работает только в браузере. Очистка браузерного хранилища приведёт к их потере, если не скачать резервную копию. Использовать для реальных заказов нельзя.</p><p>Бренд: предоставленный векторный логотип А19. Палитра и типографика сайта <a href="https://a19.group/" rel="noopener noreferrer" target="_blank">a19.group</a> требуют отдельной проверки доступного исходника сайта.</p></section>`;}
  function render(){
    document.querySelectorAll('[data-view].nav-item').forEach(n=>n.classList.toggle('active',n.dataset.view===view));
    const o=selected&&byId(selected);
    $('#app').innerHTML=view==='catalog'?catalogView():view==='about'?aboutView():o?orderView(o):ordersView();
  }
  function updateField(el){
    const scope=el.dataset.scope,path=el.dataset.edit;
    if(!scope||!path)return;
    let obj;
    if(scope==='order') obj=byId(selected);
    if(scope==='object') obj=objById(byId(selected),el.dataset.ref);
    if(scope==='catalog') obj=itemById(el.dataset.ref);
    if(scope==='line')obj=objById(byId(selected),el.dataset.object)?.lines[Number(el.dataset.line)];
    if(!obj)return;
    const keys=path.split('.');const field=keys.pop();let target=obj;
    for(const key of keys){target=target[key];if(!target)return;}
    let value=el.value;
    if(['qtyOverride','costOverride','priceOverride'].includes(field)) value=value.trim()===''?null:E.amount(value);
    if(['widthMm','heightMm','count','discountPct','otherExpense','factor','unitCost','unitPrice'].includes(field))value=Math.max(0,E.amount(value));
    if(field==='count')value=Math.max(1,Math.floor(value));
    if(field==='discountPct')value=Math.min(100,value);
    if(field==='status'&&value==='done'&&obj.status!=='done') E.freezeOrder(obj,db.catalog);
    target[field]=value;
    if(scope==='order')obj.updated=new Date().toISOString();
    save();
  }
  function download(name, content) {
    const blob=new Blob([content],{type:'application/json;charset=utf-8'});
    const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);
  }
  function printEstimate(o){
    const c=E.calcOrder(o,db.catalog);
    if(c.incomplete){alert('Для КП необходимо заполнить состав заказа, себестоимость и цену каждой позиции.');return;}
    $('#print-sheet').innerHTML=`<img src="assets/logo.svg" alt="А19"><h1>Коммерческое предложение № ${o.number}</h1><p>Клиент: ${esc(o.customer.name||'—')}<br>Дата: ${date(new Date().toISOString())}</p>${c.objects.map((obj,i)=>`<h2>${i+1}. ${esc(obj.title)}</h2><p>Габариты: ${esc(o.objects[i].widthMm)} × ${esc(o.objects[i].heightMm)} мм · Количество: ${obj.count}</p><table><thead><tr><th>Позиция</th><th>Кол-во</th><th>Цена</th><th class="right">Стоимость</th></tr></thead><tbody>${obj.lines.map(l=>`<tr><td>${esc(l.name)}</td><td>${num(l.total)} ${esc(l.unit)}</td><td>${rub(l.unitPrice)}</td><td class="right">${rub(l.sale)}</td></tr>`).join('')}</tbody></table>`).join('')}<div class="print-footer">ИТОГО: ${rub(c.total)}${c.discount?' (скидка '+rub(c.discount)+')':''}</div>`;
    window.print();
  }
  document.addEventListener('change',e=>{
    if(e.target.id==='import-input'){
      const file=e.target.files?.[0];if(!file)return;
      file.text().then(raw=>{
        let next;try{next=JSON.parse(raw);}catch(e){alert('Файл JSON повреждён');return;}
        if(!next||next.schema!==1||!Array.isArray(next.orders)||!Array.isArray(next.catalog)){alert('Неверный формат копии А19');return;}
        if(!confirm('Заменить текущие локальные данные загруженной копией? Сохраните резервную копию перед импортом.'))return;
        db={schema:1,orders:next.orders,catalog:next.catalog};selected=null;view='orders';save();
      }).catch(()=>alert('Не удалось прочитать файл')).finally(()=>{e.target.value='';});
      return;
    }
    if(e.target.matches('[data-edit]'))updateField(e.target);
  });
  document.addEventListener('click',e=>{
    const v=e.target.closest('[data-view]');if(v){view=v.dataset.view;selected=null;render();return;}
    const b=e.target.closest('[data-action]');if(!b)return;
    const act=b.dataset.action;
    const o=byId(selected);
    if(act==='new-order'){const newO=makeOrder();db.orders.push(newO);selected=newO.id;view='orders';save();}
    else if(act==='open-order'){selected=b.dataset.id;view='orders';render();}
    else if(act==='back'){selected=null;render();}
    else if(act==='add-object'&&o){o.objects.push(makeObject(o.objects.length+1));save();}
    else if(act==='delete-object'&&o){if(o.objects.length===1){alert('В заказе должен быть хотя бы один объект.');return;}if(confirm('Удалить объект и все его позиции?')){o.objects=o.objects.filter(x=>x.id!==b.dataset.object);save();}}
    else if(act==='add-line'&&o){const itemId=document.getElementById('pick-'+b.dataset.object)?.value;if(!itemById(itemId))return;const obj=objById(o,b.dataset.object);obj.lines.push({itemId,qty:1,qtyOverride:null,costOverride:null,priceOverride:null});save();}
    else if(act==='delete-line'&&o){const obj=objById(o,b.dataset.object);obj.lines.splice(Number(b.dataset.line),1);save();}
    else if(act==='copy-order'&&o){const next=clone(o);next.id=uid();next.number=Math.max(0,...db.orders.map(x=>Number(x.number)||0))+1;next.status='draft';next.created=new Date().toISOString();next.updated=next.created;next.objects.forEach(x=>{x.id=uid();x.lines.forEach(y=>delete y.fix);});db.orders.push(next);selected=next.id;save();}
    else if(act==='delete-order'&&o){if(confirm(`Удалить расчёт № ${o.number}?`)){db.orders=db.orders.filter(x=>x.id!==o.id);selected=null;save();}}
    else if(act==='print'&&o)printEstimate(o);
    else if(act==='add-item'){
      const name=$('#new-name').value.trim();if(name.length<2){alert('Укажите название позиции');return;}
      db.catalog.push({id:uid(),name,category:$('#new-category').value,unit:$('#new-unit').value.trim()||'шт',driver:$('#new-driver').value,factor:1,unitCost:E.amount($('#new-cost').value),unitPrice:E.amount($('#new-price').value)});save();
    }
    else if(act==='delete-item'){
      const item=itemById(b.dataset.id);if(!item)return;
      if(db.orders.some(o=>o.objects.some(x=>x.lines.some(y=>y.itemId===item.id)))){alert('Эта позиция используется в заказах. Чтобы сохранить историю, удаление заблокировано.');return;}
      if(confirm('Удалить позицию из справочника?')){db.catalog=db.catalog.filter(x=>x.id!==item.id);save();}
    }
    else if(act==='export'){download('a19-backup-'+new Date().toISOString().slice(0,10)+'.json',JSON.stringify(db,null,2));}
    else if(act==='import')$('#import-input').click();
  });
  render();
})();
