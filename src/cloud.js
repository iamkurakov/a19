/* A19 single-manager private GitHub storage.
   No secret is built into this public file. A limited-scope token is supplied
   by the manager on a trusted device. The public site itself is NOT an auth boundary. */
(function(root,factory){
  const api=factory();
  if(typeof module==='object'&&module.exports)module.exports=api;
  if(root)root.A19Cloud=api.browser(root);
})(typeof globalThis!=='undefined'?globalThis:this,function(){
'use strict';
const DEFAULT_REPO='iamkurakov/a19_data';
class CloudError extends Error {constructor(message,status=0){super(message);this.name='CloudError';this.status=status;}}
function enc64(str){const bytes=new TextEncoder().encode(str);let s='';for(let i=0;i<bytes.length;i+=16384)s+=String.fromCharCode(...bytes.subarray(i,i+16384));return btoa(s);}
function dec64(data){const s=atob(String(data||'').replace(/\s/g,''));const bytes=new Uint8Array(s.length);for(let i=0;i<s.length;i++)bytes[i]=s.charCodeAt(i);return bytes;}
function createClient(token,fetchFn=fetch,repo=DEFAULT_REPO){
  if(!/^[a-zA-Z0-9-]+\/[a-zA-Z0-9_.-]+$/.test(repo))throw new CloudError('Неверный репозиторий');
  const base='https://api.github.com/repos/'+repo;
  async function request(method,path,payload){
    const res=await fetchFn(base+path,{method,headers:{'Accept':'application/vnd.github+json','X-GitHub-Api-Version':'2022-11-28','Authorization':'Bearer '+token,'Content-Type':'application/json'},...(payload===undefined?{}:{body:JSON.stringify(payload)})});
    if(res.status===404){if(method==='GET')return null;throw new CloudError('Репозиторий или файл не найден. Проверьте доступ токена.',404);}
    if(!res.ok){
      let reason='GitHub вернул '+res.status;
      if(res.status===401)reason='Токен не принят GitHub';
      if(res.status===403)reason='Недостаточно прав токена либо достигнут лимит GitHub';
      if(res.status===409||res.status===422)reason='Конфликт сохранения: файл изменён в другом сеансе';
      throw new CloudError(reason,res.status);
    }
    return res.status===204?{}:res.json();
  }
  async function validate(){const repoMeta=await request('GET','');if(!repoMeta||repoMeta.full_name!==repo||repoMeta.private!==true)throw new CloudError('Нужен приватный репозиторий '+repo);return true;}
  const allowedPath=path=>/^(orders\/[a-zA-Z0-9_-]+\.json|models\/[a-zA-Z0-9_-]+\.json\.gz|settings\/[a-zA-Z0-9_-]+\.json)$/.test(path);
  async function getFile(path){if(!allowedPath(path))throw new CloudError('Недопустимый путь');const data=await request('GET','/contents/'+path);if(!data)return null;if(data.type!=='file'||!data.content)throw new CloudError('Файл слишком большой или недоступен в API');return {sha:data.sha,bytes:dec64(data.content)};}
  async function readJSON(path){const f=await getFile(path);return f?{sha:f.sha,data:JSON.parse(new TextDecoder().decode(f.bytes))}:null;}
  async function putJSON(path,data,sha){if(!allowedPath(path))throw new CloudError('Недопустимый путь');const encoded=enc64(JSON.stringify(data));const body={message:'A19: save '+path,content:encoded,...(sha?{sha}:{})};const result=await request('PUT','/contents/'+path,body);return result?.content?.sha||null;}
  async function deleteFile(path,sha){if(!allowedPath(path))throw new CloudError('Недопустимый путь');if(!sha)throw new CloudError('Нельзя удалить без подтверждённой версии');await request('DELETE','/contents/'+path,{message:'A19: delete '+path,sha});}
  async function listOrders(){const list=await request('GET','/contents/orders?per_page=1000');return Array.isArray(list)?list.filter(x=>x.type==='file'&&/^[-\w.]+\.json$/.test(x.name)):[];}
  async function workbook(){const f=await getFile('models/A19_excel_model_PRIVATE.json.gz');if(!f)throw new CloudError('Нет модели Excel в приватном репозитории');if(typeof DecompressionStream!=='function')throw new CloudError('Этот браузер не поддерживает распаковку модели. Используйте актуальный Chrome или Edge.');const stream=new Blob([f.bytes]).stream().pipeThrough(new DecompressionStream('gzip'));const content=await new Response(stream).text();const data=JSON.parse(content);if(data.format!=='a19-excel-model-v1'||!data.sheets?.['Объём1']||!data.tables)throw new CloudError('Некорректная модель Excel в облаке');return data;}
  return {validate,request,getFile,readJSON,putJSON,deleteFile,listOrders,workbook};
}
function browser(root){
  if(!root.document)return {createClient,CloudError};
  const LOCAL_KEY='a19.private.token.v1',SESSION_KEY='a19.private.session.v1';
  const QUEUE_KEY='a19.private.pending.v1',DELETES_KEY='a19.private.deletions.v1';
  const SHA_KEY='a19.private.shas.v1';
  const $=s=>root.document.querySelector(s);
  const parse=(key,fallback)=>{try{return JSON.parse(root.localStorage.getItem(key))??fallback;}catch(_){return fallback;}};
  const saveMap=(key,value)=>root.localStorage.setItem(key,JSON.stringify(value));
  const pending=()=>parse(QUEUE_KEY,{}),deletions=()=>parse(DELETES_KEY,{}),shas=()=>parse(SHA_KEY,{});
  let token=root.sessionStorage.getItem(SESSION_KEY)||root.localStorage.getItem(LOCAL_KEY)||'';
  let client=token?createClient(token,root.fetch.bind(root)):null;
  let syncing=false,debounce=null,connecting=false,lastSynced='',lastError='',connected=false,modelLoaded=false;
  const state=()=>({connected,syncing,connecting,pending:Object.keys(pending()).length+Object.keys(deletions()).length,lastSynced,error:lastError,hasToken:!!token,modelLoaded});
  function report(){const st=state();const label=st.connecting?'Подключение…':st.syncing?'Сохраняется…':st.connected?(st.pending?'В очереди: '+st.pending:'Облако подключено'):(st.error?'Ошибка облака':'Облако не подключено');
    for(const el of root.document.querySelectorAll('[data-cloud-status]'))el.textContent=label;
    const button=$('#cloud-connect');if(button)button.textContent=st.connected?'Настроить облако':'Подключить облако';
    const stamp=$('#a19-cloud-caption');if(stamp)stamp.textContent=st.connected?('А19 · приватное хранилище'+(st.lastSynced?' · сохранено '+st.lastSynced:'')):'Сейчас данные находятся только в браузере';
    root.dispatchEvent(new CustomEvent('a19-cloud-status',{detail:st}));
  }
  function schedule(){if(!token)return;clearTimeout(debounce);debounce=setTimeout(()=>{sync().catch(()=>{});},1100);report();}
  function markOrder(id){if(!id||typeof id!=='string')return;const data=pending();data[id]=Date.now();saveMap(QUEUE_KEY,data);schedule();}
  function markDeleted(id){if(!id)return;const p=pending();delete p[id];saveMap(QUEUE_KEY,p);const d=deletions();d[id]=Date.now();saveMap(DELETES_KEY,d);schedule();}
  function markExternal(){root.localStorage.setItem('a19.private.external-dirty','1');schedule();}
  async function getWorkbook(){if(!client)throw new CloudError('Сначала подключите закрытый репозиторий');await root.A19ExcelUI?.ready;const has=root.A19ExcelUI?.hasModel?.();if(!has){const wb=await client.workbook();await root.A19ExcelUI.ensurePrivateModel(wb);modelLoaded=true;}else modelLoaded=true;}
  async function syncExternal(){
    const path='settings/external_confirmations.json';const remote=await client.readJSON(path);
    if(remote?.data?.format==='a19-external-v1')await root.A19ExcelUI?.applyConfirmedExternal(remote.data.values||[]);
    if(root.localStorage.getItem('a19.private.external-dirty')==='1'){
      const values=root.A19ExcelUI?.confirmedExternal?.()||[];
      const sha=await client.putJSON(path,{format:'a19-external-v1',updatedAt:new Date().toISOString(),values},remote?.sha);
      root.localStorage.removeItem('a19.private.external-dirty');
      return sha;
    }
  }
  async function syncOrders(){
    const bridge=root.A19OrderBridge;if(!bridge)return;
    const files=await client.listOrders();const local=bridge.getAll();const localBy=new Map(local.map(o=>[o.id,o]));const stored=shas();const dirty=pending(),del=deletions();const remote=[];let conflicts=[];
    for(const file of files){
      const id=file.name.slice(0,-5),path='orders/'+file.name;
      if(del[id])continue;
      const f=await client.readJSON(path);if(!f?.data?.id||f.data.id!==id)continue;
      // Dirty local orders never overwrite an unseen remote revision.
      if(dirty[id]&&localBy.has(id)&&((stored[id]&&stored[id]!==f.sha)||(!stored[id]&&JSON.stringify(localBy.get(id))!==JSON.stringify(f.data)))){conflicts.push('Заказ '+id+' уже отличается в облаке');continue;}
      stored[id]=f.sha;remote.push(f.data);
    }
    bridge.mergeRemote(remote,Object.keys(dirty));saveMap(SHA_KEY,stored);
    for(const o of bridge.getAll()){if(!stored[o.id]&&!dirty[o.id]&&!del[o.id])dirty[o.id]=Date.now();}
    saveMap(QUEUE_KEY,dirty);
    if(conflicts.length)throw new CloudError('Конфликт локальных заказов: '+conflicts.length+'. Скачайте резервную копию и проверьте версии.');
    await flush();
  }
  async function flush(){
    const bridge=root.A19OrderBridge;if(!bridge)return;
    const dirty=pending(),del=deletions(),stored=shas();
    for(const id of Object.keys(del)){
      try{if(stored[id])await client.deleteFile('orders/'+id+'.json',stored[id]);delete stored[id];delete del[id];saveMap(SHA_KEY,stored);saveMap(DELETES_KEY,del);}
      catch(e){if(e.status===404){delete del[id];saveMap(DELETES_KEY,del);}else throw e;}
    }
    for(const id of Object.keys(dirty)){
      const order=bridge.getById(id);if(!order){delete dirty[id];saveMap(QUEUE_KEY,dirty);continue;}
      const stamp=dirty[id];
      const path='orders/'+id+'.json';
      const sha=await client.putJSON(path,order,stored[id]);
      stored[id]=sha;saveMap(SHA_KEY,stored);
      const current=pending();if(current[id]===stamp){delete current[id];saveMap(QUEUE_KEY,current);}
    }
  }
  async function sync(){
    if(!client||syncing||connecting)return false;
    syncing=true;lastError='';report();
    try{await client.validate();await getWorkbook();await syncExternal();await syncOrders();connected=true;lastSynced=new Date().toLocaleTimeString('ru-RU',{hour:'2-digit',minute:'2-digit'});return true;}
    catch(e){lastError=e.message||String(e);connected=false;return false;}
    finally{syncing=false;report();if(connected&&(Object.keys(pending()).length||Object.keys(deletions()).length||root.localStorage.getItem('a19.private.external-dirty')==='1'))schedule();}
  }
  async function connect(value,remember){
    if(!value||!/^github_pat_|^ghp_|^ghu_/.test(value))throw new CloudError('Нужен GitHub fine-grained token с доступом только к a19_data');
    connecting=true;report();
    const next=createClient(value,root.fetch.bind(root));
    try{await next.validate();token=value;client=next;if(remember){root.localStorage.setItem(LOCAL_KEY,value);root.sessionStorage.removeItem(SESSION_KEY);}else{root.sessionStorage.setItem(SESSION_KEY,value);root.localStorage.removeItem(LOCAL_KEY);}lastError='';connecting=false;report();const ok=await sync();if(!ok)throw new CloudError(lastError);return true;}
    finally{connecting=false;report();}
  }
  function disconnect(){token='';client=null;connected=false;lastError='';root.localStorage.removeItem(LOCAL_KEY);root.sessionStorage.removeItem(SESSION_KEY);report();}
  function openDialog(){const d=$('#cloud-dialog');if(!d)return;const s=state();const status=$('#cloud-dialog-status');if(status)status.textContent=s.error|| (s.connected?'Соединение установлено. Заказы автоматически записываются в приватный репозиторий.':'Подключите GitHub только к закрытому a19_data.');d.showModal();}
  root.document.addEventListener('click',async e=>{
    const target=e.target.closest('[data-cloud-action]');if(!target)return;
    const action=target.dataset.cloudAction;
    if(action==='open'){openDialog();return;}
    if(action==='close'){$('#cloud-dialog')?.close();return;}
    if(action==='disconnect'){disconnect();$('#cloud-dialog')?.close();return;}
    if(action==='sync'){await sync();$('#cloud-dialog')?.close();return;}
  });
  root.document.addEventListener('submit',async e=>{
    if(e.target.id!=='cloud-form')return;e.preventDefault();const input=$('#cloud-token'),remember=$('#cloud-remember');const btn=$('#cloud-submit');if(btn)btn.disabled=true;
    try{await connect(input.value.trim(),!!remember?.checked);input.value='';$('#cloud-dialog')?.close();root.A19OrderBridge?.rerender?.();}
    catch(err){const line=$('#cloud-dialog-status');if(line)line.textContent=err.message;}
    finally{if(btn)btn.disabled=false;}
  });
  root.addEventListener('online',()=>{if(token)sync();});
  const api={state,connect,disconnect,markOrder,markDeleted,markExternal,sync,openDialog,createClient};
  setTimeout(()=>{if(token)sync();else report();},0);
  return api;
}
return {createClient,CloudError,browser};
});
