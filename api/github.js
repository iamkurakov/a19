/* A19 shared data gateway. Deploy as a Vercel Function.
   GitHub credential stays in server environment A19_GITHUB_TOKEN.
   The web interface on GitHub Pages uses this route; users never enter tokens.
   A19_SHARED_KEY is optional: if specified, requests require a matching
   X-A19-Share-Key header (e.g. from a capability link). Without it, anyone
   who can access the public A19 site may read and edit A19 data.
*/
'use strict';
const REPO='iamkurakov/a19_data';
const GITHUB='https://api.github.com/repos/'+REPO;
const ALLOWED_ORIGINS=new Set(['https://iamkurakov.github.io']);
const GH_HEADERS={Accept:'application/vnd.github+json','X-GitHub-Api-Version':'2022-11-28'};

function classify(method,raw){
  if(typeof raw!=='string'||raw.length>200||raw.includes('..')||raw.includes('#'))return null;
  const [path,query,...rest]=raw.split('?');
  if(rest.length)return null;
  if(method==='GET'&&path===''&&!query)return {path:'',kind:'repo'};
  if(method==='GET'&&path==='/contents/orders'&&(!query||query==='per_page=1000'))return {path:'/contents/orders',kind:'list'};
  const match=path.match(/^\/contents\/(orders\/([\w-]+)\.json|settings\/external_confirmations\.json|models\/A19_excel_model_PRIVATE\.json\.gz)$/);
  if(!match||query)return null;
  const kind=match[1].startsWith('orders/')?'order':match[1].startsWith('settings/')?'settings':'model';
  if(method==='GET'||((method==='PUT'||method==='DELETE')&&kind==='order')||(method==='PUT'&&kind==='settings'))return {path,kind};
  return null;
}
function createHandler({env=process.env,fetchFn=global.fetch}={}){
  return async function handler(req,res){
    const origin=req.headers.origin||'';
    const configuredOrigins=String(env.A19_ALLOWED_ORIGINS||'').split(',').map(x=>x.trim()).filter(Boolean);
    const allowed=new Set([...ALLOWED_ORIGINS,...configuredOrigins]);
    res.setHeader('Cache-Control','no-store, max-age=0');
    res.setHeader('Vary','Origin');
    if(origin&&allowed.has(origin))res.setHeader('Access-Control-Allow-Origin',origin);
    res.setHeader('Access-Control-Allow-Methods','GET,PUT,DELETE,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers','Content-Type,X-A19-Share-Key');
    if(req.method==='OPTIONS')return res.status(origin&&!allowed.has(origin)?403:204).end();
    if(origin&&!allowed.has(origin))return res.status(403).json({error:'Недопустимый источник запроса'});
    if(env.A19_SHARED_KEY&&req.headers['x-a19-share-key']!==env.A19_SHARED_KEY)return res.status(403).json({error:'Нужна полная рабочая ссылка А19'});
    if(!env.A19_GITHUB_TOKEN)return res.status(503).json({error:'Общее хранилище ещё не подключено на сервере'});
    const raw=typeof req.query.path==='string'?req.query.path:'';
    const route=classify(req.method,raw);
    if(!route)return res.status(400).json({error:'Недопустимая операция с данными А19'});
    try{
      let body;
      if(req.method==='PUT'||req.method==='DELETE'){
        body=typeof req.body==='string'?JSON.parse(req.body):req.body;
        if(!body||typeof body!=='object'||Array.isArray(body))return res.status(400).json({error:'Неверный формат запроса'});
        if(req.method==='PUT'){
          if(typeof body.content!=='string'||body.content.length>750000||!/^[a-zA-Z0-9+/=\s]+$/.test(body.content))return res.status(413).json({error:'Некорректный размер данных'});
          const bytes=Buffer.from(body.content.replace(/\s/g,''),'base64');
          if(bytes.length>540000||bytes.length===0)return res.status(413).json({error:'Слишком большой заказ'});
          const data=JSON.parse(bytes.toString('utf8'));
          if(route.kind==='order'&&(!data.id||!/^[-\w]+$/.test(data.id)||route.path!==`/contents/orders/${data.id}.json`))return res.status(400).json({error:'Неверный идентификатор заказа'});
          if(route.kind==='settings'&&data.format!=='a19-external-v1')return res.status(400).json({error:'Неверный формат справочника'});
        }
        if(typeof body.message!=='string'||body.message.length>200||body.sha&&(!/^[a-f0-9]{40}$/.test(body.sha)))return res.status(400).json({error:'Недопустимые параметры записи'});
        body=JSON.stringify(body);
      }
      const response=await fetchFn(GITHUB+route.path,{
        method:req.method,
        headers:{...GH_HEADERS,'Authorization':'Bearer '+env.A19_GITHUB_TOKEN,'Content-Type':'application/json'},
        ...(body?{body}:{})
      });
      if(response.status===404)return res.status(404).json({message:'Not Found'});
      if(!response.ok){const code=response.status;return res.status(code).json({message:code===409||code===422?'Конфликт сохранения заказа':code===401||code===403?'Ошибка серверного подключения GitHub':'Ошибка хранилища GitHub'});}
      const responseBody=await response.json();
      if(route.kind==='repo')return res.status(200).json({full_name:responseBody.full_name,private:responseBody.private});
      if(route.kind==='list')return res.status(200).json(Array.isArray(responseBody)?responseBody.filter(x=>x.type==='file'&&/^[-\w]+\.json$/.test(x.name)).map(x=>({name:x.name,type:'file',sha:x.sha})):[]);
      if(req.method==='GET')return res.status(200).json({type:responseBody.type,sha:responseBody.sha,content:responseBody.content,encoding:responseBody.encoding});
      if(req.method==='PUT')return res.status(200).json({content:{sha:responseBody.content?.sha}});
      return res.status(200).json({ok:true});
    }catch(e){return res.status(502).json({message:'Не удалось получить ответ хранилища А19'});}
  };
}
const handler=createHandler();
module.exports=handler;
module.exports.createHandler=createHandler;
module.exports.classify=classify;
