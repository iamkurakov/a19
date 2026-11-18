const test=require('node:test');
const assert=require('node:assert/strict');
const {createClient,CloudError,parseShareToken,makeShareUrl}=require('../src/cloud.js');
const base='https://api.github.com/repos/iamkurakov/a19_data';
function makeFetch(routes,log){return async (url,opts)=>{log.push({url,method:opts.method,headers:opts.headers,body:opts.body||''});const code=routes[url+'|'+opts.method];if(!code)return {ok:false,status:404,json:async()=>({message:'Not Found'})};return {ok:code.status>=200&&code.status<300,status:code.status,json:async()=>code.json};};}
test('private repository is validated and token sent only as request header',async()=>{
 const calls=[];const mock=makeFetch({[base+'|GET']:{status:200,json:{full_name:'iamkurakov/a19_data',private:true}}},calls);
 const c=createClient('github_pat_example',mock);assert.equal(await c.validate(),true);
 assert.equal(calls.length,1);assert.equal(calls[0].headers.Authorization,'Bearer github_pat_example');assert.equal(calls[0].url.includes('github_pat_'),false);
});
test('reject public repository even with working GitHub token',async()=>{
 const mock=makeFetch({[base+'|GET']:{status:200,json:{full_name:'iamkurakov/a19_data',private:false}}},[]);
 await assert.rejects(createClient('github_pat_example',mock).validate(),/приватный репозиторий/);
});
test('orders are written only to the private data repository and encoded as UTF-8',async()=>{
 const path='orders/order-123.json';const log=[];
 const mock=makeFetch({[base+'/contents/'+path+'|PUT']:{status:201,json:{content:{sha:'newsha'}}}},log);
 const order={id:'order-123',customer:{name:'Тестовая запись'},calculator:'excel',excelOverrides:{'Объём1!J10':1200}};
 assert.equal(await createClient('github_pat_test',mock).putJSON(path,order), 'newsha');
 const sent=JSON.parse(log[0].body);assert.deepEqual(JSON.parse(Buffer.from(sent.content,'base64').toString('utf8')),order);
 assert.ok(!log[0].url.includes('ideal_potolok'));assert.equal(log[0].headers.Authorization,'Bearer github_pat_test');
});
test('write conflicts remain visible instead of overwriting other session',async()=>{
 const log=[];const mock=makeFetch({[base+'/contents/orders/x.json|PUT']:{status:409,json:{message:'conflict'}}},log);
 await assert.rejects(createClient('github_pat_test',mock).putJSON('orders/x.json',{id:'x'},'outdatedsha'),e=>e instanceof CloudError&&e.status===409);
 assert.equal(JSON.parse(log[0].body).sha,'outdatedsha');
});
test('write paths are restricted to A19 data, not GitHub settings or other projects',async()=>{
 await assert.rejects(createClient('github_pat_test',async()=>{}).putJSON('../ideal_potolok/settings.json',{}),/Недопустимый путь/);
});


test('owner generates a working single link using the already connected token',()=>{
 const token='github_pat_existingA19scoped123';
 const url=makeShareUrl(token);
 assert.equal(url,'https://iamkurakov.github.io/a19/#a19=github_pat_existingA19scoped123');
 assert.equal(parseShareToken(new URL(url).hash),token);
 assert.equal(parseShareToken('#a19=garbage'),'');
});
test('link token never changes the private GitHub API origin',async()=>{
 const calls=[];const token=parseShareToken('#a19=github_pat_AnotherExistingToken');
 const mock=makeFetch({[base+'|GET']:{status:200,json:{full_name:'iamkurakov/a19_data',private:true}}},calls);
 assert.equal(await createClient(token,mock).validate(),true);
 assert.equal(calls[0].url,base);
 assert.equal(calls[0].headers.Authorization,'Bearer '+token);
});

test('recipient link connects to private GitHub automatically without typing a token',async()=>{
 const {browser}=require('../src/cloud.js');
 const saved=new Map(); const session=new Map(); const network=[];
 const storage=(map)=>({getItem:k=>map.get(k)??null,setItem:(k,v)=>map.set(k,String(v)),removeItem:k=>map.delete(k)});
 const token='github_pat_A19SharedTokenDemo';
 const doc={hidden:false,querySelector:()=>null,querySelectorAll:()=>[],addEventListener:()=>{}};
 const root={
  document:doc,location:{hash:'#a19='+token},
  localStorage:storage(saved),sessionStorage:storage(session),
  fetch:async(url,opts)=>{
    network.push({url,opts});
    if(url===base)return {status:200,ok:true,json:async()=>({full_name:'iamkurakov/a19_data',private:true})};
    if(url.endsWith('/contents/settings/external_confirmations.json'))return {status:404,ok:false,json:async()=>({message:'not found'})};
    if(url.endsWith('/contents/orders?per_page=1000'))return {status:200,ok:true,json:async()=>([])};
    throw new Error('Unexpected request '+url);
  },
  A19ExcelUI:{ready:Promise.resolve(),hasModel:()=>true},
  A19OrderBridge:{getAll:()=>[],mergeRemote:()=>{}},
  dispatchEvent:()=>{},addEventListener:()=>{},
  setTimeout:(cb)=>{Promise.resolve().then(cb);},
 };
 const c=browser(root);
 for(let i=0;i<10 && !c.state().connected;i++)await new Promise(resolve=>setTimeout(resolve,0));
 assert.equal(c.state().connected,true);
 assert.equal(c.state().shared,true);
 assert.equal(saved.get('a19.private.token.v1'),token);
 assert.equal(saved.get('a19.private.link-mode.v1'),'1');
 assert.equal(network.some(c=>c.url===base),true);
 assert.equal(network.every(c=>c.opts.headers.Authorization==='Bearer '+token),true);
});
