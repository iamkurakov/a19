const test=require('node:test');
const assert=require('node:assert/strict');
const {createClient,CloudError}=require('../src/cloud.js');
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


test('shared client proxies all operations with no Github credential in browser',async()=>{
 const calls=[];
 const fetchFn=async (url,opts)=>{calls.push({url,opts});return {ok:true,status:200,json:async()=>({full_name:'iamkurakov/a19_data',private:true})};};
 const c=createClient('',fetchFn,'iamkurakov/a19_data','https://a19-example.vercel.app','share-link-secret');
 assert.equal(await c.validate(),true);
 assert.equal(calls.length,1);
 assert.equal(calls[0].url,'https://a19-example.vercel.app/api/github?path=');
 assert.equal(calls[0].opts.headers.Authorization,undefined);
 assert.equal(calls[0].opts.headers['X-A19-Share-Key'],'share-link-secret');
});
