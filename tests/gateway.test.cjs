const test=require('node:test');
const assert=require('node:assert/strict');
const {createHandler,classify}=require('../api/github');
function response(){return {code:200,headers:{},jsonData:null,setHeader(k,v){this.headers[k]=v;return this;},status(n){this.code=n;return this;},json(x){this.jsonData=x;return this;},end(){return this;}};}
const env={A19_GITHUB_TOKEN:'test-secret'};
const origin='https://iamkurakov.github.io';
const repo='https://api.github.com/repos/iamkurakov/a19_data';
test('gateway permits A19 objects but cannot touch other repositories',()=>{
 assert.deepEqual(classify('GET',''),{path:'',kind:'repo'});
 assert.equal(classify('PUT','/contents/models/A19_excel_model_PRIVATE.json.gz'),null);
 assert.equal(classify('PUT','/contents/../../ideal_potolok/settings.json'),null);
 assert.equal(classify('GET','/contents/orders?id=x'),null);
 assert.equal(classify('PUT','/contents/orders/abc.json').kind,'order');
});
test('server token never appears in client responses or browser headers',async()=>{
 const calls=[];const mock=async (url,params)=>{calls.push({url,params});return {ok:true,status:200,json:async()=>({full_name:'iamkurakov/a19_data',private:true})};};
 const res=response();await createHandler({env,fetchFn:mock})({method:'GET',headers:{origin},query:{path:''}},res);
 assert.equal(res.code,200);assert.deepEqual(res.jsonData,{full_name:'iamkurakov/a19_data',private:true});
 assert.equal(calls[0].params.headers.Authorization,'Bearer test-secret');
 assert.equal(JSON.stringify(res.jsonData).includes('test-secret'),false);
 assert.equal(res.headers['Access-Control-Allow-Origin'],origin);
});
test('any browser on A19 site can read shared model without GitHub token',async()=>{
 let received='';const mock=async (url)=>{received=url;return {ok:true,status:200,json:async()=>({type:'file',sha:'abc',content:'YWJj',encoding:'base64'})};};
 const res=response();await createHandler({env,fetchFn:mock})({method:'GET',headers:{origin},query:{path:'/contents/models/A19_excel_model_PRIVATE.json.gz'}},res);
 assert.equal(res.code,200);assert.equal(res.jsonData.content,'YWJj');assert.equal(received,repo+'/contents/models/A19_excel_model_PRIVATE.json.gz');
});
test('shared order writes are restricted to matching order IDs',async()=>{
 const mock=async ()=>({ok:true,status:201,json:async()=>({content:{sha:'new-sha'}})});
 const body={message:'A19 order',content:Buffer.from(JSON.stringify({id:'order-01'})).toString('base64')};
 const r=response();await createHandler({env,fetchFn:mock})({method:'PUT',headers:{origin},query:{path:'/contents/orders/order-01.json'},body},r);
 assert.equal(r.code,200);assert.equal(r.jsonData.content.sha,'new-sha');
 const bad=response();await createHandler({env,fetchFn:mock})({method:'PUT',headers:{origin},query:{path:'/contents/orders/different.json'},body},bad);
 assert.equal(bad.code,400);
});
test('unconfigured server and denied external origin show explicit errors',async()=>{
 const dummy=async()=>{throw Error('should not call')};
 const noToken=response();await createHandler({env:{},fetchFn:dummy})({method:'GET',headers:{origin},query:{path:''}},noToken);assert.equal(noToken.code,503);
 const other=response();await createHandler({env,fetchFn:dummy})({method:'GET',headers:{origin:'https://untrusted.test'},query:{path:''}},other);assert.equal(other.code,403);
});
test('optional capability key permits one shareable URL without user login',async()=>{
 const dummy=async()=>({ok:true,status:200,json:async()=>({full_name:'iamkurakov/a19_data',private:true})});
 const res=response();await createHandler({env:{...env,A19_SHARED_KEY:'abc'},fetchFn:dummy})({method:'GET',headers:{origin},query:{path:''}},res);
 assert.equal(res.code,403);
 const granted=response();await createHandler({env:{...env,A19_SHARED_KEY:'abc'},fetchFn:dummy})({method:'GET',headers:{origin,'x-a19-share-key':'abc'},query:{path:''}},granted);
 assert.equal(granted.code,200);
});
