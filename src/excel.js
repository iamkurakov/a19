/* A19 Excel workbook interpreter (no dynamic code execution / eval).
   Loads a PRIVATE, locally imported JSON Office Open XML extraction.
   Formula fidelity is reported; unsupported external dependencies are errors,
   never substituted with invented prices. */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.A19Excel = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  class ExcelError extends Error {
    constructor(code, details) { super(code); this.name = 'ExcelError'; this.code = code; this.details = details || ''; }
  }
  const error = (e, details) => new ExcelError(e, details);
  const number = (x) => {
    if (x instanceof ExcelError) throw x;
    if (Array.isArray(x)) return number(x[0]);
    if (x === null || x === undefined || x === '' || x === false) return 0;
    if (x === true) return 1;
    if (typeof x === 'number') return x;
    const v=Number(String(x).replace(/\s/g,'').replace(',','.'));
    if (!Number.isFinite(v)) throw error('#VALUE!', 'Not numeric: '+String(x));
    return v;
  };
  const text = (x) => x === null || x === undefined ? '' : Array.isArray(x) ? text(x[0]) : String(x);
  const comparable = x => typeof x === 'number' ? x : typeof x === 'boolean' ? Number(x) : x === null ? 0 : String(x).toLocaleLowerCase();
  const truth = x => { if (x instanceof ExcelError) throw x; return Array.isArray(x) ? truth(x[0]) : !(x===null||x===''||x===false||x===0); };
  function tokenizer(s) {
    const out=[];let p=0;
    const ops='()[]!,;:+-*/^&=<>';
    while(p<s.length){
      if (/\s/.test(s[p])){p++;continue;}
      const ch=s[p];
      if(ch==='"'||ch==="'") {
        const quote=ch;let q=++p, str='';
        while(q<s.length){if(s[q]===quote){if(s[q+1]===quote){str+=quote;q+=2;continue;}break;}str+=s[q++];}
        if(q===s.length)throw error('#PARSE!','Unclosed quote');
        p=q+1;out.push({t:quote==='"'?'str':'sheet',v:str});continue;
      }
      if(ch==='#'){
        const m=s.slice(p).match(/^#[A-Z0-9/!?]+/i);
        if(!m)throw error('#PARSE!','Unknown error token');out.push({t:'err',v:m[0]});p+=m[0].length;continue;
      }
      if(/\d/.test(ch) || (ch==='.'&&/\d/.test(s[p+1]||''))){
        const m=s.slice(p).match(/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+\-]?\d+)?/);
        out.push({t:'num',v:Number(m[0])});p+=m[0].length;continue;
      }
      if(ops.includes(ch)){
        const two=s.slice(p,p+2);if(['>=','<=','<>'].includes(two)){out.push({t:'op',v:two});p+=2;}else{out.push({t:'op',v:ch});p++;}continue;
      }
      let q=p;while(q<s.length&&!ops.includes(s[q])&&!/\s/.test(s[q])&&s[q]!=='"'&&s[q]!=="'")q++;
      if(q===p)throw error('#PARSE!','Unexpected '+s[p]);
      out.push({t:'id',v:s.slice(p,q)});p=q;
    }
    out.push({t:'eof',v:'EOF'});return out;
  }
  function parse(s) {
    const a=tokenizer(s);let i=0;
    const get=()=>a[i],take=()=>a[i++];
    const want=(v)=>{if(get().v!==v)throw error('#PARSE!','Expected '+v+' got '+get().v);i++;};
    const prec={'=':1,'<>':1,'<':1,'>':1,'<=':1,'>=':1,'&':2,'+':3,'-':3,'*':4,'/':4,'^':5,':':6};
    function primary(){
      const t=take();
      if(t.v==='+'||t.v==='-')return {type:'un',op:t.v,x:expr(5)};
      if(t.v==='('){const r=expr(0);want(')');return r;}
      if(t.v==='['){const id=take();want(']');want('!');const member=take();return {type:'external',book:String(id.v),name:member.v};}
      if(t.t==='num'||t.t==='str')return {type:'lit',v:t.v};
      if(t.t==='err')return {type:'error',code:t.v};
      if(t.t==='id'||t.t==='sheet'){
        if(get().v==='('){take();const args=[];while(get().v!==')'){if(get().t==='eof')throw error('#PARSE!','Unclosed function');args.push(expr(0));if(get().v===','||get().v===';'){take();continue;}break;}want(')');return {type:'call',name:t.v.toUpperCase(),args};}
        if(get().v==='!'){take();const ref=take();return {type:'ref',sheet:t.v,key:ref.v};}
        return {type:'ref',key:t.v};
      }
      throw error('#PARSE!','Unexpected token '+t.v);
    }
    function expr(min){let left=primary();
      while(get().v==='%' ){take();left={type:'binary',op:'/',left,right:{type:'lit',v:100}};}
      while(prec[get().v]!==undefined && prec[get().v]>=min){
        const op=take().v;const right=expr(prec[op]+(op==='^'?0:1));left={type:'binary',op,left,right};
      }
      return left;
    }
    const ast=expr(0);if(get().t!=='eof')throw error('#PARSE!','Trailing token '+get().v);
    return ast;
  }
  function colNum(c){let n=0;for(const ch of c.toUpperCase())n=n*26+(ch.charCodeAt(0)-64);return n;}
  function colName(n){let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;}
  function refParts(ref){const m=String(ref).replace(/\$/g,'').match(/^([A-Z]+)(\d+)$/i);return m?{c:colNum(m[1]),r:Number(m[2])}:null;}
  function quoteSheet(x){return "'"+x.replace(/'/g,"''")+"'!";}
  function normalizeTables(formula,model,sheet,cell){
    const names=Object.keys(model.tables||{}).sort((a,b)=>b.length-a.length);
    let str=formula;
    for(const name of names){
      // Regex only identifies the start; bracket parser supports nested references.
      const needle=name+'[';let at=0;
      while((at=str.indexOf(needle,at))>=0){
        if(at>0 && /[\wА-Яа-я]/.test(str[at-1])){at+=needle.length;continue;}
        let depth=0,end=at+name.length,quoted=false;
        for(;end<str.length;end++){if(str[end]==='[')depth++;if(str[end]===']' && --depth===0 && str[end]===']'){end++;break;}if(str[end]!==']')depth+=0;}
        // Safer balanced search than the above condition for nested brackets
        depth=0;end=at+name.length;
        for(;end<str.length;end++){if(str[end]==='[')depth++;else if(str[end]===']'){depth--;if(depth===0){end++;break;}}}
        if(depth!==0){at+=needle.length;continue;}
        const raw=str.slice(at+name.length,end); const tbl=model.tables[name];
        const start=refParts(tbl.ref.split(':')[0]),stop=refParts(tbl.ref.split(':')[1]);
        const row=Number(String(cell).match(/\d+/)?.[0]||0);
        let column='',thisRow=false,totals=false;
        if(raw==='[]'){ column=''; }
        else if(raw.startsWith('[[#This Row],[')&&raw.endsWith(']]')){
          column=raw.slice('[[#This Row],['.length,-2);thisRow=true;
        }else if(raw.startsWith('[[#Totals],[')&&raw.endsWith(']]')){
          column=raw.slice('[[#Totals],['.length,-2);totals=true;
        }else if(raw.startsWith('[[')&&raw.endsWith(']]'))column=raw.slice(2,-2);
        else if(raw.startsWith('[')&&raw.endsWith(']'))column=raw.slice(1,-1);
        const norm=s=>String(s).replace(/\s+/g,' ').trim().toLowerCase();
        const idx=tbl.columns.findIndex(x=>norm(x)===norm(column));
        let replacement;
        if(!column)replacement=quoteSheet(tbl.sheet)+colName(start.c)+String(start.r+1)+':'+colName(stop.c)+String(stop.r-tbl.totals);
        else if(idx<0)replacement='#REF!';
        else {
          const col=colName(start.c+idx);
          replacement=quoteSheet(tbl.sheet)+col+(totals?stop.r:thisRow?row:start.r+1);
          if(!thisRow&&!totals)replacement+=':'+col+String(stop.r-tbl.totals);
        }
        str=str.slice(0,at)+replacement+str.slice(end);at+=replacement.length;
      }
    }
    return str;
  }
  class Model {
    constructor(data){if(!data||data.format!=='a19-excel-model-v1')throw new Error('Incorrect A19 model');this.data=data;this.overrides={...(data.overrides||{})};this.cache=new Map();this.ast=new Map();this.warnings=[];this.details=new Map();this.active=new Set();}
    clear(){this.cache.clear();this.details.clear();this.active.clear();}
    edit(sheet,cell,value){if(!this.data.sheets[sheet])throw new Error('Unknown sheet');const key=sheet+'!'+cell.toUpperCase().replace(/\$/g,'');this.overrides[key]=value;this.clear();}
    restore(sheet,cell){delete this.overrides[sheet+'!'+cell.toUpperCase()];this.clear();}
    get(sheet,cell){const key=String(cell).replace(/\$/g,'').toUpperCase();const id=sheet+'!'+key;
      if(Object.prototype.hasOwnProperty.call(this.overrides,id))return this.overrides[id];
      if(this.cache.has(id))return this.cache.get(id);
      if(this.active.has(id))return error('#CIRC!','Circular reference '+id);
      const entry=this.data.sheets[sheet]?.[key];
      if(!entry)return null;
      if(!entry.f){return entry.v && /^#(REF|N\/A|DIV\/0|VALUE|NAME)/.test(entry.v)&&entry.error?error(entry.v):entry.v;}
      this.active.add(id);let value;
      try {
        let ast=this.ast.get(id);
        if(!ast){const str=normalizeTables(entry.f,this.data,sheet,key);ast=parse(str);this.ast.set(id,ast);}
        value=this.eval(ast,sheet);
        if(typeof value==='number'&&!Number.isFinite(value))throw error('#NUM!');
      }catch(e){value=e instanceof ExcelError?e:error('#PARSE!',e.message);}
      this.active.delete(id);this.cache.set(id,value);
      if(value instanceof ExcelError)this.details.set(id,value);
      return value;
    }
    getRange(sheet,start,end){const a=refParts(start),b=refParts(end);if(!a||!b||a.c>b.c||a.r>b.r)throw error('#REF!','Invalid range '+start+':'+end);
      if((b.c-a.c+1)*(b.r-a.r+1)>100000)throw error('#VALUE!','Excessive range');
      const arr=[];for(let r=a.r;r<=b.r;r++){const row=[];for(let c=a.c;c<=b.c;c++)row.push(this.get(sheet,colName(c)+r));arr.push(row);}return arr;
    }
    getName(name,scope,external){
      if(external){const val=this.data.externalNames?.[external]?.[name];if(!val)throw error('#N/A','Missing external name '+external+'!'+name);return this.externalEval(val,external);}
      const raw=this.data.localNames?.[scope]?.[name]??this.data.names?.[name];
      if(raw===undefined){
        if(name.toUpperCase()==='TRUE')return true;
        if(name.toUpperCase()==='FALSE')return false;
        throw error('#NAME?','Missing name '+name);
      }
      if(raw==='#REF!'||raw.includes('#REF!'))throw error('#REF!','Broken defined name '+name);
      if(raw.startsWith('='))return this.eval(parse(raw.slice(1)),scope);
      return this.eval(parse(raw),scope);
    }
    externalEval(value,external){
      const m=String(value).replace(/^=/,'').match(/^'?([^']+)'?!\$?([A-Z]+)\$?(\d+)$/i);
      if(!m)throw error('#REF!','External defined name not resolvable '+value);
      const cell=m[2].toUpperCase()+m[3],sheet=m[1]; const v=this.data.external?.[external]?.[sheet]?.[cell];
      if(!v)throw error('#N/A','No external cache for '+external+':'+sheet+'!'+cell);
      return v.v;
    }
    resolve(n,scope){
      if(n.type!=='ref')throw error('#REF!','Expected reference');
      const sheet=n.sheet||scope; const key=String(n.key).replace(/\$/g,'');
      if(refParts(key))return {sheet,cell:key.toUpperCase()};
      return {name:key,scope:sheet};
    }
    eval(ast,scope){
      switch(ast.type){
        case 'lit':return ast.v;
        case 'error':throw error(ast.code);
        case 'external':return this.getName(ast.name,scope,ast.book);
        case 'ref':{const r=this.resolve(ast,scope);return r.name?this.getName(r.name,r.scope):this.get(r.sheet,r.cell);}
        case 'un':return ast.op==='-'?-number(this.eval(ast.x,scope)):number(this.eval(ast.x,scope));
        case 'binary':{
          if(ast.op===':'){
            const a=this.resolve(ast.left,scope),b=this.resolve(ast.right,ast.right.sheet?scope:a.sheet);
            if(!a.cell||!b.cell||a.sheet!==b.sheet)throw error('#REF!','Invalid range');
            return this.getRange(a.sheet,a.cell,b.cell);
          }
          const x=this.eval(ast.left,scope),y=this.eval(ast.right,scope);
          if(x instanceof ExcelError)throw x;if(y instanceof ExcelError)throw y;
          switch(ast.op){
            case '+':return number(x)+number(y);case '-':return number(x)-number(y);
            case '*':return number(x)*number(y);case '/':{const d=number(y);if(!d)throw error('#DIV/0!');return number(x)/d;}
            case '^':return Math.pow(number(x),number(y));case '&':return text(x)+text(y);
            case '=':return comparable(x)===comparable(y);case '<>':return comparable(x)!==comparable(y);
            case '<':return comparable(x)<comparable(y);case '>':return comparable(x)>comparable(y);
            case '<=':return comparable(x)<=comparable(y);case '>=':return comparable(x)>=comparable(y);
          }throw error('#PARSE!','Unknown operator '+ast.op);
        }
        case 'call':return this.call(ast.name,ast.args,scope);
      }throw error('#PARSE!','Unknown AST type');
    }
    call(name,args,scope){
      const get=i=>this.eval(args[i],scope);
      const flat=x=>Array.isArray(x)?x.flat(Infinity):[x];
      if(name==='IF'){return truth(get(0))?get(1):(args.length>2?get(2):false);}
      if(name==='IFERROR'||name==='IFNA'){try{const x=get(0);if(x instanceof ExcelError)throw x;return x;}catch(e){if(name==='IFERROR'||(e instanceof ExcelError&&e.code==='#N/A'))return get(1);throw e;}}
      if(name==='ISNA'||name==='ISERR'||name==='ISERROR'){
        try{const x=get(0);if(x instanceof ExcelError)throw x;return false;}catch(e){if(!(e instanceof ExcelError))throw e;return name==='ISNA'?e.code==='#N/A':name==='ISERR'?e.code!=='#N/A':true;}
      }
      if(name==='ISNUMBER'){try{const x=get(0);return typeof x==='number'&&Number.isFinite(x);}catch(e){if(e instanceof ExcelError)return false;throw e;}}
      if(name==='AND'){for(let i=0;i<args.length;i++)if(!truth(get(i)))return false;return true;}
      if(name==='OR'){for(let i=0;i<args.length;i++)if(truth(get(i)))return true;return false;}
      if(name==='SUM'||name==='SUBTOTAL'){
        const skip=name==='SUBTOTAL'?1:0;
        if(skip&&![9,109].includes(number(get(0))))throw error('#VALUE!','Unsupported SUBTOTAL function code');
        let sum=0;for(let i=skip;i<args.length;i++)for(const x of flat(get(i)))if(typeof x==='number'&&Number.isFinite(x))sum+=x;
        return sum;
      }
      if(name==='ROUNDUP'){
        const v=number(get(0)),digits=number(get(1)),factor=Math.pow(10,digits);
        return Math.sign(v)*Math.ceil(Math.abs(v)*factor-1e-12)/factor;
      }
      if(name==='INDEX'){
        const list=get(0),r=number(get(1)),c=args.length>2?number(get(2)):1;
        if(!Array.isArray(list))throw error('#REF!');const row=list[r-1];
        if(!row||row[c-1]===undefined)throw error('#REF!');return row[c-1];
      }
      if(name==='SEARCH'){
        const needle=text(get(0)).toLowerCase(),hay=text(get(1)).toLowerCase(),start=args.length>2?number(get(2))-1:0;
        const n=hay.indexOf(needle,start);if(n<0)throw error('#VALUE!');return n+1;
      }
      if(name==='VLOOKUP'){
        const key=get(0),rows=get(1),index=number(get(2)),approx=args.length>3?truth(get(3)):true;
        if(!Array.isArray(rows)||!index||index<1)throw error('#REF!');
        const matches=[];let last=null;
        for(const row of rows){if(!Array.isArray(row)||row.length<index)continue;
          const v=row[0];if(v instanceof ExcelError)continue;
          if(comparable(v)===comparable(key))return row[index-1];
          if(approx&&comparable(v)<=comparable(key))last=row;
        }
        if(approx&&last)return last[index-1];throw error('#N/A','VLOOKUP did not match '+text(key));
      }
      throw error('#NAME?','Unsupported function '+name);
    }
    audit(){this.clear();const count={formulas:0,ok:0,error:0,parse:0,reference:0,cachedError:0,diff:0,diffExamples:[],errors:{}};
      for(const [sheet,cells] of Object.entries(this.data.sheets))for(const [a,c] of Object.entries(cells)){
        if(!c.f)continue;count.formulas++;const value=this.get(sheet,a);
        if(value instanceof ExcelError){count.error++;count.errors[value.code]=(count.errors[value.code]||0)+1;if(value.code==='#PARSE!')count.parse++;if(value.code==='#REF!')count.reference++;}
        else count.ok++;
        if(c.error)count.cachedError++;
        if(typeof c.cache==='number'&&typeof value==='number'&&Number.isFinite(value) && Math.abs(value-c.cache)>Math.max(.01,Math.abs(c.cache)*.000001)){
          count.diff++;if(count.diffExamples.length<12)count.diffExamples.push({sheet,cell:a,expected:c.cache,actual:value});
        }
      }
      return count;
    }
    serialize(){return {...this.data,overrides:{...this.overrides}};}
  }
  return {Model,ExcelError,parse,normalizeTables,refParts};
});
