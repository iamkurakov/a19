#!/usr/bin/env python3
"""Produce a PRIVATE JSON workbook model for local A19 browser import.

No third-party packages needed; reads Office Open XML through stdlib.
Do not commit output to public repository: it contains actual company prices.
"""
import sys,json,zipfile,xml.etree.ElementTree as ET,re,os
N='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

def main(src,dest):
 z=zipfile.ZipFile(src)
 wb=ET.fromstring(z.read('xl/workbook.xml'))
 rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
 relmap={x.attrib['Id']:x.attrib['Target'].lstrip('/') for x in rels}
 ss=[]
 if 'xl/sharedStrings.xml' in z.namelist():
  ss=[''.join(x.text or '' for x in a.findall('.//'+N+'t')) for a in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(N+'si')]
 def atom(c):
  t=c.attrib.get('t','n');v=c.find(N+'v');ins=c.find(N+'is')
  v=(v.text if v is not None and v.text is not None else '')
  if t=='s' and v:return ss[int(v)]
  if t=='inlineStr' and ins is not None:return ''.join(x.text or '' for x in ins.findall('.//'+N+'t'))
  if t=='b':return v=='1'
  if t in ('str','e','inlineStr'):return v
  if v!='':
   try:return float(v)
   except ValueError:pass
  return None
 sheets={};sheetnames=[];sheetpaths={}
 for sh in wb.find(N+'sheets'):
  name=sh.attrib['name'];sheetnames.append(name)
  target=relmap[sh.attrib[R+'id']];path=target if target.startswith('xl/') else 'xl/'+target
  root=ET.fromstring(z.read(path));cells={};sheetpaths[name]=path
  for c in root.findall('.//'+N+'sheetData/'+N+'row/'+N+'c'):
   formula=c.find(N+'f');value=atom(c);v=c.find(N+'v')
   if formula is None and value is None:continue
   o={'v':value}
   if formula is not None and formula.text is not None:o['f']=formula.text
   if formula is not None and value is not None:o['cache']=value
   if c.attrib.get('t')=='e':o['error']=True
   cells[c.attrib['r']]=o
  sheets[name]=cells
 names={};locals={}
 for d in wb.findall('.//'+N+'definedName'):
  key=d.attrib.get('name'); val=d.text or ''
  if not key:continue
  scope=d.attrib.get('localSheetId')
  if scope is None:names[key]=val
  else:locals.setdefault(sheetnames[int(scope)],{})[key]=val
 tables={}
 for file in z.namelist():
  if not re.match(r'xl/tables/table\d+\.xml$',file):continue
  t=ET.fromstring(z.read(file)); name=t.attrib['name'];ref=t.attrib['ref']
  head=t.find(N+'tableColumns')
  if head is None:continue
  columns=[c.attrib['name'].replace('_x000a_','\n') for c in head]
  # Table name-> worksheet located via the header cell content position (all tables unique)
  begin=ref.split(':')[0];col=re.match(r'[A-Z]+',begin).group();row=int(re.search(r'\d+',begin).group())
  matches=[]
  for sn,shpath in sheetpaths.items():
   relpath=shpath.rsplit('/',1)[0]+'/_rels/'+shpath.rsplit('/',1)[1]+'.rels'
   if relpath not in z.namelist():continue
   relation=ET.fromstring(z.read(relpath))
   for r in relation:
    if r.attrib.get('Type','').endswith('/table') and r.attrib.get('Target','').split('/')[-1]==file.split('/')[-1]:matches.append(sn)
  if len(matches)!=1:raise ValueError('ambiguous table location '+name+': '+repr(matches))
  tables[name]={'sheet':matches[0],'ref':ref,'columns':columns,'totals':int(t.attrib.get('totalsRowCount',0))}
 external={};extNames={}
 for idx in [1,2,3]:
  path=f'xl/externalLinks/externalLink{idx}.xml'
  if path not in z.namelist():continue
  r=ET.fromstring(z.read(path));
  shNames=[s.attrib['val'] for s in r.findall('.//'+N+'sheetName')]
  extSheets={}
  for sh in r.findall('.//'+N+'sheetData'):
   num=int(sh.attrib.get('sheetId','0'));name=shNames[num] if num<len(shNames) else str(num)
   sc={}
   for c in sh.findall('.//'+N+'cell'):
    val=atom(c)
    if val is not None:sc[c.attrib['r']]={'v':val}
   extSheets[name]=sc
  external[str(idx)]=extSheets
  extNames[str(idx)]={x.attrib['name']:x.attrib.get('refersTo','') for x in r.findall('.//'+N+'definedName')}
 model={'format':'a19-excel-model-v1','source':'uploaded client XLSX','sheets':sheets,'sheetOrder':sheetnames,'names':names,'localNames':locals,'tables':tables,'external':external,'externalNames':extNames,'overrides':{}}
 with open(dest,'w',encoding='utf-8') as f:json.dump(model,f,ensure_ascii=False,separators=(',',':'))
 print(json.dumps({'sheets':len(sheets),'cells':sum(len(x) for x in sheets.values()),'formulas':sum(bool(c.get('f')) for s in sheets.values() for c in s.values()),'tables':len(tables),'externalCached':sum(len(s) for book in external.values() for s in book.values()),'bytes':os.path.getsize(dest)},ensure_ascii=False))
if __name__=='__main__':
 if len(sys.argv)!=3:sys.exit('usage: python tools/export_excel_model.py source.xlsx private-model.json')
 main(sys.argv[1],sys.argv[2])
