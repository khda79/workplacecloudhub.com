// Offline PBIR migration. Operates only on an explicitly selected candidate copy.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import assert from 'node:assert/strict';
import {fileURLToPath} from 'node:url';
import {retiredTables} from './Retire-CmdbLayer.mjs';
const read=p=>JSON.parse(fs.readFileSync(p,'utf8').replace(/^\uFEFF/,''));
const clone=x=>structuredClone(x);
const id=s=>crypto.createHash('sha256').update('workplace-consolidation-v1:'+s).digest('hex').slice(0,20);
const lit=v=>({expr:{Literal:{Value:typeof v==='boolean'?String(v):typeof v==='number'?`${v}D`:`'${v.replaceAll("'","''")}'`}}});
const hasRetired=x=>retiredTables.some(t=>(JSON.stringify(x)||'').includes(t));
const destination=f=>f.page!=='CMDB'?f.page:f.source==='FactRelationshipOverview'?'Data Trust':'User Explorer';
export function reportPlan(rootPath){
  assert(rootPath,'Explicit existing Report directory required');
  const root=path.resolve(rootPath),def=path.join(root,'definition'),pagesDir=path.join(def,'pages'),marksDir=path.join(def,'bookmarks');
  const order=read(path.join(pagesDir,'pages.json'));
  const pages=order.pageOrder.map(p=>read(path.join(pagesDir,p,'page.json')));
  const cmdb=pages.find(p=>p.displayName==='CMDB');assert(cmdb,'Expected CMDB page; refusing repeated migration');
  const pageByName=new Map(pages.map(p=>[p.displayName,p]));
  const docs=new Map(),before=new Map();
  function walk(dir){for(const e of fs.readdirSync(dir,{withFileTypes:true})){const p=path.join(dir,e.name);if(e.isDirectory())walk(p);else if(p.endsWith('.json')){const s=fs.readFileSync(p,'utf8');before.set(p,s);docs.set(p,JSON.parse(s));}}}
  walk(def);
  const visualPath=(p,n)=>path.join(pagesDir,p,'visuals',n,'visual.json');
  const visualEntries=p=>[...docs].filter(([f])=>f.startsWith(path.join(pagesDir,p,'visuals')+path.sep));
  const facets=read(new URL('../config/consolidation-report-facets.json.txt',import.meta.url)).facets;
  const entries=facets.map(f=>{
    // Destination config can be updated after this migration without losing legacy IDs.
    let sourcePage=pageByName.get(f.page),g=id(sourcePage.name+':Native · '+f.label);
    if(!docs.has(visualPath(sourcePage.name,g))){sourcePage=cmdb;g=id(cmdb.name+':Native · '+f.label);}
    assert(docs.has(visualPath(sourcePage.name,g)),'Missing facet group '+f.source);
    const to=pageByName.get(destination(f));
    return {f,from:sourcePage.name,to:to.name,g,bm:id(sourcePage.name+':bookmark:Native · '+f.label)};
  });
  const moved=new Map(entries.filter(e=>e.from!==e.to).map(e=>[e.bm,e]));
  const removedVisuals=new Set(),removedMarks=new Set();
  const buttonTemplate=clone(docs.get(visualPath('a821f7deb92bb3bc4c78','34c99d7e680872d936e7')));
  assert(buttonTemplate,'Missing approved navigation template');
  // Move complete domain groups, with the same fields and IDs; do not copy generic graph visuals.
  for(const e of entries){
    if(e.from===e.to)continue;
    for(const [file,v]of visualEntries(e.from).filter(([,v])=>v.name===e.g||v.parentGroupName===e.g)){
      assert(!docs.has(visualPath(e.to,v.name)),'Visual ID collision');docs.delete(file);docs.set(visualPath(e.to,v.name),v);
    }
  }
  for(const [file,v]of [...docs]){
    if(file.startsWith(path.join(pagesDir,cmdb.name)+path.sep)){if(v.name)removedVisuals.add(v.name);docs.delete(file);continue;}
    if(file.endsWith('visual.json')&&v.visual?.visualType==='slicer'&&hasRetired(v.visual.query)){
      removedVisuals.add(v.name);docs.delete(file);
    }
  }
  // Regenerate only the native subview navigation for pages receiving migrated details.
  const receiving=new Set(entries.filter(e=>e.from!==e.to).map(e=>e.to));
  for(const p of receiving){
    const mine=entries.filter(e=>e.to===p),rows=Math.ceil(mine.length/6),width=Math.floor((1856-5*8)/6);
    for(const e of mine){
      const content=visualEntries(p).filter(([,v])=>v.parentGroupName===e.g);
      for(const [file,v]of content){
        if(v.visual?.visualType==='actionButton'){removedVisuals.add(v.name);docs.delete(file);continue;}
        if(v.visual?.visualType==='slicer'){
          const oldY=v.position.y;v.position.y=344+rows*48;v.position.width=1856;
          if(e.f.source==='EntraGroupMembership'||e.f.source==='ADMembership'){
            const key=e.f.source==='ADMembership'?'TenantADGroupKey':'TenantGroupKey';
            const q=v.visual.query.queryState.Values.projections[0];q.field.Column.Property=key;q.queryRef=e.f.table+'.'+key;q.nativeQueryRef=key;
            v.visual.objects.header[0].properties.text=lit('Group key · search / select');
            delete v.visual.objects.general;
          }
          for(const [otherFile,other]of content){
            if(other.visual?.visualType==='tableEx'&&other.position.y===oldY+96){const bottom=other.position.y+other.position.height;other.position.y=v.position.y+96;other.position.height=bottom-other.position.y;docs.set(otherFile,other);}
          }
        }
        docs.set(file,v);
      }
      mine.forEach((nav,j)=>{
        const v=clone(buttonTemplate);v.name=id('retirement:'+e.g+':nav:'+j);v.parentGroupName=e.g;delete v.isHidden;delete v.filterConfig;
        v.position={x:32+(j%6)*(width+8),y:344+Math.floor(j/6)*48,width,height:40,z:12000,tabOrder:12000+j};
        for(const t of v.visual.objects.text)t.properties.text=lit(nav.f.label);
        v.visual.visualContainerObjects.general=[{properties:{altText:lit(nav.f.label)}}];
        v.visual.visualContainerObjects.visualLink=[{properties:{show:lit(true),type:lit('Bookmark'),bookmark:lit(nav.bm)}}];
        docs.set(visualPath(p,v.name),v);
      });
    }
  }
  // All bookmarks keep their data suppression; migrated facets join the destination view switch.
  for(const [file,bm]of [...docs].filter(([f])=>f.endsWith('.bookmark.json'))){
    const oldPage=bm.explorationState?.activeSection;
    if(oldPage===cmdb.name&&!moved.has(bm.name)){removedMarks.add(bm.name);docs.delete(file);continue;}
    let p=oldPage;
    if(moved.has(bm.name)){
      const e=moved.get(bm.name);p=e.to;bm.explorationState.activeSection=p;
      bm.explorationState.sections={[p]:{visualContainers:{},visualContainerGroups:{[e.g]:{isHidden:false}}}};
    }
    if(receiving.has(p)){
      const groups=visualEntries(p).filter(([,v])=>v.visualGroup).map(([,v])=>v.name);
      const states=bm.explorationState.sections[p].visualContainerGroups||{};
      bm.explorationState.sections[p].visualContainerGroups=Object.fromEntries(groups.map(g=>[g,states[g]||{isHidden:true}]));
      bm.options.targetVisualNames=[...new Set([...(bm.options.targetVisualNames||[]).filter(n=>!removedVisuals.has(n)),...groups])];
    }
    docs.set(file,bm);
  }
  // Remove only exact generic dependencies, dead visual/bookmark references and stale explanatory wording.
  function clean(v,key=''){
    if(typeof v==='string')return v.replace('country, service and CMDB entity selections','country and service selections');
    if(Array.isArray(v))return v.filter(x=>!(typeof x==='string'&&(removedVisuals.has(x)||removedMarks.has(x)||x===cmdb.name)))
      .filter(x=>!(key==='filters'||key==='parameters')||!hasRetired(x))
      .filter(x=>key!=='visualInteractions'||(!removedVisuals.has(x.source)&&!removedVisuals.has(x.target)))
      .map(x=>clean(x,key));
    if(!v||typeof v!=='object')return v;
    const result={};for(const[k,x]of Object.entries(v))if(k!==cmdb.name&&!removedVisuals.has(k))result[k]=clean(x,k);
    if(result.children&&result.children.length===0&&result.displayName)return null;
    return result;
  }
  for(const [file,v]of docs){const c=clean(v);if(c?.items)c.items=c.items.filter(Boolean);docs.set(file,c);}
  const listFile=path.join(marksDir,'bookmarks.json'),list=docs.get(listFile);
  // Re-register moved bookmarks under their actual page (remove obsolete CMDB grouping).
  for(const item of list.items)item.children=item.children?.filter(n=>!moved.has(n));
  list.items=list.items.filter(x=>!x.children||x.children.length);
  for(const p of receiving){const children=[...moved].filter(([,e])=>e.to===p).map(([bm])=>bm);if(children.length)list.items.push({name:id('retirement:moved:'+p),displayName:pages.find(x=>x.name===p).displayName+' additional details',children});}
  const orderFile=path.join(pagesDir,'pages.json'),next=docs.get(orderFile);
  next.pageOrder=next.pageOrder.filter(p=>p!==cmdb.name);if(next.activePageName===cmdb.name)next.activePageName=next.pageOrder[0];
  assert(next.pageOrder.length===order.pageOrder.length-1);
  for(const [file,v]of docs){assert(!hasRetired(v),'Retired binding remains: '+file);assert(!JSON.stringify(v).includes(cmdb.name),'Retired navigation remains: '+file);}
  const files=[...docs].filter(([p,v])=>!before.has(p)||JSON.stringify(JSON.parse(before.get(p)))!==JSON.stringify(v)).map(([file,value])=>({file,value,before:before.get(file)??null}));
  return {files,deletes:[...before.keys()].filter(p=>!docs.has(p)),moved:entries.filter(e=>e.from!==e.to).map(e=>({table:e.f.table,page:pages.find(p=>p.name===e.to).displayName})),pageCount:next.pageOrder.length};
}
if(process.argv[1]===fileURLToPath(import.meta.url)){
  const root=process.argv[2],plan=reportPlan(root);
  if(process.argv[3]==='--apply'){
    assert(path.resolve(root).toLowerCase().includes(path.sep+'candidates'+path.sep),'Writes are restricted to an isolated candidate');
    for(const f of plan.files){if(f.before!==null)assert.equal(fs.readFileSync(f.file,'utf8'),f.before,'Concurrent edit');}
    for(const f of plan.files){fs.mkdirSync(path.dirname(f.file),{recursive:true});fs.writeFileSync(f.file,JSON.stringify(f.value,null,2)+'\n');}
    for(const f of plan.deletes)fs.unlinkSync(f);
    // Remove only now-empty candidate directories; never recursively erase a tree.
    function emptyDirs(dir){for(const e of fs.readdirSync(dir,{withFileTypes:true}))if(e.isDirectory())emptyDirs(path.join(dir,e.name));if(!fs.readdirSync(dir).length)fs.rmdirSync(dir);}
    emptyDirs(path.join(path.resolve(root),'definition','pages'));
    console.log(JSON.stringify({files:plan.files.length,deletes:plan.deletes.length,moved:plan.moved,pageCount:plan.pageCount}));
  }else console.log(JSON.stringify({files:plan.files.length,deletes:plan.deletes.length,moved:plan.moved,pageCount:plan.pageCount}));
}
