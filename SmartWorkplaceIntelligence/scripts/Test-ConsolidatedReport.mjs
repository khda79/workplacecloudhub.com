// Static source/model metadata validation only. Never opens Desktop or reads tenant CSVs.
import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
import {definitions} from './Build-ConsolidatedContext.mjs';
const root=path.resolve(process.argv[2]||''), metadataPath=process.argv[3];
assert(process.argv[2]&&metadataPath,'Explicit Report root and MCP metadata projection are required');
const read=p=>JSON.parse(fs.readFileSync(p,'utf8').replace(/^\uFEFF/,''));
const def=path.join(root,'definition'),order=read(path.join(def,'pages/pages.json'));
const snapshot=read(metadataPath);
const metadata=snapshot.model?{tables:snapshot.model.tables.map(t=>({name:t.name,columns:t.columns.map(c=>c.name),measures:(t.measures||[]).map(m=>m.name)})),relationships:snapshot.model.relationships}:snapshot;
const tables=new Map(metadata.tables.map(t=>[t.name,t]));
let checks=0;
function check(test,message){assert(test,message);checks++;}
check(order.pageOrder.length===30,'Expected 30 pages including On-prem Messaging');
check(new Set(order.pageOrder).size===30,'Duplicate page registration');
check(order.pageOrder.indexOf('on-prem-messaging')===order.pageOrder.indexOf('71a60a59e0eb52a7a7d0')+1,'On-prem Messaging must follow Messaging & Hybrid');
check(order.pageOrder.includes('migration-to-sharepoint-online'),'SharePoint migration page preserved');
const pages=order.pageOrder.map(p=>read(path.join(def,'pages',p,'page.json')));
check(!pages.some(p=>p.displayName==='CMDB')&&pages[1].displayName==='Operational Overview','Existing business pages preserved');
const all=[];
function walk(dir){for(const entry of fs.readdirSync(dir,{withFileTypes:true})){const p=path.join(dir,entry.name);if(entry.isDirectory())walk(p);else if(p.endsWith('.json'))all.push({file:p,value:read(p)});}}
walk(def);
function bindings(value,file){
 if(!value||typeof value!=='object')return;
 for(const kind of ['Column','Measure'])if(value[kind]?.Expression?.SourceRef?.Entity){
  const item=value[kind],entity=item.Expression.SourceRef.Entity,t=tables.get(entity);
  check(!!t,`Unknown table ${entity} in ${file}`);
  check((kind==='Column'?t.columns:t.measures).includes(item.Property),`Unknown ${kind} ${entity}.${item.Property} in ${file}`);
 }
 for(const item of Object.values(value))if(item&&typeof item==='object')bindings(item,file);
}
for(const item of all)bindings(item.value,item.file);
const facets=read(new URL('../config/consolidation-report-facets.json.txt',import.meta.url)).facets;
const visuals=all.filter(x=>x.file.endsWith(path.join('visual.json'))).map(x=>x.value);
const trustPage=pages.find(p=>p.displayName==='Data Trust');
const trustBanner=all.find(x=>x.file.startsWith(path.join(def,'pages',trustPage.name,'visuals')+path.sep)
 && x.value.visual?.visualContainerObjects?.title?.some(t=>t.properties?.text?.expr?.Measure?.Property==='Loaded Batch Banner'));
check(!!trustBanner,'Data Trust loaded-batch diagnostic must be bound');
check(!trustBanner.value.parentGroupName,'Loaded-batch diagnostic must remain outside bookmark-switched groups');
check(trustBanner.value.visual.visualContainerObjects.title[0].properties.show.expr.Literal.Value==='true','Loaded-batch diagnostic must be visible');
for(const facet of facets){
 const page=pages.find(p=>p.displayName===facet.page);
 check(!!page,`Missing destination page ${facet.page}`);
 const pageVisuals=all.filter(x=>x.file.startsWith(path.join(def,'pages',page.name,'visuals')+path.sep)).map(x=>x.value);
 const projections=pageVisuals.flatMap(v=>v.visual?.query?.queryState?.Values?.projections||[]).filter(p=>p.field?.Column?.Expression?.SourceRef?.Entity===facet.table).map(p=>p.field.Column.Property);
 for(const col of facet.columns)check(projections.includes(col),`Missing native field ${facet.table}.${col}`);
}
const bookmarks=all.filter(x=>x.file.endsWith('.bookmark.json')).map(x=>x.value);
for(const page of pages.filter(p=>['User Explorer','Data Trust'].includes(p.displayName))){
 const pageVisuals=all.filter(x=>x.file.startsWith(path.join(def,'pages',page.name,'visuals')+path.sep)).map(x=>x.value);
 const groups=pageVisuals.filter(v=>v.visualGroup).map(v=>v.name);
 for(const bm of bookmarks.filter(b=>b.explorationState?.activeSection===page.name)){
  const states=bm.explorationState.sections[page.name].visualContainerGroups||{};
  for(const g of groups)check(Object.hasOwn(states,g),`Bookmark ${bm.name} misses group ${g}`);
  check(groups.filter(g=>states[g]?.isHidden===false).length===1,`Bookmark ${bm.name} must show one content group`);
 }
 for(const v of pageVisuals.filter(v=>v.parentGroupName))check(groups.includes(v.parentGroupName),`Orphan visual ${v.name}`);
}
check(!visuals.some(v=>v.visual?.syncGroup?.groupName==='CMDB native entity context'),'Generic entity sync must be retired');
check(!JSON.stringify(all).match(/CMDB Entity|CMDB Relationship|CMDB Link In Context/),'Retired graph report binding');
const context=definitions();
check(context.updates.length===190&&context.filters.length===45,'Native measure/context coverage');
for(const m of [...context.updates,...context.filters])check(!m.expression.includes('CMDB Entity'),'Generic graph must not return through regeneration');
console.log(`Consolidated report/model metadata checks passed: ${checks}. Static validation only; no DAX execution or Desktop rendering.`);
