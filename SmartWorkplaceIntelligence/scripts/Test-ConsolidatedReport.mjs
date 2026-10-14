// Static source/model metadata validation only. Never opens Desktop or reads tenant CSVs.
import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
import {definitions} from './Build-ConsolidatedContext.mjs';
const root=path.resolve(process.argv[2]||''), metadataPath=process.argv[3];
assert(process.argv[2]&&metadataPath,'Explicit Report root and MCP metadata projection are required');
const read=p=>JSON.parse(fs.readFileSync(p,'utf8').replace(/^\uFEFF/,''));
const def=path.join(root,'definition'),order=read(path.join(def,'pages/pages.json'));
const metadata=read(metadataPath),tables=new Map(metadata.tables.map(t=>[t.name,t]));
let checks=0;
function check(test,message){assert(test,message);checks++;}
check(order.pageOrder.length===28,'Expected 28 pages');
check(new Set(order.pageOrder).size===28,'Duplicate page registration');
const pages=order.pageOrder.map(p=>read(path.join(def,'pages',p,'page.json')));
check(pages[2].displayName==='CMDB'&&pages[1].displayName==='Operational Overview','CMDB placement');
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
for(const facet of facets){
 const projections=visuals.flatMap(v=>v.visual?.query?.queryState?.Values?.projections||[]).filter(p=>p.field?.Column?.Expression?.SourceRef?.Entity===facet.table).map(p=>p.field.Column.Property);
 for(const col of facet.columns)check(projections.includes(col),`Missing native field ${facet.table}.${col}`);
}
for(const v of visuals.filter(v=>v.visual?.syncGroup?.groupName==='CMDB native entity context')){
 check(v.visual.query.queryState.Values.projections[0].field.Column.Expression.SourceRef.Entity==='CMDB Entity Evidence','Mixed sync entities');
 check(v.visual.objects.selection[0].properties.singleSelect.expr.Literal.Value==='true','Entity context must be single-select');
}
check(visuals.some(v=>JSON.stringify(v.filterConfig||{}).includes('CMDB Link In Context')),'Relationship visual lacks explicit context filter');
check(metadata.relationships.filter(r=>r.name.startsWith('cmdb-')).every(r=>r.crossFilteringBehavior!=='bothDirections'&&r.crossFilteringBehavior!=='BothDirections'),'Native graph must not use bidirectional shortcuts');
const context=definitions();
check(context.updates.length===190&&context.filters.length===45,'Native measure/context coverage');
for(const m of context.updates)check(m.expression.includes('OneEntity')&&m.expression.includes('BLANK()'),'Unqualified context must fail closed');
console.log(`Consolidated report/model metadata checks passed: ${checks}. Static validation only; no DAX execution or Desktop rendering.`);
