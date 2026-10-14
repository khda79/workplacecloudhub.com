// Emits an exact PBIR patch plan; caller applies it with the workspace patch tool.
// No data access, Desktop control, refresh, model edits or saves.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {fileURLToPath} from 'node:url';
const here=path.dirname(fileURLToPath(import.meta.url));
export function buildReport(rootPath){
const root=path.resolve(rootPath||'');
if(!rootPath||!fs.existsSync(path.join(root,'definition/pages/pages.json')))throw Error('Explicit existing Report directory required');
const def=path.join(root,'definition'), pages=path.join(def,'pages'), marks=path.join(def,'bookmarks');
const read=p=>JSON.parse(fs.readFileSync(p,'utf8').replace(/^\uFEFF/,''));
const clone=x=>structuredClone(x), id=s=>crypto.createHash('sha256').update('workplace-consolidation-v1:'+s).digest('hex').slice(0,20);
const lit=v=>({expr:{Literal:{Value:typeof v==='boolean'?String(v):typeof v==='number'?`${v}D`:`'${v.replaceAll("'","''")}'`}}});
const field=(table,column,kind='Column')=>({field:{[kind]:{Expression:{SourceRef:{Entity:table}},Property:column}},queryRef:`${table}.${column}`,nativeQueryRef:column});
const patch=new Map();
const put=(p,v)=>patch.set(p,v);
const templates=fs.readdirSync(path.join(pages,'a821f7deb92bb3bc4c78/visuals')).map(v=>read(path.join(pages,'a821f7deb92bb3bc4c78/visuals',v,'visual.json')));
const template=name=>clone(templates.find(v=>v.name===name));
const shell=(templateId,seed,position,group)=>{
  const v=template(templateId);v.name=id(seed);v.position={...v.position,...position};
  delete v.parentGroupName;delete v.isHidden;delete v.filterConfig;
  if(group)v.parentGroupName=group;
  return v;
};
function saveVisual(page,v){put(path.join(pages,page,'visuals',v.name,'visual.json'),v);return v;}
function contextFilter(v,entity,measure){
 v.filterConfig={filters:[{name:id(v.name+':context'),field:field(entity,measure,'Measure').field,type:'Advanced',filter:{Version:2,From:[{Name:'c',Entity:entity,Type:0}],Where:[{Condition:{Comparison:{ComparisonKind:2,Left:{Measure:{Expression:{SourceRef:{Source:'c'}},Property:measure}},Right:{Literal:{Value:'0L'}}}}}]}}]};
}
function entityContext(page,seed,pos,group){
 const v=slicer(page,seed,'CMDB Entity Evidence','Selection','CMDB context · clear to use inventory scope',pos,group);
 v.visual.syncGroup={groupName:'CMDB native entity context',fieldChanges:true,filterChanges:true};
 v.visual.objects.selection=[{properties:{singleSelect:lit(true)}}];return v;
}
function title(v,text){v.visual.visualContainerObjects.title=[{properties:{show:lit(true),text:lit(text),fontSize:lit(15)}}];v.visual.visualContainerObjects.general=[{properties:{altText:lit(text)}}];}
function table(page,seed,entity,columns,label,pos,group,kind='Column'){
  const v=shell('346ed8e16b6d4212dfdb',seed,pos,group);
  v.visual.query={queryState:{Values:{projections:columns.map(c=>field(entity,c,kind))}}};
  delete v.visual.objects.columnWidth;delete v.visual.objects.general;
  title(v,label);return saveVisual(page,v);
}
function text(page,seed,value,pos,group){
  const v=shell('55651abd11a3bf8eb812',seed,pos,group);
  v.visual.objects.general[0].properties.paragraphs=[{textRuns:[{value,textStyle:{fontFamily:'Segoe UI',fontSize:'14px',color:'#334E68'}}]}];
  return saveVisual(page,v);
}
function slicer(page,seed,entity,column,label,pos,group){
  const v=shell('95aecf75f46b84134141',seed,pos,group);
  v.visual.query={queryState:{Values:{projections:[{...field(entity,column),active:true}]}}};
  v.visual.objects.header[0].properties.text=lit(label);return saveVisual(page,v);
}
function button(page,seed,label,target,pos,group,type='Bookmark'){
  const v=shell('34c99d7e680872d936e7',seed,pos,group);
  for(const t of v.visual.objects.text)t.properties.text=lit(label);
  v.visual.visualContainerObjects.general=[{properties:{altText:lit(label)}}];
  v.visual.visualContainerObjects.visualLink=[{properties:{show:lit(true),type:lit(type),...(type==='Back'?{}:{[type==='Bookmark'?'bookmark':'navigationSection']:lit(target)})}}];
  return saveVisual(page,v);
}
function group(page,name,hidden=true){const v=shell('0494a34823b0fb981f15',page+':'+name,{},null);v.visualGroup.displayName=name;v.isHidden=hidden;return saveVisual(page,v).name;}
const bookmarkList=read(path.join(marks,'bookmarks.json'));
const existingMarks=fs.readdirSync(marks).filter(x=>x.endsWith('.bookmark.json')).map(f=>({file:path.join(marks,f),value:read(path.join(marks,f))}));
function bookmark(page,name,selected,groups){
  const nameId=id(page+':bookmark:'+name);
  const v={$schema:'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/bookmark/2.1.0/schema.json',name:nameId,displayName:name,
    options:{applyOnlyToTargetVisuals:true,targetVisualNames:groups,suppressActiveSection:true,suppressData:true},
    explorationState:{version:'1.0',activeSection:page,sections:{[page]:{visualContainers:{},visualContainerGroups:Object.fromEntries(groups.map(g=>[g,{isHidden:g!==selected}]))}}}};
  put(path.join(marks,nameId+'.bookmark.json'),v);return nameId;
}
const pageList=fs.readdirSync(pages,{withFileTypes:true}).filter(x=>x.isDirectory()).map(x=>read(path.join(pages,x.name,'page.json')));
const cmdb=id('CMDB');
if(pageList.some(p=>p.name===cmdb))throw Error('CMDB page already exists; do not blindly regenerate private edits');
const p=read(path.join(pages,'a821f7deb92bb3bc4c78/page.json'));
p.name=cmdb;p.displayName='CMDB';delete p.pageBinding;delete p.filterConfig;
p.filterConfig={filters:[{name:id('cmdb-country-filter'),field:field('CMDB Entity Evidence','Country').field,type:'Categorical',howCreated:'User'}]};
put(path.join(pages,cmdb,'page.json'),p);pageList.push(p);
for(const t of ['02003f4ce2fa690e0ca0','847d60f9ba460a0874bc','f5116a2458ac9c723513'])saveVisual(cmdb,shell(t,'cmdb:'+t,{},null));
const header=shell('7de0f77aaaead86f4acb','cmdb-title',{},null);
header.visual.objects.general[0].properties.paragraphs[0].textRuns[0].value='Which assets, identities and relationships make up the workplace?';saveVisual(cmdb,header);
text(cmdb,'cmdb-subtitle','CMDB · governed native identities and observed relationships · one prepared batch',{x:96,y:88,width:664,height:32});
slicer(cmdb,'cmdb-type','CMDB Entity Evidence','Entity Type','Entity type',{x:784,y:16,width:264,height:80});
slicer(cmdb,'cmdb-source','CMDB Entity Evidence','Source System','Source system',{x:1064,y:16,width:264,height:80});
entityContext(cmdb,'cmdb-selection',{x:1344,y:16,width:544,height:80});
const card=shell('524feb1a40c49d095418','cmdb-kpis',{},null);
card.visual.query.queryState.Data.projections=[field('CMDB Entity Evidence','CMDB Observed Entities','Measure'),field('CMDB Relationship Evidence','CMDB Observed Relationships','Measure'),field('CMDB Relationship Evidence','CMDB Unresolved Relationships','Measure')];
// Explicit spacing/padding remove dependence on an absent external base-theme file.
// Required height: border 2 + VCO 16 + value 42 + label 18 + spacing 2 = 80 <= 120.
// Each callout has >600 px width; full integer counts and labels fit without display-unit shortening.
card.visual.objects.spacing=[{selector:{id:'default'},properties:{verticalSpacing:lit(2)}}];saveVisual(cmdb,card);
const cmdbGroups={Inventory:group(cmdb,'Inventory',false),Relationships:group(cmdb,'Relationships'),Evidence:group(cmdb,'Evidence')};
table(cmdb,'cmdb-inventory','CMDB Entity Evidence',['Display Name','Entity Type','Source System','State','Native ID','Acquired UTC','Entity Key'],'Observed identities · not verified physical machines',{x:32,y:336,width:1856,height:416},cmdbGroups.Inventory);
table(cmdb,'cmdb-alias','CMDB Entity Alias',['Source Table','Source Key','Canonical Account Key','Canonical Device Key','Entity Key'],'Selected identity · native keys and canonical estate links',{x:32,y:776,width:1856,height:216},cmdbGroups.Inventory);
contextFilter(table(cmdb,'cmdb-links','CMDB Relationship Evidence',['From Name','Relationship Type','To Name','Resolution','Qualification','Source System','Source Table','Source Row Key','Acquired UTC','Link Key'],'Observed incoming and outgoing links · unresolved endpoints retained',{x:32,y:336,width:1856,height:656},cmdbGroups.Relationships),'CMDB Relationship Evidence','CMDB Link In Context');
table(cmdb,'cmdb-source-health','Native Source Health',['SourceName','Status','Coverage','SourceRowsLabel','CompletedDateTimeLabel','Evidence'],'Captured native source qualification · tenant scope',{x:32,y:384,width:1856,height:280},cmdbGroups.Evidence);
table(cmdb,'cmdb-findings','Native Entity Finding',['EntityType','EntityId','FindingType','Severity','Description','RecommendedAction','LinkStatus'],'Native reconciliation findings · tenant scope',{x:32,y:688,width:1856,height:304},cmdbGroups.Evidence);
text(cmdb,'cmdb-evidence-note','Source qualification is tenant-level and does not follow entity selection. Unknown is not healthy.',{x:32,y:336,width:1856,height:40},cmdbGroups.Evidence);
const facets=read(path.join(here,'../config/consolidation-report-facets.json.txt')).facets;
facets.push({...facets.find(f=>f.source==='LicenseAssignmentPath'),page:'CMDB',label:'Group license paths'});
const reportCoverage=[];
for(const page of pageList){
  const mine=facets.filter(f=>f.page===page.displayName);if(!mine.length)continue;
  const pageDir=path.join(pages,page.name);
  const oldGroups=page.name===cmdb?Object.values(cmdbGroups):fs.readdirSync(path.join(pageDir,'visuals')).map(v=>read(path.join(pageDir,'visuals',v,'visual.json'))).filter(v=>v.visualGroup).map(v=>v.name);
  if(!oldGroups.length)throw Error('Existing page lacks governed view groups: '+page.displayName);
  const groups=mine.map(f=>group(page.name,'Native · '+f.label));
  const allGroups=[...oldGroups,...groups];
  for(const entry of existingMarks.filter(x=>x.value.explorationState?.activeSection===page.name)){
    const bm=clone(entry.value);bm.options.targetVisualNames=[...new Set([...bm.options.targetVisualNames,...groups])];
    const states=bm.explorationState.sections[page.name].visualContainerGroups||={};
    for(const g of groups)states[g]={isHidden:true};put(entry.file,bm);
  }
  const markIds=mine.map((f,i)=>bookmark(page.name,'Native · '+f.label,groups[i],allGroups));
  bookmarkList.items.push({name:id(page.name+':native-tabs'),displayName:page.displayName+' native facets',children:markIds});
  if(page.name!==cmdb)button(page.name,page.name+':native-entry','Native detail',markIds[0],{x:32,y:1008,width:260,height:28,z:12500,tabOrder:2000});
  if(page.pageBinding?.type==='Drillthrough')button(page.name,page.name+':context-back','Back',null,{x:1640,y:1008,width:248,height:28,z:12500,tabOrder:2001},null,'Back');
  for(let i=0;i<mine.length;i++){
    const f=mine[i],g=groups[i],seed=page.name+':'+f.source;
    text(page.name,seed+':selected','Selected view: Native · '+f.label,{x:32,y:312,width:1856,height:24,z:12000},g);
    const rowCount=Math.ceil(mine.length/6),buttonWidth=Math.floor((1856-5*8)/6);
    mine.forEach((view,j)=>button(page.name,seed+':nav:'+j,view.label,markIds[j],{x:32+(j%6)*(buttonWidth+8),y:344+Math.floor(j/6)*48,width:buttonWidth,height:40,z:12000},g));
    const filterY=344+rowCount*48;
    slicer(page.name,seed+':filter',f.table,f.primary,'Native selection · '+f.label,{x:32,y:filterY,width:600,height:80,z:12000},g);
    entityContext(page.name,seed+':scope',{x:656,y:filterY,width:1232,height:80,z:12000},g);
    const y=filterY+96,hasMetrics=f.measures.length>0;
    const ordered=[f.primary,...f.columns.filter(c=>c!==f.primary&& !['TenantKey','OrganizationKey','EnvironmentKey','TenantId'].includes(c)),...f.columns.filter(c=>['TenantKey','OrganizationKey','EnvironmentKey','TenantId'].includes(c))];
    const detail=table(page.name,seed+':detail',f.table,ordered,f.label+' · complete governed fields',{x:32,y,width:1856,height:(hasMetrics?824:992)-y,z:12000},g);
    contextFilter(detail,f.table,'Native Context Rows · '+f.table);
    if(hasMetrics)table(page.name,seed+':metrics',f.table,f.measures,'Native metrics · qualified source population, not workforce totals',{x:32,y:848,width:1856,height:144,z:12000},g,'Measure');
    reportCoverage.push({sourceTable:f.source,table:f.table,page:page.displayName,bookmark:markIds[i],visual:detail.name,columns:f.columns,measures:f.measures});
  }
  if(page.name===cmdb){
    const main=[['Inventory',bookmark(cmdb,'Inventory',cmdbGroups.Inventory,allGroups)],['Relationships',bookmark(cmdb,'Relationships',cmdbGroups.Relationships,allGroups)],['Groups',markIds[0]],['Evidence',bookmark(cmdb,'Evidence',cmdbGroups.Evidence,allGroups)]];
    main.forEach(([name,bm],i)=>button(cmdb,'cmdb-main:'+name,name,bm,{x:32+i*466,y:264,width:458,height:48}));
    bookmarkList.items.push({name:id('cmdb-main'),displayName:'CMDB views',children:main.filter(x=>x[0]!=='Groups').map(x=>x[1])});
  }
}
text(cmdb,'cmdb-footer','Observed links only. No inferred transitive impact, effective permissions or physical-machine guarantee.',{x:32,y:1008,width:1200,height:28});
for(const [i,name]of ['User Explorer','Device Explorer','Data Trust'].entries()){
  const target=pageList.find(p=>p.displayName===name);button(cmdb,'cmdb-explorer:'+name,name,target.name,{x:32+i*280,y:1042,width:264,height:28},null,'PageNavigation');
}
const order=read(path.join(pages,'pages.json'));const index=order.pageOrder.indexOf(pageList.find(p=>p.displayName==='Operational Overview').name);order.pageOrder.splice(index+1,0,cmdb);
put(path.join(pages,'pages.json'),order);put(path.join(marks,'bookmarks.json'),bookmarkList);
const files=[...patch].map(([file,value])=>({file,value,before:fs.existsSync(file)?fs.readFileSync(file,'utf8'):null}));
return {pageCount:order.pageOrder.length,fileCount:files.length,coverage:reportCoverage,files};
}
if(typeof process!=='undefined'&&process.argv[1]&&fileURLToPath(import.meta.url)===process.argv[1]){
 const result=buildReport(process.argv[2]);
 const start=Number(process.argv[3]||0), count=Number(process.argv[4]||result.files.length);
 if(!Number.isInteger(start)||!Number.isInteger(count)||start<0||count<0)throw Error('Invalid patch slice');
 process.stdout.write(JSON.stringify({...result,coverage:process.argv[3]===undefined?result.coverage:[],files:result.files.slice(start,start+count)}));
}
