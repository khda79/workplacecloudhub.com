// Emits a guarded PBIR patch plan. No writes, model access, data read or refresh.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {fileURLToPath} from 'node:url';
export function buildPatch(rootPath) {
  if (!rootPath) throw Error('Explicit private Report directory required');
  const root=path.resolve(rootPath), def=path.join(root,'definition'), pages=path.join(def,'pages'), marks=path.join(def,'bookmarks');
  const read=p=>JSON.parse(fs.readFileSync(p,'utf8').replace(/^\uFEFF/,''));
  const id=s=>crypto.createHash('sha256').update('teams-spo-v1:'+s).digest('hex').slice(0,20);
  const lit=v=>({expr:{Literal:{Value:typeof v==='boolean'?String(v):typeof v==='number'?`${v}D`:`'${v.replaceAll("'","''")}'`}}});
  const field=(t,c,k='Column')=>({field:{[k]:{Expression:{SourceRef:{Entity:t}},Property:c}},queryRef:`${t}.${c}`,nativeQueryRef:c});
  const patch=new Map(), put=(p,v)=>patch.set(p,v);
  const template=n=>read(path.join(pages,'a821f7deb92bb3bc4c78/visuals',n,'visual.json'));
  function shell(n,seed,pos,group) {
    const v=template(n);v.name=id(seed);v.position={...v.position,...pos,z:13000,tabOrder:13000};
    delete v.parentGroupName;delete v.isHidden;delete v.filterConfig;
    if(group)v.parentGroupName=group;
    return v;
  }
  function save(p,v){put(path.join(pages,p,'visuals',v.name,'visual.json'),v);return v;}
  function text(p,seed,value,y,h,g){
    const v=shell('55651abd11a3bf8eb812',seed,{x:32,y,width:1856,height:h},g);
    v.visual.objects.general[0].properties.paragraphs=[{textRuns:[{value,textStyle:{fontFamily:'Segoe UI',fontSize:'14px',color:'#334E68'}}]}];return save(p,v);
  }
  function table(p,seed,t,cols,label,y,h,g){
    const v=shell('346ed8e16b6d4212dfdb',seed,{x:32,y,width:1856,height:h},g);
    v.visual.query={queryState:{Values:{projections:cols.map(c=>field(t,c))}}};
    delete v.visual.objects.columnWidth;delete v.visual.objects.general;delete v.visual.objects.columnFormatting;
    v.visual.objects.columnHeaders[0].properties.autoSizeColumnWidth=lit(true);
    v.visual.objects.columnHeaders[0].properties.columnAdjustment=lit('growToFit');
    v.visual.visualContainerObjects.title=[{properties:{show:lit(true),text:lit(label),fontSize:lit(15)}}];
    v.visual.visualContainerObjects.subTitle=[{properties:{show:lit(false)}}];
    v.visual.visualContainerObjects.general=[{properties:{altText:lit(label)}}];
    v.visual.visualContainerObjects.stylePreset=[{properties:{name:lit('None')}}];return save(p,v);
  }
  function cards(p,seed,items,g){
    const v=shell('524feb1a40c49d095418',seed,{x:32,y:376,width:1856,height:136},g);
    v.visual.query={queryState:{Data:{projections:items.map(([t,c,label])=>({...field(t,c,'Measure'),nativeQueryRef:label,displayName:label}))}}};
    v.visual.objects.value[0].properties.fontSize=lit(26);
    v.visual.objects.label[0].properties.fontSize=lit(12);
    delete v.visual.objects.label[0].properties.text;
    v.visual.objects.spacing=[{selector:{id:'default'},properties:{verticalSpacing:lit(2)}}];
    // Explicit 0 vertical content margins, VCO 8px, 1px border. Height 78 <=136.
    // Five callouts: >=350px after padding; value width <=220px at 26pt.
    v.visual.visualContainerObjects.general=[{properties:{altText:lit('Receipt-qualified Teams and SharePoint observed scope')}}];
    return save(p,v);
  }
  const bmList=read(path.join(marks,'bookmarks.json'));
  const bmFiles=fs.readdirSync(marks).filter(f=>f.endsWith('.bookmark.json')).map(f=>({path:path.join(marks,f),value:read(path.join(marks,f))}));
  const allPages=fs.readdirSync(pages,{withFileTypes:true}).filter(e=>e.isDirectory()).map(e=>read(path.join(pages,e.name,'page.json')));
  const targetNames=['Collaboration Adoption','Cloud Content & Storage','Service & Content Explorer','Data Trust'];
  const coverage=[];
  for(const title of targetNames){
    const page=allPages.find(p=>p.displayName===title);if(!page)throw Error('Missing existing page: '+title);
    const p=page.name,g=id(p+':group'),bm=id(p+':bookmark');
    if(fs.existsSync(path.join(pages,p,'visuals',g,'visual.json')))throw Error('Teams/SPO view already exists; inspect before reapplying');
    const old=fs.readdirSync(path.join(pages,p,'visuals')).filter(n=>fs.existsSync(path.join(pages,p,'visuals',n,'visual.json'))).map(n=>read(path.join(pages,p,'visuals',n,'visual.json')));
    const groups=old.filter(v=>v.visualGroup).map(v=>v.name);if(!groups.length)throw Error('Expected grouped page');
    const group=shell('0494a34823b0fb981f15',p+':group',{},null);group.visualGroup.displayName='Teams & SharePoint';group.isHidden=true;save(p,group);
    for(const item of bmFiles.filter(b=>b.value.explorationState?.activeSection===p)){
      const v=structuredClone(item.value);v.options.targetVisualNames=[...new Set([...(v.options.targetVisualNames||[]),g])];
      (v.explorationState.sections[p].visualContainerGroups||={})[g]={isHidden:true};put(item.path,v);
    }
    put(path.join(marks,bm+'.bookmark.json'),{$schema:'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/bookmark/2.1.0/schema.json',name:bm,displayName:title+' · Teams & SharePoint',
      options:{applyOnlyToTargetVisuals:true,targetVisualNames:[...groups,g],suppressActiveSection:true,suppressData:true},
      explorationState:{version:'1.0',activeSection:p,sections:{[p]:{visualContainers:{},visualContainerGroups:Object.fromEntries([...groups,g].map(n=>[n,{isHidden:n!==g}]))}}}});
    bmList.items.push({name:id(p+':bookmark-group'),displayName:title+' Teams & SharePoint view',children:[bm]});
    const button=shell('34c99d7e680872d936e7',p+':entry',{x:320,y:1008,width:280,height:28},null);
    for(const t of button.visual.objects.text)t.properties.text=lit('Teams & SharePoint');
    button.visual.visualContainerObjects.visualLink=[{properties:{show:lit(true),type:lit('Bookmark'),bookmark:lit(bm)}}];
    button.visual.visualContainerObjects.general=[{properties:{altText:lit('Open Teams and SharePoint observed scope')}}];save(p,button);
    text(p,p+':heading','Selected view: Teams & SharePoint · observed tenant scope',312,24,g);
    text(p,p+':scope','This view is tenant-level: country and service selections above do not filter it. Source scope excludes OneDrive. Unmatched is not invalid; no observed link is not proof of a standalone site.',340,32,g);
    const site='SharePoint Site Links',cap='SharePoint Tenant Capacity';
    cards(p,p+':metrics',title==='Cloud Content & Storage' ? [
      [site,'Observed SPO Storage (GB)','Observed site storage (GB)'],[site,'Teams Linked SPO Storage (GB)','Teams-linked site storage (GB)'],
      [cap,'SPO Tenant Capacity (GB)','Tenant capacity (GB)'],[cap,'SPO Tenant Utilization (%)','Tenant utilization'],[cap,'SPO Tenant Capacity State','Capacity qualification']
    ] : [
      [site,'Observed SPO Sites','Observed SPO sites'],[site,'SPO Sites With Teams Link','Sites with observed Teams link'],
      [site,'SPO Sites With Teams Link (%)','Linked / observed sites'],['Teams Site Links','Teams Links Outside SPO Scope','Links outside SPO scope']
    ],g);
    if(title==='Data Trust') {
      table(p,p+':proof','Collaboration Source Qualification',['Source','Receipt Version','Completed UTC','Source Rows','Link Coverage Qualified','Qualification','Receipt Run ID'],'Acquisition proof · receipt rows are not prepared-table totals',536,256,g);
      table(p,p+':capacity',cap,['Capacity State','Capacity Source','Storage Used GB','Capacity GB','Utilization','Collected UTC'],'Tenant capacity · missing is not zero; no site quota substitution',816,136,g);
    } else if(title==='Cloud Content & Storage') {
      table(p,p+':sites',site,['Site','Teams Link State','Storage Used GB','Link Count','Usage Report Date','Collection Scope','Site ID'],'Observed sites · storage counted once per site',536,256,g);
      table(p,p+':capacity',cap,['Capacity State','Capacity Source','Storage Used GB','Capacity GB','Utilization','Collected UTC'],'Tenant capacity · independent of site selection',816,136,g);
    } else if(title==='Collaboration Adoption') {
      table(p,p+':sites',site,['Site','Teams Link State','Link Count','Storage Used GB','Usage Report Date','Site URL','Site ID'],'Observed sites · link coverage is not usage or adoption',536,416,g);
    } else {
      table(p,p+':links','Teams Site Links',['Team','Channel','Link Type','Link Status','Resolution','Site URL','Site ID','Team ID','Channel ID','Collected UTC'],'Direct observed Team / channel / site links · no effective-access inference',536,416,g);
    }
    text(p,p+':qualification','A standard channel has no separate site. Percentages require complete observed link coverage and captures within 24 hours. Capacity remains Not measured without a qualified tenant quota. Sharing and hub status are not inferred.',964,32,g);
    coverage.push({page:title,group:g,bookmark:bm});
  }
  put(path.join(marks,'bookmarks.json'),bmList);
  return {coverage,files:[...patch].map(([file,value])=>({file,value,before:fs.existsSync(file)?fs.readFileSync(file,'utf8'):null}))};
}
if(process.argv[1]===fileURLToPath(import.meta.url))process.stdout.write(JSON.stringify(buildPatch(process.argv[2])));
