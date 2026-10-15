// Static/offline checks. No Desktop, data access, refresh or publication.
import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
const root=process.argv[2];assert(root,'Explicit existing private Report directory required');
const read=p=>JSON.parse(fs.readFileSync(p,'utf8').replace(/^\uFEFF/,''));
const id=s=>crypto.createHash('sha256').update('teams-spo-v1:'+s).digest('hex').slice(0,20);
const def=path.join(root,'definition'),pages=path.join(def,'pages'),marks=path.join(def,'bookmarks');
const order=read(path.join(pages,'pages.json'));assert.equal(order.pageOrder.length,28);
const recipe=read(new URL('../config/teams-sharepoint-model-patch.json.txt',import.meta.url));
const source=read(new URL('../config/prepared-source-contract.json.txt',import.meta.url));
const output=read(new URL('../config/native-output-contract.json.txt',import.meta.url));
let checks=1;
const check=(ok,message)=>{assert(ok,message);checks++;};
for(const file of ['M365_Teams_Channels.csv','M365_SPO_Tenant.csv'])check(source.currentFiles.includes(file),'Missing capture input');
for(const table of recipe.tables){
  const contract=output.tables.find(t=>t.table===table.name);check(!!contract,'Missing output contract');
  check(JSON.stringify(table.columns.map(c=>c.name))===JSON.stringify(contract.columns.map(c=>c.name)),'Model and prepared schema differ');
  check(table.mExpression.includes('GetPreparedEvidenceFile('),'Ungoverned model source');
}
check(recipe.measureUpdates[0].expression.includes('[Service] <> "Teams"'),'Double-counted Teams storage');
check(recipe.measures.find(m=>m.name==='SPO Sites With Teams Link (%)').expression.startsWith('IF([Teams Site Coverage Qualified] = 1'),'Unqualified percentage');
check(recipe.relationships.every(r=>r.crossFilteringBehavior==='OneDirection'&&!r.relyOnReferentialIntegrity),'Unsafe relationship');
const bmList=read(path.join(marks,'bookmarks.json'));
for(const name of ['Collaboration Adoption','Cloud Content & Storage','Service & Content Explorer','CMDB','Data Trust']){
  const page=order.pageOrder.map(p=>read(path.join(pages,p,'page.json'))).find(p=>p.displayName===name);
  check(!!page,'Missing existing page');const g=id(page.name+':group'),bm=id(page.name+':bookmark');
  const files=fs.readdirSync(path.join(pages,page.name,'visuals')).filter(v=>fs.existsSync(path.join(pages,page.name,'visuals',v,'visual.json')));
  const visuals=files.map(v=>read(path.join(pages,page.name,'visuals',v,'visual.json')));
  check(visuals.find(v=>v.name===g)?.isHidden===true,'New view must start hidden');
  const bookmark=read(path.join(marks,bm+'.bookmark.json'));
  check(bookmark.options.suppressData===true,'Bookmark changes user filters');
  check(bmList.items.some(x=>x.children?.includes(bm)),'Unregistered bookmark');
  const selected=bookmark.explorationState.sections[page.name].visualContainerGroups;
  for(const v of visuals.filter(v=>v.visualGroup))check(selected[v.name]?.isHidden===(v.name!==g),'Overlapping view groups');
  for(const f of fs.readdirSync(marks).filter(f=>f.endsWith('.bookmark.json'))){
    const old=read(path.join(marks,f));if(old.name===bm||old.explorationState?.activeSection!==page.name)continue;
    check(old.explorationState.sections[page.name].visualContainerGroups[g]?.isHidden===true,'Existing bookmark leaves new view visible');
  }
  const children=visuals.filter(v=>v.parentGroupName===g);
  for(const v of children){
    const p=v.position;check(p.x>=0&&p.y>=312&&p.x+p.width<=1920&&p.y+p.height<=1000,'View outside reserved body');
    for(const role of Object.values(v.visual?.query?.queryState||{}))for(const p of role.projections||[]){
      const kind=p.field.Column?'Column':'Measure',f=p.field[kind],t=recipe.tables.find(t=>t.name===f.Expression.SourceRef.Entity);
      check(!!t,'Unknown new-view table');
      check(kind==='Column'?t.columns.some(c=>c.name===f.Property):recipe.measures.some(m=>m.tableName===t.name&&m.name===f.Property),'Unknown new-view binding');
    }
  }
  for(let i=0;i<children.length;i++)for(let j=i+1;j<children.length;j++){
    const a=children[i].position,b=children[j].position;
    check(!(a.x<b.x+b.width&&a.x+a.width>b.x&&a.y<b.y+b.height&&a.y+a.height>b.y),'Overlapping new visuals');
  }
}
console.log(`Teams/SPO static report checks passed: ${checks}. No DAX execution or Desktop rendering.`);
