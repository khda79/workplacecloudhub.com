'use strict';
// Version: 3.8.0-beta.2 BETA. Build tests operate only on disposable copies.
const fs=require('fs'),path=require('path'),os=require('os'),crypto=require('crypto'),cp=require('child_process'),assert=require('assert/strict');
const app=path.resolve(process.argv[2]||path.join(__dirname,'..'));
const temp=fs.mkdtempSync(path.join(os.tmpdir(),'smart-dashboard-beta-')),copy=path.join(temp,'app');
function copyBuildInputs(source,target){fs.cpSync(source,target,{recursive:true,filter:p=>!path.relative(source,p).split(path.sep).some(part=>['private','.pbi','.git','local_memory.md'].includes(part.toLowerCase()))})}
copyBuildInputs(app,copy);
let passed=0;
function build(env={}){return cp.spawnSync(process.execPath,[path.join(copy,'scripts/Build-SmartWorkplaceDashboard.js')],{encoding:'utf8',env:{...process.env,SMART_M365_DATA_ROOT:'',...env}})}
function hashes(dir){let result={};for(const e of fs.readdirSync(dir,{withFileTypes:true})){const p=path.join(dir,e.name);if(e.isDirectory())for(const [k,v]of Object.entries(hashes(p)))result[e.name+'/'+k]=v;else result[e.name]=crypto.createHash('sha256').update(fs.readFileSync(p)).digest('hex')}return result}
function test(name,fn){fn();passed++;console.log('PASS '+name)}
function mutate(file,edit,pattern){const p=path.join(copy,file),original=fs.readFileSync(p),obj=JSON.parse(original);edit(obj);fs.writeFileSync(p,JSON.stringify(obj));try{const r=build();assert.notEqual(r.status,0);assert.match(r.stderr,pattern)}finally{fs.writeFileSync(p,original)}}
try{
test('build tests exclude local caches and private files',()=>{
 const probe=path.join(temp,'copy-probe'),output=path.join(temp,'copy-probe-output');
 for(const rel of ['public.txt','private/cache.abf','pbip/Model.SemanticModel/.pbi/cache.abf','LOCAL_MEMORY.md']){const p=path.join(probe,rel);fs.mkdirSync(path.dirname(p),{recursive:true});fs.writeFileSync(p,'synthetic marker only')}
 copyBuildInputs(probe,output);
 assert.equal(fs.readFileSync(path.join(output,'public.txt'),'utf8'),'synthetic marker only');
 for(const rel of ['private','pbip/Model.SemanticModel/.pbi','LOCAL_MEMORY.md'])assert(!fs.existsSync(path.join(output,rel)),rel);
});
test('clean package builds with no data root',()=>assert.equal(build().status,0));
test('license measures avoid the Rows variable rejected by Desktop',()=>{
 const model=JSON.parse(fs.readFileSync(path.join(copy,'pbip/SmartWorkplaceDashboard.SemanticModel/model.bim')));
 const measures=model.model.tables.flatMap(t=>t.measures||[]);
 for(const m of measures)assert(!/\bVAR\s+Rows\b/i.test(m.expression),m.name);
 const names=['M365 E3 Purchased','M365 E3 Consumed','M365 E5 Purchased','M365 E5 Consumed','M365 F3 Purchased','M365 F3 Consumed','SharePoint Estimated Capacity TB'];
 for(const name of names){const m=measures.find(m=>m.name===name);assert(m,name);assert.match(m.expression,/VAR _LicenseRows=FILTER\(/);assert(!/\bRows\b/.test(m.expression),name)}
});
test('all generated KPI cards use supported numeric formatting',()=>{
  const pages=path.join(copy,'pbip/SmartWorkplaceDashboard.Report/definition/pages');let cards=0;
  for(const page of fs.readdirSync(pages,{withFileTypes:true}).filter(e=>e.isDirectory())){
    const visuals=path.join(pages,page.name,'visuals');
    for(const id of fs.readdirSync(visuals)){
      const doc=JSON.parse(fs.readFileSync(path.join(visuals,id,'visual.json')));
      if(doc.visual?.visualType!=='cardVisual')continue;
      cards++;
      const value=doc.visual.objects.value.find(e=>e.selector?.id==='default').properties;
      assert.equal(value.labelDisplayUnits.expr.Literal.Value,'0D');
      assert.equal(value.labelPrecision.expr.Literal.Value,'1L');
      assert(!('displayUnits' in value));assert(!('decimalPlaces' in value));
      assert.equal(doc.visual.query.queryState.Data.projections.length,1);
    }
  }
  assert.equal(cards,161);
});
const before=hashes(path.join(copy,'pbip'));
test('all generated textboxes reserve their full height for text',()=>{
 const pages=path.join(copy,'pbip/SmartWorkplaceDashboard.Report/definition/pages');let count=0;
 for(const page of fs.readdirSync(pages,{withFileTypes:true}).filter(e=>e.isDirectory()))for(const id of fs.readdirSync(path.join(pages,page.name,'visuals'))){
  const visual=JSON.parse(fs.readFileSync(path.join(pages,page.name,'visuals',id,'visual.json')));
  if(visual.visual?.visualType!=='textbox')continue;count++;
  const padding=visual.visual.visualContainerObjects.padding[0].properties;
  for(const side of ['top','bottom','left','right'])assert.equal(padding[side].expr.Literal.Value,'0D',visual.name+' '+side);
  assert.equal(visual.visual.visualContainerObjects.visualHeader[0].properties.show.expr.Literal.Value,'false');
 }
 assert.equal(count,204);
});
test('validator accepts generator and Desktop metadata but rejects unexpected CSV fields',()=>{
 const modelPath=path.join(copy,'pbip/SmartWorkplaceDashboard.SemanticModel/model.bim'),original=fs.readFileSync(modelPath),model=JSON.parse(original);
 const fixture=path.join(temp,'validator-fixture'),last=path.join(fixture,'DATA-LAST');fs.mkdirSync(last,{recursive:true});
 for(const [name,schema] of Object.entries(JSON.parse(fs.readFileSync(path.join(copy,'source-schema.json'))).files))fs.writeFileSync(path.join(last,name),schema.columns.map(c=>'"'+c.replace(/"/g,'""')+'"').join(schema.delimiter||',')+'\r\n');
 function validate(){return cp.spawnSync('pwsh',['-NoProfile','-ExecutionPolicy','AllSigned','-File',path.join(copy,'scripts/Validate-SmartWorkplaceDashboard.ps1'),'-CheckSourceFiles','-SkipPbirValidation'],{encoding:'utf8',windowsHide:true,env:{...process.env,SMART_M365_DATA_ROOT:fixture}})}
 try{
  let r=validate();assert.equal(r.status,0,r.stderr);
  model.compatibilityLevel=1606;model.model.tables[0].columns.unshift({type:'rowNumber',name:'RowNumber-synthetic',dataType:'int64',isHidden:true});fs.writeFileSync(modelPath,JSON.stringify(model));
  r=validate();assert.equal(r.status,0,r.stderr);
  model.model.tables[0].columns[0]={name:'RowNumber-synthetic',dataType:'string',sourceColumn:'RowNumber-synthetic'};fs.writeFileSync(modelPath,JSON.stringify(model));
  r=validate();assert.notEqual(r.status,0);assert.match(r.stderr,/CSV\/model columns differ/);
 }finally{fs.writeFileSync(modelPath,original)}
});
test('rebuild is deterministic',()=>{assert.equal(build().status,0);assert.deepEqual(hashes(path.join(copy,'pbip')),before)});
test('stable version rejected',()=>mutate('version.json',j=>j.version='3.8.0',/beta version is required/));
test('path traversal source rejected',()=>mutate('source-selection.json',j=>j.includedFiles[0]='../other.csv',/Unsafe CSV name/));
test('duplicate schema header rejected',()=>mutate('source-schema.json',j=>{const v=Object.values(j.files)[0];v.columns.push(v.columns[0])},/Invalid header/));
test('missing explicit input root fails',()=>{const r=build({SMART_M365_DATA_ROOT:path.join(temp,'missing')});assert.notEqual(r.status,0);assert.match(r.stderr,/Missing selected CSV/)});
test('failed builds preserve generated outputs',()=>assert.deepEqual(hashes(path.join(copy,'pbip')),before));
console.log(JSON.stringify({passed,scope:'Real Node.js build execution on isolated copies'}));
}finally{const resolved=path.resolve(temp);assert(resolved.startsWith(path.resolve(os.tmpdir())+path.sep)&&path.basename(resolved).startsWith('smart-dashboard-beta-'));fs.rmSync(resolved,{recursive:true,force:true})}
