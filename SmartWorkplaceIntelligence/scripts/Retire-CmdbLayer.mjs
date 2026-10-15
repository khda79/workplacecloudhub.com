// Guarded offline migration plans. Model writes belong to the modeling MCP.
// Does not read production CSVs, refresh, publish, or save a Desktop session.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import assert from 'node:assert/strict';
import {fileURLToPath} from 'node:url';

export const retiredTables = ['CMDB Entity Evidence','CMDB Entity Alias','CMDB Relationship Evidence','CMDB Relationship Endpoint'];
export const retiredFiles = ['CMDBEntityEvidence.csv','CMDBEntityAlias.csv','CMDBRelationshipEvidence.csv','CMDBRelationshipEndpoint.csv'];
export const modelDescription = 'Business-facing semantic model backed by governed SmartM365 domain evidence and observed histories. Generic CMDB projections and synthetic demonstration tables are retired. Offline validation does not imply production or Desktop qualification.';
export const contextDescription = 'Existing qualified workforce/device scope via exact native IDs; otherwise the domain selection. No name matching or inferred transitive membership.';
const text = v => Array.isArray(v) ? v.join('\n') : v;
const read = p => JSON.parse(fs.readFileSync(p,'utf8').replace(/^\uFEFF/,''));
const references = value => retiredTables.some(t => (JSON.stringify(value)||'').includes(t));

// Split only top-level function arguments, respecting DAX strings, identifiers,
// escaped quotes/brackets, nested calls, and table constructors. Fail closed.
export function argumentsOf(value) {
  let start=0, depth=0, quote=null; const args=[];
  for(let i=0;i<value.length;i++) {
    const c=value[i];
    if(quote){if(c===quote){if(value[i+1]===quote)i++;else quote=null;}continue;}
    if(c==='"'||c==="'"||c==='['){quote=c==='['?']':c;continue;}
    if(c==='('||c==='{')depth++;
    else if(c===')'||c==='}')depth--;
    else if(c===','&&depth===0){args.push(value.slice(start,i).trim());start=i+1;}
    assert(depth>=0,'Unbalanced DAX');
  }
  assert(depth===0&&!quote,'Unterminated DAX'); args.push(value.slice(start).trim());return args;
}

export function removeEntityContext(expression) {
  const prefix="VAR HasEntityContext = ISFILTERED('CMDB Entity Evidence'[Selection]) || ISFILTERED('CMDB Entity Evidence'[Entity Key])\nVAR OneEntity = HASONEVALUE('CMDB Entity Evidence'[Entity Key])\n";
  assert(expression.startsWith(prefix),'Unexpected entity wrapper; manual review required');
  const marker='\nRETURN IF(HasEntityContext,', split=expression.lastIndexOf(marker);
  assert(split>=0 && expression.endsWith(')'),'Unexpected entity RETURN');
  const args=argumentsOf(expression.slice(split+'\nRETURN IF('.length,-1));
  assert(args.length===3&&args[0]==='HasEntityContext'&&args[1].startsWith('IF(OneEntity,'),'Unexpected entity branches');
  const vars=expression.slice(prefix.length,split).split('\n');
  const retained=[];
  for(const v of vars){
    if(/^VAR NativeKeys\d+ = CALCULATETABLE\(VALUES\('CMDB Entity Alias'/.test(v))continue;
    assert(/^VAR Inventory(Users|Devices) = /.test(v)||!v.trim(),'Unrecognized wrapper variable');
    if(v.trim())retained.push(v);
  }
  const result=retained.length ? retained.join('\n')+'\nRETURN '+args[2] : args[2];
  assert(!references(result)&&!result.includes('NativeKeys'),'Retired dependency remains');
  return result;
}

export function modelPlan(bim) {
  const model=bim.model||bim, updates=[], removedMeasures=[];
  for(const t of model.tables)for(const m of t.measures||[]){
    if(retiredTables.includes(t.name)){removedMeasures.push({tableName:t.name,name:m.name});continue;}
    const expression=text(m.expression);
    if(references(expression))updates.push({tableName:t.name,name:m.name,expression:removeEntityContext(expression),...(m.name.startsWith('Native Context Rows · ')?{description:contextDescription}:{})});
  }
  assert(retiredTables.every(n=>model.tables.some(t=>t.name===n)),'Expected original graph tables');
  const relationships=(model.relationships||[]).filter(r=>retiredTables.includes(r.fromTable)||retiredTables.includes(r.toTable));
  const columns=model.tables.filter(t=>!retiredTables.includes(t.name)).flatMap(t=>(t.columns||[]).filter(c=>references(c.expression)).map(c=>{
    assert(c.name==='CMDB Entity Key'&&['Native Group Evidence','AD Group Evidence'].includes(t.name),'Unexpected calculated dependency');
    return {tableName:t.name,name:c.name};
  }));
  return {updates,removedMeasures,columns,relationships:relationships.map(r=>({name:r.name})),tables:retiredTables.map(name=>({name}))};
}

export function checkModel(before,after) {
  before=before.model||before;after=after.model||after;
  const plan=modelPlan(before), changed=new Map(plan.updates.map(m=>[m.tableName+'|'+m.name,m.expression]));
  assert.equal(after.tables.length,before.tables.length-retiredTables.length);
  for(const t of before.tables.filter(t=>!retiredTables.includes(t.name))){
    const a=after.tables.find(x=>x.name===t.name);assert(a,'Lost domain table '+t.name);
    assert.deepEqual(a.columns,t.columns.filter(c=>!plan.columns.some(x=>x.tableName===t.name&&x.name===c.name)),'Changed domain columns '+t.name);
    assert.deepEqual(a.partitions,t.partitions,'Changed domain partition '+t.name);
    assert.equal(a.measures?.length||0,t.measures?.length||0,'Lost domain measure');
    for(const m of t.measures||[]){const n=a.measures.find(x=>x.name===m.name);assert(n);
      assert.equal(text(n.expression),changed.get(t.name+'|'+m.name)||text(m.expression),'Business expression drift '+m.name);
      const description=plan.updates.find(x=>x.tableName===t.name&&x.name===m.name)?.description;
      assert.deepEqual({...n,expression:null},{...m,...(description?{description}:{}),expression:null},'Measure metadata drift '+m.name);
    }
  }
  assert.deepEqual(after.relationships,before.relationships.filter(r=>!plan.relationships.some(x=>x.name===r.name)),'Domain relationship drift');
  assert.deepEqual(after.expressions,before.expressions,'Reader or parameters changed');
  assert.equal(after.description,modelDescription,'Model scope description');
  assert(!references(after.tables),'Retired graph reference remains');
  return {tables:after.tables.length,measures:after.tables.reduce((n,t)=>n+(t.measures?.length||0),0),rewritten:plan.updates.length,structuralOnly:true};
}

if(process.argv[1]===fileURLToPath(import.meta.url)){
  const [mode,input,after]=process.argv.slice(2);
  if(mode==='model')console.log(JSON.stringify(modelPlan(read(input))));
  else if(mode==='check-model')console.log(JSON.stringify(checkModel(read(input),read(after))));
  else throw Error('Expected model <MCP BIM snapshot> or check-model <before> <after>');
}
