// Offline regression and fault-injection checks. No DAX execution or refresh.
import fs from 'node:fs';
import assert from 'node:assert/strict';
import {modelPlan,checkModel,removeEntityContext,argumentsOf,retiredFiles} from './Retire-CmdbLayer.mjs';
import {definitions} from './Build-ConsolidatedContext.mjs';
const read=p=>JSON.parse(fs.readFileSync(p,'utf8'));
const [beforePath,afterPath]=process.argv.slice(2);
assert(beforePath&&afterPath,'Explicit before/after MCP BIM snapshots required');
const before=read(beforePath),after=read(afterPath),plan=modelPlan(before);
assert.equal(plan.updates.length,235);assert.equal(plan.columns.length,2);
assert.equal(plan.removedMeasures.length,4);assert.equal(plan.relationships.length,5);
const result=checkModel(before,after);
const changed=structuredClone(after);changed.model.tables[0].columns[0].name='Unexpected rename';
assert.throws(()=>checkModel(before,changed),/Changed domain columns/);
const broken=structuredClone(after);broken.model.tables.find(t=>t.measures?.length).measures[0].expression='0';
assert.throws(()=>checkModel(before,broken),/Business expression drift/);
const lost=structuredClone(after);lost.model.tables.pop();assert.throws(()=>checkModel(before,lost));
assert.throws(()=>removeEntityContext('COUNTROWS(\'CMDB Entity Evidence\')'),/Unexpected/);
assert.deepEqual(argumentsOf('a, IF(x, "a,b", \'a,b\'[c]]d]), {1,2}'),['a','IF(x, "a,b", \'a,b\'[c]]d])','{1,2}']);
assert.throws(()=>argumentsOf('a, IF(x'),/Unterminated/);
const generated=definitions();
for(const m of [...generated.updates,...generated.filters]){
  const expected=plan.updates.find(x=>x.tableName===m.tableName&&x.name===m.name);
  assert(expected,'Unknown regenerated measure '+m.name);
  // Regeneration must use exactly the preserved fallback from the live baseline.
  assert.equal(m.expression,expected.expression,'Generator drift '+m.name);
}
const contract=read(new URL('../config/native-output-contract.json.txt',import.meta.url));
assert.equal(contract.tables.length,50);
assert(!contract.tables.some(t=>retiredFiles.includes(t.file)));
assert(!fs.existsSync(new URL('./workplace_graph.py',import.meta.url)));
assert(!fs.readFileSync(new URL('./workplace_native_build.py',import.meta.url),'utf8').includes('build_graph'));
console.log(JSON.stringify({...result,regeneratedMeasures:generated.updates.length+generated.filters.length,negativeControls:6,retiredFiles:4,productionQualification:false}));
