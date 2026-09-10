'use strict';
// Version: 3.8.0-beta.2 BETA. Offline regression contracts, not an M/DAX engine.
const fs=require('fs'),path=require('path'),assert=require('node:assert/strict');
const app=path.resolve(process.argv[2]||path.join(__dirname,'..'));
const model=JSON.parse(fs.readFileSync(path.join(app,'pbip/SmartWorkplaceDashboard.SemanticModel/model.bim')));
const expr=n=>model.model.expressions.find(e=>e.name===n)?.expression||'';
const measure=n=>model.model.tables.flatMap(t=>t.measures||[]).find(m=>m.name===n)?.expression||'';
const part=n=>model.model.tables.find(t=>t.name===n).partitions[0].source.expression;
let pass=0,fail=0;
function test(name,fn){try{fn();pass++;console.log('PASS '+name)}catch(e){fail++;console.log('FAIL '+name+': '+e.message.split('\n')[0])}}
test('unknown booleans excluded from disabled/noncompliant',()=>{for(const m of model.model.tables.flatMap(t=>t.measures||[]))assert(!/(?<!=)=FALSE\(\)/.test(m.expression),m.name)});
test('missing IDs do not count as one entity',()=>{for(const m of model.model.tables.flatMap(t=>t.measures||[]))assert(!/\bDISTINCTCOUNT\(/.test(m.expression),m.name)});
test('complement rates do not invent full success or usage from absent inputs',()=>{
 for(const name of ['Migration Success Rate','Exchange Storage Used Rate'])assert.match(measure(name),/IF\(ISBLANK\(_Total\)\|\|_Total<=0/);
 assert.match(measure('Exchange Storage Used Rate'),/ISBLANK\(_Free\)/);
});
test('unknown sign-in is not proven inactive',()=>assert(!measure('Enabled Users Inactive 90d').includes('ISBLANK(LastActivity)||')));
test('tenant scope is mandatory',()=>assert(expr('fnGetSourceFiles').includes('SmartInventoryConfiguration')));
test('current folder is nonrecursive',()=>assert(expr('fnGetSourceFiles').includes('Folder.Contents')));
test('ambiguous current CSV fails',()=>assert(expr('fnLoadSourceTable').includes('SmartInventorySourceAmbiguous')));
test('missing timestamp is not replaced with now',()=>assert(!expr('fnLoadSourceTable').includes('then DateTime.LocalNow()')));
test('missing source columns fail',()=>assert(expr('fnLoadSourceTable').includes('SmartInventorySchemaMismatch')));
test('row tenant mismatch fails',()=>assert(expr('fnLoadSourceTable').includes('SmartInventoryTenantMismatch')));
test('partial inventory fails',()=>assert(expr('fnLoadSourceTable').includes('SmartInventoryPartialSource')));
test('history row floor rounds up',()=>assert(expr('fnLoadSourceTable').includes('Number.RoundUp(Table.RowCount(CurrentTable)')));
test('fractional integer is not silently rounded',()=>assert(expr('fnToInt64').includes('Number.RoundTowardZero')));
test('nonfinite numeric values are unknown',()=>assert(expr('fnToNumber').includes('Number.IsNaN')));
test('oldest content timestamp determines freshness',()=>assert(expr('fnLoadSourceTable').includes('List.Min(UsableDates)')));
test('blank and future age return blank',()=>{for(const n of ['Users','Devices','Licenses','Exchange'])assert(measure(n+' Source Age Hours').includes('ISBLANK(Stamp)||Stamp>NOW()'))});
test('immutable Base64 remains case sensitive',()=>assert(!/Norm\(\[[AEX]_ImmutableId\]\)/.test(part('UserDetail'))));
test('conflicting canonical matches fail',()=>{for(const n of ['UserDetail','DeviceDetail'])assert(part(n).includes('SmartInventoryIdentityConflict'))});
test('ambiguous deduplication fails',()=>assert(expr('fnUnique').includes('SmartInventoryDuplicateKey')));
test('missing canonical keys fail',()=>assert(expr('fnRequireKey').includes('SmartInventoryIdentityMissing')));
test('future activity is unknown',()=>{for(const n of ['UserDetail','DeviceDetail'])assert(part(n).includes('[LastActivityDateTime]>DateTime.LocalNow()'))});
test('beta model annotation exists',()=>assert(model.model.annotations.some(a=>a.name==='SmartWorkplaceDashboard.Version'&&a.value==='3.8.0-beta.2')));
test('all page subtitles identify beta',()=>{const pages=path.join(app,'pbip/SmartWorkplaceDashboard.Report/definition/pages');for(const d of fs.readdirSync(pages,{withFileTypes:true}).filter(d=>d.isDirectory())){const v=JSON.parse(fs.readFileSync(path.join(pages,d.name,'visuals',d.name+'subtitle','visual.json')));assert(JSON.stringify(v).includes('BETA 3.8.0-beta.2'),d.name)}});
console.log(JSON.stringify({passed:pass,failed:fail,scope:'Offline generated-code regression contracts; M and DAX not executed'}));process.exitCode=fail?1:0;
