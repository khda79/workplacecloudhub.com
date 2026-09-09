'use strict';
// Version: 3.8.0-beta.1 BETA. Build tests operate only on disposable copies.
const fs=require('fs'),path=require('path'),os=require('os'),crypto=require('crypto'),cp=require('child_process'),assert=require('assert/strict');
const app=path.resolve(process.argv[2]||path.join(__dirname,'..'));
const temp=fs.mkdtempSync(path.join(os.tmpdir(),'smart-dashboard-beta-')),copy=path.join(temp,'app');
fs.cpSync(app,copy,{recursive:true});
let passed=0;
function build(env={}){return cp.spawnSync(process.execPath,[path.join(copy,'scripts/Build-SmartWorkplaceDashboard.js')],{encoding:'utf8',env:{...process.env,SMART_M365_DATA_ROOT:'',...env}})}
function hashes(dir){let result={};for(const e of fs.readdirSync(dir,{withFileTypes:true})){const p=path.join(dir,e.name);if(e.isDirectory())for(const [k,v]of Object.entries(hashes(p)))result[e.name+'/'+k]=v;else result[e.name]=crypto.createHash('sha256').update(fs.readFileSync(p)).digest('hex')}return result}
function test(name,fn){fn();passed++;console.log('PASS '+name)}
function mutate(file,edit,pattern){const p=path.join(copy,file),original=fs.readFileSync(p),obj=JSON.parse(original);edit(obj);fs.writeFileSync(p,JSON.stringify(obj));try{const r=build();assert.notEqual(r.status,0);assert.match(r.stderr,pattern)}finally{fs.writeFileSync(p,original)}}
try{
test('clean package builds with no data root',()=>assert.equal(build().status,0));
const before=hashes(path.join(copy,'pbip'));
test('rebuild is deterministic',()=>{assert.equal(build().status,0);assert.deepEqual(hashes(path.join(copy,'pbip')),before)});
test('stable version rejected',()=>mutate('version.json',j=>j.version='3.8.0',/beta version is required/));
test('path traversal source rejected',()=>mutate('source-selection.json',j=>j.includedFiles[0]='../other.csv',/Unsafe CSV name/));
test('duplicate schema header rejected',()=>mutate('source-schema.json',j=>{const v=Object.values(j.files)[0];v.columns.push(v.columns[0])},/Invalid header/));
test('missing explicit input root fails',()=>{const r=build({SMART_M365_DATA_ROOT:path.join(temp,'missing')});assert.notEqual(r.status,0);assert.match(r.stderr,/Missing selected CSV/)});
test('failed builds preserve generated outputs',()=>assert.deepEqual(hashes(path.join(copy,'pbip')),before));
console.log(JSON.stringify({passed,scope:'Real Node.js build execution on isolated copies'}));
}finally{const resolved=path.resolve(temp);assert(resolved.startsWith(path.resolve(os.tmpdir())+path.sep)&&path.basename(resolved).startsWith('smart-dashboard-beta-'));fs.rmSync(resolved,{recursive:true,force:true})}
