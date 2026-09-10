'use strict';
// Version: 3.8.0-beta.2 BETA. Generates in-memory M tests; no data-source calls.
const fs=require('fs'),path=require('path');
const app=path.resolve(process.argv[2]||path.join(__dirname,'..')),out=process.argv[3];
if(!out)throw Error('Usage: node New-SyntheticQuery.cjs <app> <output.pq>');
const model=JSON.parse(fs.readFileSync(path.join(app,'pbip/SmartWorkplaceDashboard.SemanticModel/model.bim')));
const expression=n=>model.model.expressions.find(e=>e.name===n)?.expression||'(input as table,keys as any) => input';
const funcs=model.model.expressions.filter(e=>e.name.startsWith('fn')&&!['fnGetSourceFiles','fnLoadSourceTable'].includes(e.name)).map(e=>e.name+' = '+e.expression);
if(!model.model.expressions.some(e=>e.name==='fnUnique'))funcs.push('fnUnique=(input,keys)=>Table.Distinct(input,keys)');
if(!model.model.expressions.some(e=>e.name==='fnRequireKey'))funcs.push('fnRequireKey=(input,key)=>Table.SelectRows(input,each Record.Field(_,key)<>null)');
const q=s=>'"'+String(s).replace(/"/g,'""').replace(/\n/g,'#(lf)')+'"';
const val=v=>v===null?'null':typeof v==='boolean'||typeof v==='number'?String(v):q(v);
const tests=[];const add=(name,e)=>tests.push('Check('+q(name)+',()=> '+e+')');
add('invalid SourceMode rejected','ErrorIs(()=>Adapter("Invalid","tenant-a","DATA-LAST"),"SmartInventoryConfiguration")');
add('blank tenant configuration rejected','ErrorIs(()=>Adapter("LocalFolder","","DATA-LAST"),"SmartInventoryConfiguration")');
add('unknown relative folder rejected','ErrorIs(()=>Adapter("LocalFolder","tenant-a","../Other"),"SmartInventoryConfiguration")');
add('local current uses direct enumeration','Adapter("LocalFolder","tenant-a","DATA-LAST"){0}[Name]="direct.csv"');
add('local history uses recursive enumeration','Adapter("LocalFolder","tenant-a","DATA-ALL"){0}[Name]="recursive.csv"');
add('SharePoint current excludes nested folders','Table.RowCount(Adapter("SharePoint","tenant-a","DATA-LAST"))=1');
const date='#datetime(2026,1,14,0,0,0)',old='#datetime(2026,1,7,0,0,0)';
const file=(csv,stamp=date)=>'File('+q(csv)+','+stamp+')';
const basic='TenantKey,Value\ntenant-a,1';
const load=(csv,cols='{"TenantKey","Value"}',types='{{"Value","int64"}}',stamp=date,hist='{}')=>'Load({'+file(csv,stamp)+'},'+hist+','+cols+','+types+')';
const err=(e,reason)=>'ErrorIs(()=> '+e+','+q(reason)+')';
add('logical missing remains null','fnToLogical(null)=null and fnToLogical("")=null and fnToLogical("garbage")=null');
add('logical false/true French and numeric','fnToLogical("non")=false and fnToLogical("oui")=true and fnToLogical("0")=false and fnToLogical("1")=true');
add('decimal comma and dot','fnToNumber("1,25")=1.25 and fnToNumber("1.25")=1.25');
add('invalid and nonfinite numeric unknown','fnToNumber("invalid")=null and fnToNumber("NaN")=null and fnToNumber("Infinity")=null');
add('fractional integer unknown','fnToInt64("1.5")=null and fnToInt64("-1.5")=null and fnToInt64("12")=12');
add('invalid date unknown','fnToDateTime("invalid")=null and fnToDateTime("")=null');
add('comma import',load(basic)+'{0}[Value]=1');
add('semicolon import',load('TenantKey;Value\ntenant-a;2')+'{0}[Value]=2');
add('quoted numeric comma import',load('TenantKey,Value\ntenant-a,"1,25"','{"TenantKey","Value"}','{{"Value","number"}}')+'{0}[Value]=1.25');
add('header-only valid empty export','Table.RowCount('+load('TenantKey,Value\n')+')=0');
add('missing CSV',err('Load({}, {}, {"TenantKey","Value"}, {})','SmartInventorySourceMissing'));
add('duplicate CSV',err('Load({'+file(basic)+','+file(basic)+'}, {}, {"TenantKey","Value"}, {})','SmartInventorySourceAmbiguous'));
add('missing source field',err(load('TenantKey\ntenant-a'),'SmartInventorySchemaMismatch'));
add('duplicate source header',err(load('TenantKey,TenantKey\ntenant-a,1'),'SmartInventorySchemaMismatch'));
add('empty file',err(load(''),'SmartInventorySchemaMismatch'));
add('cross tenant row',err(load('TenantKey,Value\ntenant-b,1'),'SmartInventoryTenantMismatch'));
add('blank tenant row',err(load('TenantKey,Value\n,1'),'SmartInventoryTenantMismatch'));
add('tenant normalization',load('TenantKey,Value\n TENANT-A ,3')+'{0}[Value]=3');
add('partial inventory',err(load('TenantKey,IsPartialInventory\ntenant-a,true','{"TenantKey","IsPartialInventory"}','{}'),'SmartInventoryPartialSource'));
add('unknown completeness',err(load('TenantKey,IsPartialInventory\ntenant-a,','{"TenantKey","IsPartialInventory"}','{}'),'SmartInventoryPartialSource'));
add('known complete inventory',load('TenantKey,IsPartialInventory\ntenant-a,false','{"TenantKey","IsPartialInventory"}','{}')+'{0}[TenantKey]="tenant-a"');
add('missing timestamp remains unknown',load(basic,undefined,undefined,'null')+'{0}[__SnapshotDateTime]=null');
add('future timestamp rejected',err(load(basic,undefined,undefined,'#datetime(2099,1,1,0,0,0)'),'SmartInventoryFutureTimestamp'));
add('oldest content date',load('TenantKey,ReportRefreshDate\ntenant-a,2026-01-02\ntenant-a,2026-01-01','{"TenantKey","ReportRefreshDate"}','{}')+'{0}[__SnapshotDateTime]=#datetime(2026,1,1,0,0,0)');
add('partly missing content date unknown',load('TenantKey,ReportRefreshDate\ntenant-a,2026-01-02\ntenant-a,','{"TenantKey","ReportRefreshDate"}','{}')+'{0}[__SnapshotDateTime]=null');
const hist='{HistoryFile('+q(basic)+','+old+')}';
add('weekly history retained','Table.RowCount('+load(basic,undefined,undefined,date,hist)+')=2');
add('history floor rounds up for three current rows','Table.RowCount('+load('TenantKey,Value\ntenant-a,1\ntenant-a,2\ntenant-a,3',undefined,undefined,date,hist)+')=3');
add('historical tenant mismatch rejected',err(load(basic,undefined,undefined,date,'{HistoryFile('+q('TenantKey,Value\ntenant-b,1')+','+old+')}'),'SmartInventoryTenantMismatch'));
add('historical duplicate timestamps rejected',err(load(basic,undefined,undefined,date,'{HistoryFile('+q(basic)+','+old+'),HistoryFile('+q(basic)+','+old+')}'),'SmartInventoryHistoryAmbiguous'));
add('identical duplicate rows collapse','Table.RowCount(fnUnique(#table({"Id","Value"},{{"a",1},{"a",1}}),{"Id"}))=1');
add('conflicting duplicate key rejected',err('fnUnique(#table({"Id","Value"},{{"a",1},{"a",2}}),{"Id"})','SmartInventoryDuplicateKey'));
add('missing canonical ID rejected',err('fnRequireKey(#table({"Id"},{{null}}),"Id")','SmartInventoryIdentityMissing'));
const selected=model.model.tables.filter(t=>!['UserDetail','DeviceDetail'].includes(t.name));
function detail(name,rows,condition){
 const defs=selected.map(t=>{const columns=t.columns.map(c=>c.name);const data=(rows[t.name]||[]).map(r=>'{'+columns.map(c=>val(c==='__IsCurrent'?true:r[c]??null)).join(',')+'}');return '#'+q(t.name)+'=#table({'+columns.map(q).join(',')+'},{'+data.join(',')+'})'});
 return '(let '+defs.join(',\n')+', Result=('+model.model.tables.find(t=>t.name===name).partitions[0].source.expression+') in '+condition+')';
}
add('empty UserDetail',detail('UserDetail',{},'Table.RowCount(Result)=0'));
add('empty DeviceDetail',detail('DeviceDetail',{},'Table.RowCount(Result)=0'));
add('unmatched users retained',detail('UserDetail',{'M365_Users_Active':[{'Object Id':'u-a','User principal name':'a@example.invalid',AccountEnabled:'true'}]},'Table.RowCount(Result)=1 and Result{0}[InEntra]=true and Result{0}[InAD]=false'));
add('Base64 immutable IDs remain distinct',detail('UserDetail',{'M365_Users_Active':[{'Object Id':'u-a',OnPremisesImmutableId:'AbCd=='},{'Object Id':'u-b',OnPremisesImmutableId:'abcd=='}],'AD_Users_AllDomains':[{ObjectGUID:'ad-a',ImmutableId_AD:'AbCd=='}]},'Table.RowCount(Result)=2 and Table.RowCount(Table.SelectRows(Result,each [InAD]=true and [EntraObjectId]="u-a"))=1'));
add('conflicting user identifiers rejected',err(detail('UserDetail',{'M365_Users_Active':[{'Object Id':'u-a',OnPremisesImmutableId:'AbCd==','User principal name':'a@example.invalid'},{'Object Id':'u-b','User principal name':'b@example.invalid'}],'AD_Users_AllDomains':[{ObjectGUID:'ad-a',ImmutableId_AD:'AbCd==',UserPrincipalName:'b@example.invalid'}]},'Result'),'SmartInventoryIdentityConflict'));
add('blank user canonical ID rejected',err(detail('UserDetail',{'M365_Users_Active':[{'User principal name':'a@example.invalid'}]},'Result'),'SmartInventoryIdentityMissing'));
add('unknown user activity is not healthy',detail('UserDetail',{'M365_Users_Active':[{'Object Id':'u-a'}]},'Result{0}[ActionSeverity]<>"Healthy"'));
add('unmatched device retained',detail('DeviceDetail',{'M365_Entra_Devices':[{ObjectId:'e-a',DeviceId:'d-a'}]},'Table.RowCount(Result)=1 and Result{0}[InEntra]=true'));
add('conflicting device identifiers rejected',err(detail('DeviceDetail',{'M365_Entra_Devices':[{ObjectId:'e-a',DeviceId:'d-a'},{ObjectId:'e-b',DeviceId:'d-b'}],'Intune_Devices_Inventory':[{'Device ID':'i-a','Entra ObjectId':'e-a','Azure AD Device ID':'d-b'}]},'Result'),'SmartInventoryIdentityConflict'));
add('future activity is not active',detail('UserDetail',{'M365_Users_Active':[{'Object Id':'u-a',LastSignInDateTime:'2099-01-01T00:00:00'}]},'Result{0}[ActivityBucket]="Unknown"'));
const sourceDefinitions=selected.map(t=>'#'+q(t.name)+'=('+t.partitions[0].source.expression+')').join(',\n');
add('all 90 generated source partitions evaluate on empty synthetic schemas','(let fnLoadSourceTable=(file,delimiter,columns,types,history)=>#table(List.Combine({columns,{"__SnapshotDate","__SnapshotDateTime","__SnapshotPeriod","__IsCurrent","__SourceFile","__SourceFolder"}}),{}),'+sourceDefinitions+', Counts={'+selected.map(t=>'Table.RowCount(#'+q(t.name)+')').join(',')+'} in List.Sum(Counts)=0)');
const code=`let
${funcs.join(',\n')},
Adapter=(mode as text,key as text,folder as text) as table => let
 SourceMode=mode,ExpectedTenantKey=key,DataRootPath="synthetic",SharePointSiteUrl="https://example.invalid/site",SharePointDataFolderUrl="https://example.invalid/site/data",HistoryMonths=24,HistoryMinRowRatio=0.5,
 MockFile=(name,folder)=>[Name=name,Content=Text.ToBinary("x"),Extension=".csv",#"Folder Path"=folder,#"Date modified"=#datetime(2026,1,1,0,0,0),Size=1],
 MockFolderContents=(p)=>Table.FromRecords({MockFile("direct.csv",p)}),MockFolderFiles=(p)=>Table.FromRecords({MockFile("recursive.csv",p)}),
 MockSharePointFiles=(url,options)=>Table.FromRecords({MockFile("direct.csv","https://example.invalid/site/data/DATA-LAST/"),MockFile("nested.csv","https://example.invalid/site/data/DATA-LAST/other/")}),
 Get=${expression('fnGetSourceFiles').replaceAll('Folder.Contents(','MockFolderContents(').replaceAll('Folder.Files(','MockFolderFiles(').replaceAll('SharePoint.Files(','MockSharePointFiles(')}
 in Get(folder),
File=(csv as text,stamp as nullable datetime) as record => [Name="test.csv",Content=Text.ToBinary(csv),Extension=".csv",#"Folder Path"="synthetic/DATA-LAST/",#"Date modified"=stamp,Size=Text.Length(csv)],
HistoryFile=(csv as text,stamp as datetime) as record => Record.TransformFields(File(csv,stamp),{{"Folder Path",each "synthetic/DATA-ALL/WeeklyHistory/"}}),
Load=(current as list,history as list,columns as list,types as list) as table => let
 SourceMode="LocalFolder",ExpectedTenantKey="tenant-a",HistoryMonths=24,HistoryMinRowRatio=0.5,
 fnGetSourceFiles=(folder as text) as table => Table.FromRecords(if folder="DATA-LAST" then current else history,type table [Name=text,Content=binary,Extension=text,#"Folder Path"=text,#"Date modified"=nullable datetime,Size=number]),
 fnLoadSourceTable=${expression('fnLoadSourceTable')}
 in fnLoadSourceTable("test.csv",",",columns,types,not List.IsEmpty(history)),
ErrorIs=(f as function,reason as text) as logical => let R=try Table.Buffer(f()) in R[HasError] and R[Error][Reason]=reason,
Check=(name as text,f as function) as record => let R=try f() in [Name=name,Passed=if R[HasError] then false else R[Value]=true,Error=if R[HasError] then R[Error][Message] else ""],
Tests={${tests.join(',\n')}}
in Table.FromRecords(Tests)`;
fs.writeFileSync(out,code);console.log('Generated '+tests.length+' in-memory M scenarios.');
