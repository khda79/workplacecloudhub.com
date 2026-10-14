// Emit neutral model definitions for the MCP model-authoring API. No model I/O.
import fs from 'node:fs';
import {fileURLToPath} from 'node:url';
const input=JSON.parse(fs.readFileSync(new URL('../config/consolidation-context-input.json.txt',import.meta.url),'utf8'));
const q=s=>`'${s.replaceAll("'","''")}'`, ref=(t,c)=>`${q(t)}[${c.replaceAll(']',']]')}]`;
const roots=[['DimUser','TenantUserKey'],['DimDevice','TenantDeviceKey'],['DimGroup','TenantGroupKey'],['ADUserSource','TenantADUserKey'],['ADComputerSource','TenantADComputerKey'],['ADGroupSource','TenantADGroupKey'],['ADDirectoryObject','TenantADObjectKey'],['DimIntuneManagedDevice','TenantIntuneDeviceKey'],['DimDetectedApplication','TenantApplicationKey'],['DimTeam','TenantTeamKey'],['DimSharePointSite','TenantSiteKey'],['DimLicenseSku','TenantSkuKey'],['DimLicenseServicePlan','TenantServicePlanKey'],['FactMailbox','TenantMailboxKey'],['FactAutopilotDevice','TenantAutopilotDeviceKey']];
function reaches(source,target,seen=new Set()){
 if(source===target)return true;if(seen.has(source))return false;seen.add(source);
 return input.relationships.filter(r=>r.toTable===source&&r.isActive&&r.crossFilteringBehavior==='OneDirection').some(r=>reaches(r.fromTable,target,seen));
}
function targetFilter(source,key,target){
 const t=input.tables.find(t=>t.name===input.targets[target]);
 // Direct native foreign key takes precedence over a longer relationship path.
 if(t.columns.includes(key))return ref(t.name,key);
 if(reaches(source,target))return ref(input.targets[source],key);
 return null;
}
function wrap(expression,table){
 const source=Object.keys(input.targets).find(k=>input.targets[k]===table);
 const candidates=roots.map(([s,k])=>({source:s,key:k,filter:targetFilter(s,k,source)})).filter(x=>x.filter);
 const prelude=`VAR HasEntityContext = ISFILTERED('CMDB Entity Evidence'[Selection]) || ISFILTERED('CMDB Entity Evidence'[Entity Key])\nVAR OneEntity = HASONEVALUE('CMDB Entity Evidence'[Entity Key])\n`;
 let vars='',branches=[];
 candidates.forEach((c,i)=>{vars+=`VAR NativeKeys${i} = CALCULATETABLE(VALUES('CMDB Entity Alias'[Source Key]), KEEPFILTERS('CMDB Entity Alias'[Source Table] = "${c.source}"))\n`;branches.push(`COUNTROWS(NativeKeys${i}) > 0, CALCULATE((${expression}), KEEPFILTERS(TREATAS(NativeKeys${i}, ${c.filter})))`);});
 const userFilter=targetFilter('DimUser','TenantUserKey',source);
 const deviceFilter=targetFilter('DimDevice','TenantDeviceKey',source);
 const inventory=[];
 if(userFilter){vars+=`VAR InventoryUsers = CALCULATETABLE(VALUES('Native User Evidence'[TenantUserKey]), REMOVEFILTERS('Native User Evidence'), TREATAS(VALUES('User Inventory Evidence'[User Source ID]), 'Native User Evidence'[SourceUserId]))\n`;inventory.push(`ISCROSSFILTERED('User Inventory Evidence'), CALCULATE((${expression}), KEEPFILTERS(TREATAS(InventoryUsers, ${userFilter})))`);}
 if(deviceFilter){vars+=`VAR InventoryDevices = CALCULATETABLE(VALUES('Native Managed Device Evidence'[TenantDeviceKey]), REMOVEFILTERS('Native Managed Device Evidence'), TREATAS(VALUES('Device Inventory Evidence'[Device Source ID]), 'Native Managed Device Evidence'[ManagedDeviceId]))\n`;inventory.push(`ISCROSSFILTERED('Device Inventory Evidence'), CALCULATE((${expression}), KEEPFILTERS(TREATAS(InventoryDevices, ${deviceFilter})))`);}
 const native=branches.length?`SWITCH(TRUE(), ${branches.join(',\n')}, BLANK())`:'BLANK()';
 const fallback=inventory.length?`SWITCH(TRUE(), ${inventory.join(',\n')}, (${expression}))`:`(${expression})`;
 return prelude+vars+`RETURN IF(HasEntityContext, IF(OneEntity, ${native}, BLANK()), ${fallback})`;
}
export function definitions(){
 const updates=input.measures.map(m=>({...m,expression:wrap(m.expression,m.tableName)}));
 const filters=input.tables.map(t=>({tableName:t.name,name:'Native Context Rows · '+t.name,expression:wrap(`COUNTROWS(${q(t.name)})`,t.name),isHidden:true,formatString:'#,0',description:'Exact native-key entity context or existing inventory context. No name matching or inferred transitive membership. Unrelated entity selections return blank.'}));
 return {updates,filters,policy:'Explicit TREATAS projection; all native relationships remain one-directional. A single synchronized native entity selection takes precedence. Otherwise existing qualified workforce/device scope applies where a native key/path exists.'};
}
if(typeof process!=='undefined'&&process.argv[1]&&fileURLToPath(import.meta.url)===process.argv[1]){
 const result=definitions(),kind=process.argv[2];
 process.stdout.write(JSON.stringify(kind?result[kind].slice(Number(process.argv[3]||0),Number(process.argv[3]||0)+Number(process.argv[4]||5)):result));
}
