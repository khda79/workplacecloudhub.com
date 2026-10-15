"""Observed entity/relationship projection of governed native facets.

No name matching, transitive membership, fabricated parent, truncation or API.
Links stream to disk; only entity/alias indexes are held in memory.
"""
import csv
import hashlib
import json
from contextlib import ExitStack
from pathlib import Path

ENTITY_COLUMNS = ['TenantKey','Entity Key','Entity Type','Source System','Native ID',
                  'Display Name','Selection','State','Country','Acquired UTC','Source Table']
LINK_COLUMNS = ['TenantKey','Link Key','From Entity Key','To Entity Key','From Name',
                'To Name','Relationship Type','Source System','Resolution','Qualification',
                'Source Table','Source Row Key','Acquired UTC']
ENDPOINT_COLUMNS = ['TenantKey','Link Key','Entity Key']
ALIAS_COLUMNS = ['TenantKey','Entity Key','Source Table','Source Key','Canonical Account Key','Canonical Device Key']


def rows(root, name):
    with (root / (name + '.csv')).open(encoding='utf-8-sig', newline='') as stream:
        yield from csv.DictReader(stream)


def stable(*parts):
    return hashlib.sha256(json.dumps(parts, ensure_ascii=False, separators=(',', ':')).encode()).hexdigest()


def build_graph(root, tenant):
    root = Path(root)
    entities, aliases, native = {}, {}, {}
    alias_rows = []

    def entity(table, row, typ, system, id_field, label, source_key, state='', country=''):
        identifier = row.get(id_field, '').strip()
        if not identifier:
            raise ValueError('Observed entity lacks native identity: ' + table)
        key = stable(tenant, system, typ, identifier.lower())
        title = row.get(label, '') or identifier
        item = {'TenantKey':tenant,'Entity Key':key,'Entity Type':typ,'Source System':system,
                'Native ID':identifier,'Display Name':title,'Selection':title+' | '+typ+' | '+identifier,
                'State':row.get(state,'') or 'Unknown','Country':row.get(country,'') or 'Unknown',
                'Acquired UTC':row.get('SourceCollectedDateTime',''),'Source Table':table}
        if key in entities and entities[key] != item:
            raise ValueError('Conflicting observed entity identity: ' + table)
        entities[key] = item
        native[(system, typ, identifier.lower())] = key
        alias(table, row.get(source_key,''), key)
        return key

    def alias(table, value, key, account='', device=''):
        if not value:
            return
        identity = (table, value)
        if identity in aliases:
            if aliases[identity] != key:
                raise ValueError('Ambiguous entity alias: ' + table)
            return
        aliases[identity] = key
        alias_rows.append(dict(zip(ALIAS_COLUMNS,[tenant,key,table,value,account,device])))

    specs = [
        ('DimUser','User','Entra','SourceUserId','DisplayName','TenantUserKey','AccountStatusLabel','CountryLabel'),
        ('DimDevice','Device','Entra / Intune','CmdbDeviceId','DeviceName','TenantDeviceKey','ManagementStateLabel','CountryLabel'),
        ('DimGroup','Group','Entra','SourceGroupId','DisplayName','TenantGroupKey','MemberStatus',''),
        ('DimTeam','Team','Teams','TeamId','DisplayName','TenantTeamKey','ActivityState',''),
        ('DimSharePointSite','Site','SharePoint Online','SiteId','SiteName','TenantSiteKey','ActivityState',''),
        ('DimDetectedApplication','Application version','Intune discovered','AppId','DisplayName','TenantApplicationKey','RelationshipCoverageStatus',''),
        ('DimLicenseSku','License SKU','Microsoft 365','SkuId','SkuPartNumber','TenantSkuKey','',''),
        ('DimLicenseServicePlan','Service plan','Microsoft 365','TenantServicePlanKey','ServicePlanName','TenantServicePlanKey','ProvisioningStatus',''),
        ('FactMailbox','Mailbox','Exchange Online','CmdbMailboxId','DisplayName','TenantMailboxKey','LinkStatusLabel',''),
        ('FactAutopilotDevice','Autopilot registration','Intune','AutopilotDeviceId','DisplayName','TenantAutopilotDeviceKey','EnrollmentState',''),
        ('DimVerifiedDomain','Domain','Entra','DomainId','DomainId','TenantDomainKey','AvailabilityStatus',''),
        ('ADDomainSource','Domain','Active Directory','DNSRoot','DNSRoot','TenantADDomainKey','','')]
    for table, typ, system, identifier, label, source_key, state, country in specs:
        for row in rows(root, table):
            source_system = ('Exchange On-premises' if row['SourceSystem']=='local_mailboxes' else 'Exchange Online') if table=='FactMailbox' else system
            entity(table,row,typ,source_system,identifier,label,source_key,state,country)
    for row in rows(root,'ADDirectoryObject'):
        entity('ADDirectoryObject',row,'AD '+row['ObjectClass'],'Active Directory','ObjectGUID','DistinguishedName','TenantADObjectKey')
    for table, field, typ, label in [('ADUserSource','TenantADUserKey','AD user','DisplayName'),
                                   ('ADComputerSource','TenantADComputerKey','AD computer','Name'),
                                   ('ADGroupSource','TenantADGroupKey','AD group','Name')]:
        for row in rows(root,table):
            key = native.get(('Active Directory',typ,row['ObjectGUID'].lower()))
            if key is None:
                key=entity(table,row,typ,'Active Directory','ObjectGUID',label,field)
            else:
                alias(table,row[field],key)
                if row.get(label):
                    entities[key]['Display Name']=row[label]
                    entities[key]['Selection']=row[label]+' | '+typ+' | '+row['ObjectGUID']
    for row in rows(root,'DimIntuneManagedDevice'):
        key=aliases.get(('DimDevice',row['TenantDeviceKey']))
        if not key:
            raise ValueError('Managed-device bridge has no observed parent')
        alias('IntuneManagedDeviceId',row['ManagedDeviceId'].lower(),key)
        alias('DimIntuneManagedDevice',row['TenantIntuneDeviceKey'],key)
    # Direct observed Teams site relationships; never infer from display names.
    site_links = list(rows(root, 'TeamsSiteLinkEvidence')) if (root/'TeamsSiteLinkEvidence.csv').exists() else []
    for row in site_links:
        if row['Channel ID']:
            channel = dict(row, NativeChannelId=row['Team ID'].lower()+'|'+row['Channel ID'],
                           SourceCollectedDateTime=row['Collected UTC'])
            entity('TeamsSiteLinkEvidence',channel,'Channel','Teams','NativeChannelId','Channel','Link Key')
    for (system, typ, identifier), entity_key in native.items():
        if (system, typ) in [('Teams','Team'), ('SharePoint Online','Site')]:
            alias('TeamsSiteNative:'+typ, identifier, entity_key)
    # Exact canonical estate identifiers only. The estate's strong-link decision
    # is reused, not reimplemented from names, UPN or SMTP.
    for filename, field, mappings in [
        ('AllAccountsEvidence','Account Key', [('Entra Object ID','Entra','User'),('AD Object GUID','Active Directory','AD user')]),
        ('AllDevicesEvidence','Device Key',[('AD Object GUID','Active Directory','AD computer'),('Intune Device ID','IntuneManagedDeviceId','')])]:
        if not (root/(filename+'.csv')).exists():
            continue  # Native-only unit tests do not create estate outputs.
        for row in rows(root,filename):
            if row.get('TenantKey',tenant) != tenant:
                raise ValueError('Canonical estate tenant mismatch')
            for column, system, typ in mappings:
                identifier=row.get(column,'').strip().lower()
                key=aliases.get((system,identifier)) if not typ else native.get((system,typ,identifier))
                if key:
                    alias_rows.append(dict(zip(ALIAS_COLUMNS,[tenant,key,filename,row[field],row[field] if field=='Account Key' else '',row[field] if field=='Device Key' else ''])))
    counts={}
    with ExitStack() as stack:
        def writer(name, columns):
            stream=stack.enter_context((root/(name+'.csv')).open('x',encoding='utf-8',newline=''))
            result=csv.DictWriter(stream,fieldnames=columns); result.writeheader(); counts[name]=0
            return result
        ew=writer('CMDBEntityEvidence',ENTITY_COLUMNS)
        lw=writer('CMDBRelationshipEvidence',LINK_COLUMNS)
        bw=writer('CMDBRelationshipEndpoint',ENDPOINT_COLUMNS)
        aw=writer('CMDBEntityAlias',ALIAS_COLUMNS)
        for item in entities.values(): ew.writerow(item); counts['CMDBEntityEvidence']+=1
        for item in alias_rows: aw.writerow(item); counts['CMDBEntityAlias']+=1

        def link(table, row, row_key, from_table, from_key, to_table, to_key, kind, system, qualification='', raw_from='', raw_to=''):
            fk=aliases.get((from_table,from_key),''); tk=aliases.get((to_table,to_key),'')
            key=stable(tenant,table,row_key,kind)
            resolution='Resolved' if fk and tk else 'Unresolved endpoint'
            lw.writerow(dict(zip(LINK_COLUMNS,[tenant,key,fk,tk,
                entities[fk]['Display Name'] if fk else raw_from or from_key or 'Unknown',
                entities[tk]['Display Name'] if tk else raw_to or to_key or 'Unknown',
                kind,system,resolution,qualification or 'Observed native link; not transitive or effective access',
                table,row_key,row.get('SourceCollectedDateTime','')])))
            counts['CMDBRelationshipEvidence']+=1
            for endpoint in sorted({fk,tk}-{''}):
                bw.writerow(dict(zip(ENDPOINT_COLUMNS,[tenant,key,endpoint])));counts['CMDBRelationshipEndpoint']+=1

        for r in rows(root,'ADMembership'):
            link('ADMembership',r,r['TenantADMembershipKey'],'ADGroupSource',r['TenantADGroupKey'],'ADDirectoryObject',r['TenantADObjectKey'],r['MembershipKind'],'Active Directory',r['ResolutionStatus'],r['GroupSID'],r['MemberDistinguishedName'])
        for r in rows(root,'EntraGroupMembership'):
            target, field=next(((t,f) for t,f in [('DimUser','TenantUserKey'),('DimDevice','TenantDeviceKey'),('DimGroup','NestedTenantGroupKey')] if r.get(f)),('DimUser','TenantUserKey'))
            link('EntraGroupMembership',r,r['TenantMembershipKey'],'DimGroup',r['TenantGroupKey'],target,r.get(field,''),r['MembershipKind'],'Entra',r['LinkStatus'],raw_to=r['MemberId'])
        for r in rows(root,'FactUserDeviceRelationship'):
            link('FactUserDeviceRelationship',r,r['TenantRelationshipKey'],'DimUser',r['TenantUserKey'],'DimDevice',r['TenantDeviceKey'],r['RelationshipType'],r['SourceSystem'],'Reported user-device association; not console activity',r['Account'],r['Device'])
        for r in rows(root,'FactTeamMember'):
            link('FactTeamMember',r,r['TenantTeamMemberKey'],'DimTeam',r['TenantTeamKey'],'DimUser',r['TenantUserKey'],'Team '+r['Role'],'Teams',raw_to=r['UserId'])
        for r in site_links:
            r['SourceCollectedDateTime'] = r['Collected UTC']
            if r['Channel ID']:
                link('TeamsSiteLinkEvidence',r,r['Link Key'],'TeamsSiteNative:Team',r['Team ID'].lower(),
                     'TeamsSiteLinkEvidence',r['Link Key'],'Team channel','Teams',r['Link Type'],raw_from=r['Team ID'])
            if r['Link Type'] != 'standard':
                link('TeamsSiteLinkEvidence',r,r['Link Key'],
                     'TeamsSiteLinkEvidence' if r['Channel ID'] else 'TeamsSiteNative:Team',
                     r['Link Key'] if r['Channel ID'] else r['Team ID'].lower(),
                     'TeamsSiteNative:Site',r['Site ID'].lower() if r['Link Status']=='Collected' else '',
                     'Channel site' if r['Channel ID'] else 'Team primary site','Teams / SharePoint Online',
                     r['Link Status']+'; '+r['Resolution']+'; no effective access inference',raw_to=r['Site ID'])
        for r in rows(root,'FactDeviceApplication'):
            link('FactDeviceApplication',r,r['TenantDeviceApplicationKey'],'IntuneManagedDeviceId',r['ManagedDeviceId'].lower(),'DimDetectedApplication',r['TenantApplicationKey'],'Discovered application version','Intune',r['ApplicationLinkStatus']+'; '+r['DeviceLinkStatus'],r['ManagedDeviceId'],r['AppId'])
        for r in rows(root,'FactUserLicense'):
            link('FactUserLicense',r,stable(r['SourceUserId'],r['TenantSkuKey']),'DimUser',r['TenantUserKey'],'DimLicenseSku',r['TenantSkuKey'],'Reported SKU assignment','Microsoft 365',r['UserLinkStatus'],r['SourceUserId'],r['SkuId'])
        for r in rows(root,'LicenseAssignmentPath'):
            link('LicenseAssignmentPath',r,r['TenantAssignmentPathKey'],'DimUser',r['TenantUserKey'],'DimLicenseSku',r['TenantSkuKey'],'License '+r['AssignmentRoute'],'Microsoft 365',r['AssignmentState']+'; '+r['AssignmentError'],r['SourceUserId'],r['SkuId'])
            if r['AssignedByGroupId']:
                link('LicenseAssignmentPath',r,r['TenantAssignmentPathKey'],'DimGroup',r['TenantGroupKey'],'DimUser',r['TenantUserKey'],'Observed group license path','Microsoft 365',r['GroupLinkStatus']+'; '+r['UserLinkStatus'],r['AssignedByGroupId'],r['SourceUserId'])
        for table, field, target, target_field in [('ADUserSource','TenantADUserKey','DimUser','TenantUserKey'),('ADComputerSource','TenantADComputerKey','DimDevice','TenantDeviceKey'),('ADGroupSource','TenantADGroupKey','DimGroup','TenantGroupKey')]:
            for r in rows(root,table):
                if r[target_field]:
                    link(table,r,r[field],table,r[field],target,r[target_field],'Qualified native identity link','AD / Entra',r['CloudMatchStatus'])
        for r in rows(root,'FactMailbox'):
            system='Exchange On-premises' if r['SourceSystem']=='local_mailboxes' else 'Exchange Online'
            link('FactMailbox',r,r['TenantMailboxKey'],'FactMailbox',r['TenantMailboxKey'],'DimUser',r['TenantUserKey'],'Mailbox account',system,r['LinkStatusLabel'],raw_to=r['CmdbUserId'])
        for r in rows(root,'DimLicenseServicePlan'):
            link('DimLicenseServicePlan',r,r['TenantServicePlanKey'],'DimLicenseSku',r['TenantSkuKey'],'DimLicenseServicePlan',r['TenantServicePlanKey'],'SKU service plan','Microsoft 365',r['ProvisioningStatus'],r['SkuId'],r['ServicePlanId'])
        for r in rows(root,'FactUserServicePlan'):
            link('FactUserServicePlan',r,stable(r['SourceUserId'],r['TenantServicePlanKey']),'DimUser',r['TenantUserKey'],'DimLicenseServicePlan',r['TenantServicePlanKey'],'Reported user service plan','Microsoft 365','Enabled='+r['IsEnabled']+'; '+r['AssignmentState']+'; '+r['UserLinkStatus'],r['SourceUserId'],r['ServicePlanId'])
        for r in rows(root,'FactAutopilotDevice'):
            if r['ManagedDeviceId']:
                link('FactAutopilotDevice',r,r['TenantAutopilotDeviceKey'],'FactAutopilotDevice',r['TenantAutopilotDeviceKey'],'IntuneManagedDeviceId',r['ManagedDeviceId'].lower(),'Autopilot managed identity','Intune','Observed managed device identifier; no serial-number matching',raw_to=r['ManagedDeviceId'])
        for r in rows(root,'ADUserSource'):
            if r['ManagerDistinguishedName']:
                link('ADUserSource',r,r['TenantADUserKey'],'ADUserSource',r['TenantADUserKey'],'ADDirectoryObject',r['ManagerADObjectKey'],'Directory manager','Active Directory','Observed manager attribute; not inferred organizational authority',raw_to=r['ManagerDistinguishedName'])
    return counts
