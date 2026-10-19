"""Synthetic integration-test inputs; no tenant data or network access."""
import argparse
import csv
import datetime as dt
import hashlib
import json
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET
import workplace_spo_migration as m


def fixture(root):
    raw, staged = root/'raw', root/'prepared'
    last = raw/'DATA-LAST'
    last.mkdir(parents=True)
    staged.mkdir()
    template = Path(__file__).resolve().parents[1]/'config/templates/SmartWorkplaceIntelligence-Migration-to-SPO.xlsx'
    workbook = raw/'SmartWorkplaceIntelligence-Migration-to-SPO.xlsx'
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    def set_cell(sheet, address, value):
        row_num = ''.join(c for c in address if c.isdigit())
        data = sheet.find(ns+'sheetData')
        row = next((r for r in data if r.get('r') == row_num), None)
        if row is None:
            row = ET.SubElement(data, ns+'row', {'r': row_num})
        cell = next((c for c in row if c.get('r') == address), None)
        if cell is None:
            cell = ET.SubElement(row, ns+'c', {'r': address})
        for child in list(cell):
            cell.remove(child)
        cell.set('t', 'inlineStr')
        ET.SubElement(ET.SubElement(cell, ns+'is'), ns+'t').text = str(value)
    with zipfile.ZipFile(template) as source, zipfile.ZipFile(workbook, 'w', zipfile.ZIP_DEFLATED) as dest:
        strings=[''.join(e.itertext()) for e in ET.fromstring(source.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in source.namelist() else []
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename in ['xl/worksheets/sheet2.xml','xl/worksheets/sheet3.xml','xl/worksheets/sheet4.xml']:
                sheet = ET.fromstring(data)
                if item.filename.endswith('sheet4.xml'):
                    for address, value in [('B4','sample-project'),('B5','sample-prod'),('B6','2026-10-19')]:
                        set_cell(sheet,address,value)
                else:
                    # Resolve headers from the actual worksheet; preserve all unrelated template content.
                    headers={}
                    for c in sheet.find(ns+'sheetData').find(ns+'row'):
                        val=''.join(c.find(ns+'is').itertext()) if c.get('t')=='inlineStr' else strings[int(c.findtext(ns+'v'))] if c.get('t')=='s' else c.findtext(ns+'v','')
                        headers[val]=''.join(ch for ch in c.get('r') if ch.isalpha())
                    values={'MigrationID':'MIG-000001','Disposition':'MIGRATE','ExecutionStatus':'DONE','ValidationStatus':'NOT_RECORDED',
                            'ActualDate':'2026-10-12','TargetURL':'/sites/sample','TargetReviewStatus':'CONFIRMED','SourceScope':'SELF'}
                    if item.filename.endswith('sheet2.xml'):
                        values.update(SourceURL='http://portal.example.invalid',SourceName='Sample portal',SourceObjectType='WEB')
                    else:
                        values.update(SourcePath=r'\\files.example.invalid\share\folder',FolderName='Sample folder')
                    for key,value in values.items():
                        set_cell(sheet,headers[key]+'2',value)
                data=ET.tostring(sheet,encoding='utf-8',xml_declaration=True)
            dest.writestr(item,data)
    # Match a native source identity, keeping keys syntactically independent of client values.
    from test_spo_migration import web
    native=web()
    def write(name, rows, fields=None, folder=last, delimiter=';'):
        fields=fields or list(rows[0])
        with (folder/name).open('w',encoding='utf-8-sig',newline='') as stream:
            writer=csv.DictWriter(stream,fieldnames=fields,delimiter=delimiter)
            writer.writeheader();writer.writerows(rows)
    collection={k:v for k,v in native.items() if k not in ['WebId','Title']}
    collection['StorageBytes']='1024'
    write('SharePoint_OnPrem_Webs.csv',[native])
    write('SharePoint_OnPrem_SiteCollections.csv',[collection])
    write('SharePoint_OnPrem_SiteAdministrators.csv',[],['TenantKey','RunId'])
    write('SharePoint_OnPrem_CollectionCoverage.csv',[dict(native,Status='Complete')])
    infra=[]
    for suffix in ['Farms','WebApplications','WebApplicationZones','ContentDatabases','Servers','ServiceApplications']:
        name='SharePoint_OnPrem_'+suffix+'.csv';write(name,[native]);infra.append(name)
    def receipt(producer,names):
        entries=[]
        for name in names:
            entries.append({'File':name,'Rows':len(m.read_csv(last/name,';')),'SHA256':m.sha(last/name)})
        document={'Owner':'SmartInventory-SourceReceipt','ContractVersion':'1.2','TenantKey':'sample-prod','RunId':'run-a',
                  'Producer':producer+'.ps1','Status':'Completed','CompletedAtUtc':'2026-10-19T10:00:00Z',
                  'ConsumerScopeQualified':False,'FullInventoryQualified':False,'ScopeQualification':'ConfiguredOutputsOnly',
                  'Files':entries,'RequiredFiles':names}
        (last/('SmartInventory_'+producer+'.current.json.txt')).write_text(json.dumps(document),encoding='utf-8')
    receipt('SmartM365-SharePoint-OnPrem-Content-Inventory',['SharePoint_OnPrem_'+n+'.csv' for n in ['Webs','SiteCollections','SiteAdministrators','CollectionCoverage']])
    receipt('SmartM365-SharePoint-OnPrem-Infrastructure-Inventory',infra)
    write('DimSharePointSite.csv',[{'TenantKey':'sample-prod','SiteUrl':'https://sample.sharepoint.com/sites/sample'}],folder=staged,delimiter=',')
    (root/'fixture.json.txt').write_text(json.dumps({'WorkbookSHA256':m.sha(workbook)}),encoding='utf-8')


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--root',type=Path,required=True)
    fixture(parser.parse_args().root)
