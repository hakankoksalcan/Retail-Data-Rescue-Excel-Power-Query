"""Patch presentation XML while preserving Excel's original Power Query package.
Usage: python scripts/prepare_workbook.py /path/to/Excel-authored.xlsx
The input is never overwritten. No openpyxl save is used (it can drop Mashup parts).
"""
from pathlib import Path
from datetime import datetime, timezone
import argparse, json, hashlib, posixpath
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as ET

ROOT = Path(__file__).resolve().parents[1]
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'s': S}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('input', type=Path)
    args = parser.parse_args()
    dest = ROOT / 'Retail_Data_Rescue.xlsx'
    if args.input.resolve() == dest.resolve():
        raise ValueError('Use the untouched Excel-authored input, not the output file')
    with ZipFile(args.input) as z:
        parts = {item.filename: z.read(item.filename) for item in z.infolist()}
    workbook = ET.fromstring(parts['xl/workbook.xml'])
    rels = ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
    targets = {r.get('Id'): posixpath.normpath(posixpath.join('xl', r.get('Target'))) for r in rels}
    sheet_parts = {s.get('name'): targets[s.get('{'+R+'}id')] for s in workbook.find('s:sheets', NS)}
    updated = {}
    def sheet(name):
        part = sheet_parts[name]
        if part not in updated: updated[part] = ET.fromstring(parts[part])
        return updated[part]
    def cell(name, address, text, keep_formula=False):
        node = sheet(name).find('.//s:c[@r="'+address+'"]', NS)
        if node is None: raise ValueError('Missing cell '+name+'!'+address)
        if keep_formula:
            node.set('t','str')
            val = node.find('s:v', NS)
            if val is None: val = ET.SubElement(node, '{'+S+'}v')
            val.text = text
        else:
            for child in list(node): node.remove(child)
            node.set('t','inlineStr')
            ET.SubElement(ET.SubElement(node,'{'+S+'}is'),'{'+S+'}t').text = text
    mode = 'Reference dashboard | Embedded audit queries refreshed in Excel 2021; full-model integration pending'
    cell('Settings','A6',r'C:\RetailDataRescue\raw\Online Retail.xlsx')
    cell('Settings','A11','Four audit queries are embedded. Full reporting M sources are in power_query/; see the setup guide.')
    cell('RefreshInfo','B2',mode)
    import openpyxl
    initial = openpyxl.load_workbook(args.input, data_only=True, read_only=True)
    reference_time = str(initial['RefreshInfo']['A2'].value)
    initial.close()
    cell('Dashboard','B16',mode+' | Reference built UTC: '+reference_time,keep_formula=True)
    # Keep formula evaluation consistent with the published cached text.
    f = sheet('Dashboard').find('.//s:c[@r="B16"]/s:f',NS)
    f.text = 'RefreshInfo!B2&" | Reference built UTC: "&RefreshInfo!A2'
    names = list(sheet_parts)
    new_names = ['Dashboard'] + [name for name in names if name != 'Dashboard']
    sheet_list = workbook.find('s:sheets', NS)
    elements = {node.get('name'): node for node in sheet_list}
    for node in list(sheet_list): sheet_list.remove(node)
    for name in new_names: sheet_list.append(elements[name])
    for node in workbook.findall('s:definedNames/s:definedName', NS):
        local = node.get('localSheetId')
        if local is not None: node.set('localSheetId', str(new_names.index(names[int(local)])))
    view = workbook.find('s:bookViews/s:workbookView', NS)
    view.set('activeTab',str(new_names.index('Dashboard')))
    view.set('firstSheet',str(new_names.index('Dashboard')))
    for name in names:
        for v in sheet(name).findall('s:sheetViews/s:sheetView',NS):
            v.set('tabSelected','1' if name=='Dashboard' else '0')
    calc = workbook.find('s:calcPr',NS)
    calc.set('calcMode','auto');calc.set('fullCalcOnLoad','1')
    # Remove a local Desktop path Excel saved in workbook presentation metadata.
    for node in workbook.xpath('//*[local-name()="absPath"]'):
        parent=node.getparent();parent.remove(node)
    updated['xl/workbook.xml']=workbook
    core=ET.fromstring(parts['docProps/core.xml'])
    for node in core:
        if ET.QName(node).localname=='description':
            node.text='Retail data quality case study. Four embedded Power Query audit queries; reference dashboard; full-model M sources supplied separately.'
        if ET.QName(node).localname=='modified':
            node.text=datetime.now(timezone.utc).isoformat(timespec='seconds').replace('+00:00','Z')
    updated['docProps/core.xml']=core
    for part,root in updated.items(): parts[part]=ET.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
    with ZipFile(dest,'w',ZIP_DEFLATED) as out:
        for name,data in parts.items():out.writestr(name,data)
    with ZipFile(args.input) as old, ZipFile(dest) as new:
        protected=[n for n in old.namelist() if n.startswith(('customXml/','xl/connections','xl/queryTables/','xl/tables/','xl/charts/'))]
        assert all(old.read(n)==new.read(n) for n in protected),'Excel query/table/chart parts changed'
        assert new.testzip() is None
    report={'input_sha256':hashlib.sha256(args.input.read_bytes()).hexdigest(),'output_sha256':hashlib.sha256(dest.read_bytes()).hexdigest(),'power_query_package_preserved':True,'protected_parts_unchanged':protected,'scope':'Presentation metadata, path normalization and dashboard activation only; no M runtime execution'}
    (ROOT/'validation'/'workbook-preservation.json').write_text(json.dumps(report,indent=2)+'\n')
    print('Preserved',len(protected),'query/table/chart parts; wrote',dest.name)

if __name__=='__main__':main()
