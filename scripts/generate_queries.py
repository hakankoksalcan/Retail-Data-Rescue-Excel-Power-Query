"""Insert the shared explicit code rules and create simple Excel query wrappers."""
from pathlib import Path
import json
from build_reference import RULES
ROOT=Path(__file__).resolve().parents[1]
folder=ROOT/'power_query'
p=folder/'fnRetailModel.pq'
template=p.read_text(encoding='utf-8')
replacement='#table(type table [StockCode=text,CodeCategory=text],{'+','.join(
    '{'+json.dumps(k)+','+json.dumps(v)+'}' for k,v in RULES.items())+'})'
if '__CODE_RULES__' in template:
 template=template.replace('__CODE_RULES__',replacement)
else:
 template='\n'.join('    CodeRules = '+replacement+',' if line.startswith('    CodeRules = ') else line for line in template.split('\n'))
p.write_text(template,encoding='utf-8')
(folder/'qModel.pq').write_text('let\n    Config = Excel.CurrentWorkbook(){[Name="tblSource"]}[Content]{0},\n    Result = fnRetailModel(Text.From(Config[SourcePath]), Logical.From(Config[ValidateBaseline]))\nin\n    Result\n',encoding='utf-8')
fields=['KPI','Monthly','Country','Product','ProductDictionary','Reconciliation','Bridge','Quality','ReviewRecords',
 'ReviewExamples','DuplicateScenario','Controls','BeforeAfter','SpecialOperations','Customers','Facts','DashboardData']
for name in fields:
 (folder/('q'+name+'.pq')).write_text('let\n    Result = qModel['+name+']\nin\n    Result\n',encoding='utf-8')
(ROOT/'docs'/'code-rules.json').write_text(json.dumps(RULES,indent=2),encoding='utf-8')
print('Available',len(list(folder.glob('*.pq'))),'query sources plus explicit code rules')
