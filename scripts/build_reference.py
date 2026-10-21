"""Independent Python reference, not an Excel/Power Query execution.
Run from the delivery folder: python3 scripts/build_reference.py
"""
from pathlib import Path
from datetime import datetime, timezone
from decimal import Decimal
import argparse, hashlib, json, re
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
EXPECTED_HASH = '43465a06f2ccf7c8b5bd2892bc7defb52f97487934fe93b16ae4c3936424676d'
RAW_FIELDS = ['InvoiceNo','StockCode','Description','Quantity','InvoiceDate','UnitPrice','CustomerID','Country']
RULES = {
    'POST':'Logistics', 'C2':'Logistics', 'DOT':'Logistics', 'D':'Discount',
    'M':'Manual', 'm':'Manual', 'BANK CHARGES':'BankFee', 'S':'Sample',
    'AMAZONFEE':'PlatformFee', 'B':'AccountingAdjustment', 'CRUK':'Commission',
    'gift_0001_10':'Voucher', 'gift_0001_20':'Voucher', 'gift_0001_30':'Voucher',
    'gift_0001_40':'Voucher', 'gift_0001_50':'Voucher',
    'PADS':'MerchandiseException', 'DCGS0076':'MerchandiseException',
    'DCGS0003':'MerchandiseException', 'DCGS0070':'MerchandiseException',
    'DCGS0055':'MerchandiseException', 'DCGS0072':'MerchandiseException',
    'DCGS0074':'MerchandiseException', 'DCGS0069':'MerchandiseException',
    'DCGS0057':'MerchandiseException', 'DCGSSBOY':'MerchandiseException',
    'DCGSSGIRL':'MerchandiseException', 'DCGS0004':'MerchandiseException',
    'DCGS0073':'MerchandiseException', 'DCGS0071':'MerchandiseException',
    'DCGS0068':'MerchandiseException', 'DCGS0067':'MerchandiseException',
    'DCGS0066P':'MerchandiseException',
}
SPECIAL = set(RULES) - {k for k,v in RULES.items() if v=='MerchandiseException'}
FLAG_DEFS = {
 'HasMissingCustomer':('Missing customer ID','review','Customer ID null; sale remains eligible'),
 'HasMissingDescription':('Missing description','review','Description null or whitespace-only text'),
 'HasNonTextDescription':('Non-text description','review','Nonempty description is not text'),
 'HasOuterWhitespace':('Description outer whitespace','format','Raw text differs from trimmed text'),
 'IsExactDuplicateMember':('Exact duplicate group member','review','All eight raw fields equal; all group members'),
 'IsDuplicateBeyondFirst':('Exact duplicate beyond first','scenario','SourceRow greater than group minimum'),
 'IsDescriptionVariantDuplicate':('Description-only duplicate candidate','review','Seven other raw fields equal; descriptions differ'),
 'HasMultipleDescriptions':('Code with multiple descriptions','review','Code has >1 nonblank raw text description'),
 'IsNonCNegative':('Negative quantity without C','review','Quantity < 0 and no C invoice prefix'),
 'IsZeroPrice':('Zero price','review','UnitPrice == 0'),
 'IsNegativePrice':('Negative price','review','UnitPrice < 0'),
 'IsLargeAmount':('Large absolute line amount','review','abs(LineAmount) >= GBP 10000'),
 'IsUnmappedCode':('Unmapped nonstandard code','review','Neither explicit code rule nor common merchandise format'),
 'IsInvoicePatternException':('Invoice pattern outside numeric/C','context','Neither six digits nor C + six digits'),
 'IsCancellationCode':('C cancellation prefix','context','InvoiceNo starts with uppercase C'),
}
REVIEW_FLAGS = [k for k,(label,kind,definition) in FLAG_DEFS.items() if kind=='review']

def jsonable(v):
 if isinstance(v, (datetime,pd.Timestamp)): return v.isoformat()
 if isinstance(v,dict): return {str(k):jsonable(x) for k,x in v.items()}
 if isinstance(v,(list,tuple)): return [jsonable(x) for x in v]
 if hasattr(v,'item'): return jsonable(v.item())
 if v is None or pd.isna(v): return None
 return v

def identifier(v):
 if v is None or pd.isna(v): return None
 if isinstance(v,(int,float)) and float(v).is_integer(): return str(int(v))
 return str(v)

def build(path):
 digest=hashlib.sha256(path.read_bytes()).hexdigest()
 print('Reading workbook:',path.name,flush=True)
 source=pd.read_excel(path,sheet_name='Online Retail',dtype=object,engine='openpyxl')
 if list(source.columns)!=RAW_FIELDS: raise ValueError('Exact eight-column source schema required')
 if source.isna().all(axis=1).any(): raise ValueError('Fully empty source row requires explicit policy')
 if not source.Quantity.map(lambda v:isinstance(v,(int,float)) and not isinstance(v,bool) and float(v).is_integer()).all():
  raise ValueError('Quantity must be a native whole number')
 if not source.UnitPrice.map(lambda v:isinstance(v,(int,float)) and not isinstance(v,bool)).all():
  raise ValueError('UnitPrice must be a native number')
 if not source.InvoiceDate.map(lambda v:isinstance(v,datetime)).all(): raise ValueError('Native datetime required')
 if source[['InvoiceNo','StockCode','Country']].isna().any().any(): raise ValueError('Missing critical key/country')
 prices=[]
 for v in source.UnitPrice:
  exact=Decimal(str(v))*10000
  if exact!=exact.to_integral_value(): raise ValueError('Price precision >4 decimals: explicit policy required')
  prices.append(int(exact))
 df=source.rename(columns={c:c+'_Raw' for c in RAW_FIELDS}).copy()
 df.insert(0,'SourceRow',np.arange(2,len(df)+2,dtype='int64'))
 df.insert(0,'SourceSheet','Online Retail'); df.insert(0,'SourceFile',path.name)
 df['RecordID']=path.name+'|Online Retail|'+df.SourceRow.map(lambda n:f'{n:09d}')
 for c in ['InvoiceNo','StockCode','CustomerID']: df[c]=source[c].map(identifier)
 df['DescriptionTrim']=source.Description.map(lambda v:v.strip() if isinstance(v,str) else None)
 df['Quantity']=pd.to_numeric(source.Quantity).astype('int64')
 df['UnitPrice']=pd.to_numeric(source.UnitPrice)
 df['InvoiceDate']=pd.to_datetime(source.InvoiceDate)
 df['Country']=source.Country
 df['Month']=df.InvoiceDate.dt.strftime('%Y-%m')
 df['PriceUnits4']=np.array(prices,dtype='int64')
 df['AmountUnits4']=df.Quantity*df.PriceUnits4
 df['LineAmount']=df.AmountUnits4/10000
 df['HasMissingCustomer']=df.CustomerID.isna() | (df.CustomerID=='')
 df['HasMissingDescription']=source.Description.isna() | (df.DescriptionTrim=='')
 df['HasNonTextDescription']=source.Description.notna() & ~source.Description.map(lambda v:isinstance(v,str))
 df['HasOuterWhitespace']=source.Description.map(lambda v:isinstance(v,str) and v!=v.strip())
 raw_cols=[c+'_Raw' for c in RAW_FIELDS]
 df['IsExactDuplicateMember']=df.duplicated(raw_cols,keep=False)
 df['IsDuplicateBeyondFirst']=df.duplicated(raw_cols,keep='first')
 exact_groups=df.loc[df.IsExactDuplicateMember,raw_cols].drop_duplicates().shape[0]
 desc_key=[c for c in raw_cols if c!='Description_Raw']
 desc_counts=df.groupby(desc_key,dropna=False,sort=False).Description_Raw.nunique(dropna=False)
 df['IsDescriptionVariantDuplicate']=pd.MultiIndex.from_frame(df[desc_key]).isin(desc_counts[desc_counts>1].index)
 dictionary=[]
 for code,g in df.groupby('StockCode',sort=True):
  descriptions=sorted(set(v for v in g.Description_Raw if isinstance(v,str) and v.strip()))
  trimmed=sorted(set(v.strip() for v in descriptions))
  normalized=sorted(set(' '.join(v.split()).upper() for v in descriptions))
  dictionary.append({'StockCode':code,'DescriptionVariantCount':len(descriptions),
      'TrimmedVariantCount':len(trimmed),'NormalizedVariantCount':len(normalized),
      'RawDescriptions':' | '.join(descriptions),'TrimmedDescriptions':' | '.join(trimmed),
      'SafeDisplayLabel':trimmed[0] if len(trimmed)==1 else '[multiple/absent descriptions; use code]',
      'SourceRows':len(g)})
 dictionary=pd.DataFrame(dictionary)
 before_rows=len(df); before_amount=int(df.AmountUnits4.sum())
 df=df.merge(dictionary[['StockCode','DescriptionVariantCount']],on='StockCode',how='left',validate='many_to_one',sort=False)
 if len(df)!=before_rows or int(df.AmountUnits4.sum())!=before_amount: raise AssertionError('Join inflation')
 df['HasMultipleDescriptions']=df.DescriptionVariantCount>1
 df['CodeCategory']=df.StockCode.map(RULES)
 common=df.StockCode.str.fullmatch(r'[0-9]{5}[A-Za-z]*')
 df.loc[df.CodeCategory.isna() & common,'CodeCategory']='Merchandise'
 df['CodeCategory']=df.CodeCategory.fillna('UnmappedCode')
 df['IsUnmappedCode']=df.CodeCategory=='UnmappedCode'
 df['IsCancellationCode']=df.InvoiceNo.str.startswith('C')
 df['IsInvoicePatternException']=~df.InvoiceNo.str.fullmatch(r'(?:C)?[0-9]{6}')
 regular_invoice=df.InvoiceNo.str.fullmatch(r'[0-9]{6}')
 df['IsNonCNegative']=(df.Quantity<0) & ~df.IsCancellationCode
 df['IsZeroPrice']=df.PriceUnits4==0; df['IsNegativePrice']=df.PriceUnits4<0
 df['IsLargeAmount']=df.AmountUnits4.abs()>=10000*10000
 special=df.StockCode.isin(SPECIAL)
 hard_review=df.HasNonTextDescription | df.IsUnmappedCode | df.IsInvoicePatternException
 df['TransactionClass']=np.select([
  special,hard_review,
  df.IsCancellationCode & (df.Quantity<0) & (df.PriceUnits4>0),
  df.IsNonCNegative & df.IsZeroPrice,
  regular_invoice & (df.Quantity>0) & (df.PriceUnits4>0),
 ],['SpecialOperation','UnresolvedReview','MerchandiseCancellation','ZeroPriceAdjustment','MerchandiseSale'],
 default='UnresolvedReview')
 df['ReportScope']=df.TransactionClass.isin(['MerchandiseSale','MerchandiseCancellation'])
 df['CustomerAnalysisEligible']=(df.TransactionClass=='MerchandiseSale') & ~df.HasMissingCustomer
 df['NeedsReview']=df[REVIEW_FLAGS].any(axis=1) | (df.TransactionClass=='UnresolvedReview')
 reasons=np.array(['']*len(df),dtype=object)
 for flag in REVIEW_FLAGS:
  mask=df[flag].to_numpy(); label=FLAG_DEFS[flag][0]
  reasons[mask]=[old+('; ' if old else '')+label for old in reasons[mask]]
 unresolved=(df.TransactionClass=='UnresolvedReview').to_numpy()
 reasons[unresolved]=[old+('; ' if old else '')+'Unresolved transaction class' for old in reasons[unresolved]]
 df['ReviewReasons']=reasons
 assert df.RecordID.is_unique and len(df)==len(source)
 assert digest==hashlib.sha256(path.read_bytes()).hexdigest()
 return df,dictionary,source,digest,exact_groups

def aggregates(df):
 tables={}
 class_names=['MerchandiseSale','MerchandiseCancellation','SpecialOperation','ZeroPriceAdjustment','UnresolvedReview']
 rows=[]
 for c in class_names:
  g=df.loc[df.TransactionClass==c]
  rows.append({'TransactionClass':c,'Rows':len(g),'InvoiceIDs':g.InvoiceNo.nunique(),
    'AmountUnits4':int(g.AmountUnits4.sum()),'SignedAmountGBP':int(g.AmountUnits4.sum())/10000,
    'ReviewFlagRows':int(g.NeedsReview.sum())})
 tables['Reconciliation']=pd.DataFrame(rows)
 sale=df.loc[df.TransactionClass=='MerchandiseSale']; cancel=df.loc[df.TransactionClass=='MerchandiseCancellation']
 known=sale.loc[~sale.HasMissingCustomer]; report=df.loc[df.ReportScope]
 totals={'SourceRows':len(df),'RawInvoiceIDs':df.InvoiceNo.nunique(),
 'RawSignedAmountGBP':int(df.AmountUnits4.sum())/10000,
 'ReportableSalesGBP':int(sale.AmountUnits4.sum())/10000,
 'MerchandiseCancellationGBP':-int(cancel.AmountUnits4.sum())/10000,
 'MerchandiseBalanceGBP':int(report.AmountUnits4.sum())/10000,
 'EligibleSalesInvoiceIDs':sale.InvoiceNo.nunique(),'SalesRows':len(sale),
 'KnownCustomerSalesRows':len(known),'KnownCustomerRowCoverage':len(known)/len(sale) if len(sale) else 0,
 'KnownCustomerAmountCoverage':int(known.AmountUnits4.sum())/int(sale.AmountUnits4.sum()) if sale.AmountUnits4.sum() else 0,
 'KnownCustomerIDs':known.CustomerID.nunique(),'ReviewRows':int(df.NeedsReview.sum()),
 'ReviewSignedAmountGBP':int(df.loc[df.NeedsReview,'AmountUnits4'].sum())/10000,
 'DuplicateExtraRows':int(df.IsDuplicateBeyondFirst.sum()),
 'DuplicateExtraSignedAmountGBP':int(df.loc[df.IsDuplicateBeyondFirst,'AmountUnits4'].sum())/10000}
 tables['KPI']=pd.DataFrame([{'Metric':k,'Value':v} for k,v in totals.items()])
 months=[]
 for month,g in df.groupby('Month',sort=True):
  s=g.loc[g.TransactionClass=='MerchandiseSale']; c=g.loc[g.TransactionClass=='MerchandiseCancellation']
  months.append({'Month':month,'SourceRows':len(g),'SalesRows':len(s),'SalesInvoiceIDs':s.InvoiceNo.nunique(),
    'SalesGBP':int(s.AmountUnits4.sum())/10000,'CancellationGBP':-int(c.AmountUnits4.sum())/10000,
    'BalanceGBP':int(s.AmountUnits4.sum()+c.AmountUnits4.sum())/10000,
    'ReviewRows':int(g.NeedsReview.sum()),'FirstObservedDate':g.InvoiceDate.min().date().isoformat(),
    'LastObservedDate':g.InvoiceDate.max().date().isoformat(),
    'PeriodNote':'Edge month: observed transaction window only' if month in ['2010-12','2011-12'] else 'Observed month; no completeness certification'})
 tables['Monthly']=pd.DataFrame(months)
 for name,key in [('Country','Country'),('Product','StockCode')]:
  result=[]
  for value,g in report.groupby(key,sort=True):
   s=g.loc[g.TransactionClass=='MerchandiseSale']; c=g.loc[g.TransactionClass=='MerchandiseCancellation']
   result.append({key:value,'SalesGBP':int(s.AmountUnits4.sum())/10000,
    'CancellationGBP':-int(c.AmountUnits4.sum())/10000,'BalanceGBP':int(g.AmountUnits4.sum())/10000,
    'SalesRows':len(s),'SalesInvoiceIDs':s.InvoiceNo.nunique()})
  tables[name]=pd.DataFrame(result).sort_values('SalesGBP',ascending=False,kind='stable').reset_index(drop=True)
 quality=[]
 for flag,(label,kind,definition) in FLAG_DEFS.items():
  g=df.loc[df[flag]]
  quality.append({'Flag':flag,'Label':label,'Kind':kind,'Rows':len(g),
      'SourceRowExample':int(g.SourceRow.iloc[0]) if len(g) else None,'Definition':definition})
 tables['Quality']=pd.DataFrame(quality)
 scenario=[]
 for label,g in [('Main: retain all rows',df),('Alternative: first exact record only',df.loc[~df.IsDuplicateBeyondFirst])]:
  s=g.loc[g.TransactionClass=='MerchandiseSale']; c=g.loc[g.TransactionClass=='MerchandiseCancellation']
  scenario.append({'Scenario':label,'SourceRows':len(g),'RawAmountUnits4':int(g.AmountUnits4.sum()),
   'RawSignedGBP':int(g.AmountUnits4.sum())/10000,'SalesGBP':int(s.AmountUnits4.sum())/10000,
   'CancellationGBP':-int(c.AmountUnits4.sum())/10000,'BalanceGBP':int(s.AmountUnits4.sum()+c.AmountUnits4.sum())/10000,
   'SalesInvoiceIDs':s.InvoiceNo.nunique(),'CustomerSalesIDs':s.CustomerID.nunique(),
   'ReviewRows':int(g.NeedsReview.sum())})
 tables['DuplicateScenario']=pd.DataFrame(scenario)
 raw=int(df.AmountUnits4.sum()); bridge=[{'Step':'Raw signed transactions','ChangeGBP':raw/10000,'BalanceGBP':raw/10000}]
 running=raw
 for c in ['SpecialOperation','ZeroPriceAdjustment','UnresolvedReview']:
  change=-int(df.loc[df.TransactionClass==c,'AmountUnits4'].sum()); running+=change
  bridge.append({'Step':'Separate '+c,'ChangeGBP':change/10000,'BalanceGBP':running/10000})
 bridge.append({'Step':'Merchandise signed balance','ChangeGBP':0.,'BalanceGBP':running/10000})
 tables['Bridge']=pd.DataFrame(bridge)
 special=[]
 for code,g in df.loc[df.TransactionClass=='SpecialOperation'].groupby('StockCode',sort=True):
  special.append({'StockCode':code,'Category':g.CodeCategory.iloc[0],'Rows':len(g),
   'SignedGBP':int(g.AmountUnits4.sum())/10000,'SourceRowExample':int(g.SourceRow.iloc[0])})
 tables['SpecialOperations']=pd.DataFrame(special)
 return tables,totals

def controls(df,source,digest,exact_groups,tables,baseline):
 out=[]
 def check(name,expected,actual,scope='Python reference'):
  good=expected==actual
  if isinstance(expected,float): good=abs(float(actual)-expected)<=0.0001
  out.append({'Control':name,'Expected':expected,'Actual':actual,'Status':'PASS' if good else 'FAIL','Scope':scope})
 check('Unique source RecordID',len(df),df.RecordID.nunique())
 check('Source rows preserved',len(source),len(df))
 check('Exactly one class per row',len(df),int(tables['Reconciliation'].Rows.sum()))
 check('No unassigned class',0,int(df.TransactionClass.isna().sum()))
 check('Class amounts reconcile (integer units)',int(df.AmountUnits4.sum()),int(tables['Reconciliation'].AmountUnits4.sum()))
 check('Product dictionary join rows preserved',len(source),len(df))
 check('Product dictionary nonmissing matches',0,int(df.DescriptionVariantCount.isna().sum()))
 check('Monthly source rows reconcile',len(df),int(tables['Monthly'].SourceRows.sum()))
 check('Monthly sales reconcile',int(df.loc[df.TransactionClass=='MerchandiseSale','AmountUnits4'].sum())/10000,float(tables['Monthly'].SalesGBP.sum()))
 check('Country sales reconcile',int(df.loc[df.TransactionClass=='MerchandiseSale','AmountUnits4'].sum())/10000,float(tables['Country'].SalesGBP.sum()))
 check('Product sales reconcile',int(df.loc[df.TransactionClass=='MerchandiseSale','AmountUnits4'].sum())/10000,float(tables['Product'].SalesGBP.sum()))
 check('Bridge closes to reporting scope',int(df.loc[df.ReportScope,'AmountUnits4'].sum())/10000,float(tables['Bridge'].BalanceGBP.iloc[-1]))
 check('Duplicate scenario rows removed',int(df.IsDuplicateBeyondFirst.sum()),int(tables['DuplicateScenario'].SourceRows.iloc[0]-tables['DuplicateScenario'].SourceRows.iloc[1]))
 check('Duplicate scenario amount delta (integer units)',int(df.loc[df.IsDuplicateBeyondFirst,'AmountUnits4'].sum()),int(tables['DuplicateScenario'].RawAmountUnits4.iloc[0]-tables['DuplicateScenario'].RawAmountUnits4.iloc[1]))
 if baseline:
  for name,expected,actual in [
   ('Original source SHA256',EXPECTED_HASH,digest),('Baseline rows',541909,len(df)),
   ('Baseline invoice IDs',25900,source.InvoiceNo.nunique()),('Baseline missing CustomerID',135080,int(source.CustomerID.isna().sum())),
   ('Baseline missing Description',1454,int(source.Description.isna().sum())),('Baseline duplicate groups',4879,exact_groups),
   ('Baseline duplicate members',10147,int(df.IsExactDuplicateMember.sum())),('Baseline duplicate extras',5268,int(df.IsDuplicateBeyondFirst.sum())),
   ('Baseline C-prefix rows',9288,int(df.IsCancellationCode.sum())),('Baseline non-C negative',1336,int(df.IsNonCNegative.sum())),
   ('Baseline zero prices',2515,int(df.IsZeroPrice.sum())),('Baseline negative prices',2,int(df.IsNegativePrice.sum())),
   ('Baseline multidescription codes',650,df.loc[df.HasMultipleDescriptions,'StockCode'].nunique()),
   ('Baseline outer whitespace',113452,int(df.HasOuterWhitespace.sum())),('Baseline raw amount units',97477479340,int(df.AmountUnits4.sum()))]:
   check(name,expected,actual)
 out.append({'Control':'Real Excel / Power Query execution and refresh','Expected':'User runs Excel checklist',
             'Actual':'Desktop Excel unavailable in Linux environment','Status':'PENDING','Scope':'Excel required'})
 return pd.DataFrame(out)

def main():
 parser=argparse.ArgumentParser(); parser.add_argument('--input',type=Path,default=ROOT/'raw'/'Online Retail.xlsx')
 parser.add_argument('--output',type=Path,default=ROOT/'reference'); parser.add_argument('--no-fact-export',action='store_true')
 args=parser.parse_args(); args.output.mkdir(parents=True,exist_ok=True)
 df,dictionary,source,digest,exact_groups=build(args.input)
 if args.input.resolve()==(ROOT/'raw'/'Online Retail.xlsx').resolve() and digest!=EXPECTED_HASH:
  raise ValueError('Original raw checksum changed; use an explicitly separate input for new/test data')
 tables,totals=aggregates(df); tables['ProductDictionary']=dictionary
 baseline=digest==EXPECTED_HASH
 tables['Controls']=controls(df,source,digest,exact_groups,tables,baseline)
 examples=[]
 for f in REVIEW_FLAGS:
  sample=df.loc[df[f]].head(8).copy(); sample.insert(0,'ExampleReason',f); examples.append(sample)
 tables['ReviewExamples']=pd.concat(examples).drop_duplicates(['RecordID','ExampleReason'])
 selected=[12,487,541,624,1445,9304,199126,199146,299985,420393,540423,540424]
 tables['BeforeAfter']=df.loc[df.SourceRow.isin(selected),[
  'RecordID','SourceRow','InvoiceNo_Raw','StockCode_Raw','Description_Raw','DescriptionTrim',
  'CustomerID_Raw','CustomerID','TransactionClass','ReviewReasons','LineAmount']]
 for name,table in tables.items(): table.to_csv(args.output/(name+'.csv'),index=False,encoding='utf-8-sig')
 summary={'engine':'Python reference; Power Query execution not claimed','rules_version':'1.0',
    'source_file':args.input.name,'sha256':digest,'rows':len(df),'totals':totals,
    'classes':tables['Reconciliation'].to_dict('records'),'checks':tables['Controls'].to_dict('records'),
    'price_scale':10000,'tolerance_gbp':0.0001,'exact_duplicate_groups':exact_groups,
    'date_min':df.InvoiceDate.min(),'date_max':df.InvoiceDate.max(),
    'built_at_utc':datetime.now(timezone.utc).isoformat()}
 (args.output/'summary.json').write_text(json.dumps(jsonable(summary),ensure_ascii=False,indent=2),encoding='utf-8')
 if not args.no_fact_export:
  df.to_csv(args.output/'ClassifiedTransactions.csv.gz',index=False,encoding='utf-8-sig',compression={'method':'gzip','mtime':0})
  review=ROOT/'review'; review.mkdir(exist_ok=True)
  df.loc[df.NeedsReview].to_csv(review/'ReviewRecords.csv.gz',index=False,encoding='utf-8-sig',compression={'method':'gzip','mtime':0})
  (ROOT/'.cache').mkdir(exist_ok=True)
  df.to_pickle(ROOT/'.cache'/'classified.pkl')
 failures=tables['Controls'].loc[tables['Controls'].Status=='FAIL']
 print(json.dumps(jsonable({'totals':totals,'classes':summary['classes'],'failed_controls':failures.to_dict('records')}),indent=2),flush=True)
 if len(failures): raise SystemExit('Reference controls failed')

if __name__=='__main__': main()
