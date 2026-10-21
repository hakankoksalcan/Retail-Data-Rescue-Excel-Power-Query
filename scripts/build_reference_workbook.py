"""Build the transparent reference workbook, actual-output figures, and delivery documents."""
from pathlib import Path
from datetime import datetime
import json,math,textwrap,os
os.environ.setdefault('MPLCONFIGDIR','/tmp/retail-matplotlib-cache')
import numpy as np,pandas as pd,xlsxwriter
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.lib.pagesizes import A4,landscape

ROOT=Path(__file__).resolve().parents[1]
REF=ROOT/'reference'; DOC=ROOT/'docs'; IMG=ROOT/'images'
BLUE='#2563EB'; ORANGE='#E58B37'; NAVY='#10263F'; TEAL='#177E89'; GRAY='#64748B'
summary=json.loads((REF/'summary.json').read_text()); totals=summary['totals']
tables={p.stem:pd.read_csv(p,keep_default_na=False) for p in REF.glob('*.csv')}
df=pd.read_pickle(ROOT/'.cache'/'classified.pkl')
from build_reference import REVIEW_FLAGS
tables['BeforeAfter']=df.loc[df.SourceRow.isin(tables['BeforeAfter'].SourceRow),list(tables['BeforeAfter'].columns)].copy()
tables['ReviewExamples']=pd.concat([df.loc[df[f]].head(8).assign(ExampleReason=f) for f in REVIEW_FLAGS]).drop_duplicates(['RecordID','ExampleReason'])
sales=df.TransactionClass=='MerchandiseSale'; canc=df.TransactionClass=='MerchandiseCancellation'
cube_work=df[['Country','Month','InvoiceNo','AmountUnits4','CustomerAnalysisEligible','NeedsReview']].copy()
cube_work['SaleUnits']=np.where(sales,df.AmountUnits4,0)
cube_work['CancelUnits']=np.where(canc,-df.AmountUnits4,0)
cube_work['SalesRows']=sales.astype(int)
cube=cube_work.groupby(['Country','Month'],sort=True).agg(
 SalesUnits=('SaleUnits','sum'),CancelUnits=('CancelUnits','sum'),SalesRows=('SalesRows','sum'),
 KnownCustomerSalesRows=('CustomerAnalysisEligible','sum'),ReviewRows=('NeedsReview','sum')).reset_index()
inv=df.loc[sales].groupby(['Country','Month']).InvoiceNo.nunique().rename('SalesInvoiceIDs').reset_index()
cube=cube.merge(inv,on=['Country','Month'],how='left',validate='one_to_one').fillna({'SalesInvoiceIDs':0})
cube['MonthDate']=pd.to_datetime(cube.Month+'-01')
cube['SalesGBP']=cube.SalesUnits/10000;cube['CancellationGBP']=cube.CancelUnits/10000
cube=cube[['Country','Month','MonthDate','SalesGBP','CancellationGBP','SalesRows','KnownCustomerSalesRows','SalesInvoiceIDs','ReviewRows']]
cube.SalesInvoiceIDs=cube.SalesInvoiceIDs.astype(int)
cube.to_csv(REF/'DashboardData.csv',index=False,encoding='utf-8-sig'); tables['DashboardData']=cube
customers=df.loc[df.CustomerAnalysisEligible].groupby('CustomerID',sort=True).agg(
 SalesRows=('RecordID','size'),AmountUnits4=('AmountUnits4','sum'),SalesInvoiceIDs=('InvoiceNo','nunique')).reset_index()
customers['SalesGBP']=customers.AmountUnits4/10000;customers=customers[['CustomerID','SalesRows','SalesGBP','SalesInvoiceIDs']]
customers.to_csv(REF/'Customers.csv',index=False,encoding='utf-8-sig');tables['Customers']=customers
cube_check={'Control':'Sales invoice cube partition','Expected':totals['EligibleSalesInvoiceIDs'],
 'Actual':int(cube.SalesInvoiceIDs.sum()),'Status':'PASS' if int(cube.SalesInvoiceIDs.sum())==totals['EligibleSalesInvoiceIDs'] else 'FAIL','Scope':'Python reference'}
controls=tables['Controls'];controls=controls.loc[controls.Control!='Sales invoice cube partition']
tables['Controls']=pd.concat([controls,pd.DataFrame([cube_check])],ignore_index=True)
tables['Controls'].to_csv(REF/'Controls.csv',index=False,encoding='utf-8-sig')
assert cube_check['Status']=='PASS'

bookpath=REF/'Reference_Workbook.xlsx'
wb=xlsxwriter.Workbook(bookpath,{'strings_to_formulas':False,'strings_to_urls':False,'nan_inf_to_errors':True})
wb.set_properties({'title':'Retail Data Rescue — Excel & Power Query delivery','author':'Hakan',
 'comments':'Independent portfolio. Python reference snapshot; Excel Power Query execution pending.'})
fmt={
 'title':wb.add_format({'bold':True,'font_size':24,'font_color':'white','bg_color':NAVY,'valign':'vcenter'}),
 'subtitle':wb.add_format({'font_size':10,'font_color':'white','bg_color':NAVY}),
 'label':wb.add_format({'bold':True,'font_color':GRAY,'font_size':10}),
 'money':wb.add_format({'num_format':'#,##0.00;[Red](#,##0.00)','font_color':NAVY}),
 'precise':wb.add_format({'num_format':'#,##0.0000;[Red](#,##0.0000)'}),
 'count':wb.add_format({'num_format':'#,##0'}),
 'percent':wb.add_format({'num_format':'0.0%'}),
 'date':wb.add_format({'num_format':'yyyy-mm-dd'}),
 'datetime':wb.add_format({'num_format':'yyyy-mm-dd hh:mm:ss'}),
 'input':wb.add_format({'bg_color':'#E0F2FE','font_color':NAVY,'bold':True,'border':1}),
 'wrap':wb.add_format({'text_wrap':True,'valign':'top','font_size':10}),
 'status':wb.add_format({'bg_color':'#FFF4DB','font_color':'#855C13','text_wrap':True,'font_size':10}),
 'card':wb.add_format({'bold':True,'font_size':22,'font_color':NAVY,'bg_color':'#F1F5F9','num_format':'#,##0.00'}),
 'card_count':wb.add_format({'bold':True,'font_size':22,'font_color':NAVY,'bg_color':'#F1F5F9','num_format':'#,##0'}),
 'card_pct':wb.add_format({'bold':True,'font_size':22,'font_color':TEAL,'bg_color':'#F1F5F9','num_format':'0.0%'}),
}
dash=wb.add_worksheet('Dashboard');dash.hide_gridlines(2);dash.set_zoom(85)
dash.set_column('A:A',2);dash.set_column('B:M',12);dash.set_column('N:N',2)
dash.set_row(0,40);dash.merge_range('B1:M1','RETAIL DATA RESCUE',fmt['title'])
dash.merge_range('B2:M2','Merchandise reporting | GBP | independent portfolio by Hakan',fmt['subtitle'])
dash.write('B4','Country',fmt['label']);dash.write('F4','Start month',fmt['label']);dash.write('J4','End month',fmt['label'])
dash.merge_range('B5:D5','All',fmt['input']);dash.merge_range('F5:H5','2010-12',fmt['input']);dash.merge_range('J5:L5','2011-12',fmt['input'])
dash.merge_range('B6:M6','Filters affect cards and monthly trend. Country/product rankings show full source.',fmt['label'])
countrycrit='IF($B$5="All","*",$B$5)'
start='DATE(VALUE(LEFT($F$5,4)),VALUE(RIGHT($F$5,2)),1)'
end='DATE(VALUE(LEFT($J$5,4)),VALUE(RIGHT($J$5,2)),1)'
def sumcube(column):return f'SUMIFS(DashboardData!${column}:${column},DashboardData!$A:$A,{countrycrit},DashboardData!$C:$C,">="&{start},DashboardData!$C:$C,"<="&{end})'
cards=[('B8:E8','B9:E10','REPORTABLE SALES (GBP)',sumcube('D'),totals['ReportableSalesGBP'],'card'),
 ('F8:I8','F9:I10','MERCHANDISE CANCELLATIONS',sumcube('E'),totals['MerchandiseCancellationGBP'],'card'),
 ('J8:M8','J9:M10','MERCHANDISE BALANCE (GBP)',sumcube('D')+'-'+sumcube('E'),totals['MerchandiseBalanceGBP'],'card'),
 ('B12:E12','B13:E14','ELIGIBLE SALES INVOICES',sumcube('H'),totals['EligibleSalesInvoiceIDs'],'card_count'),
 ('F12:I12','F13:I14','KNOWN CUSTOMER ROW COVERAGE','IFERROR('+sumcube('G')+'/'+sumcube('F')+',0)',totals['KnownCustomerRowCoverage'],'card_pct'),
 ('J12:M12','J13:M14','ROWS WITH REVIEW FLAGS',sumcube('I'),totals['ReviewRows'],'card_count')]
for labelrange,valuerange,label,formula,value,fkey in cards:
 dash.merge_range(labelrange,label,fmt['label']);dash.merge_range(valuerange,'',fmt[fkey])
 dash.write_formula(valuerange.split(':')[0],'='+formula,fmt[fkey],value)
dash.merge_range('B16:M17','',fmt['status']);dash.write_formula('B16','=RefreshInfo!B2&" | Built / refreshed UTC: "&RefreshInfo!A2',fmt['status'],
 'Python reference snapshot — Excel refresh pending | '+summary['built_at_utc'])
dash.merge_range('B18:M19','Observed edge windows: Dec 2010 = 01–23 Dec; Dec 2011 = 01–09 Dec. No full-year growth claim.',fmt['wrap'])

settings=wb.add_worksheet('Settings');settings.set_column('A:A',78);settings.set_column('B:B',25)
settings.write('A1','Source parameter — change to your local full file path',fmt['label'])
settings.add_table('A5:B6',{'name':'tblSource','style':'Table Style Medium 2','columns':[{'header':'SourcePath'},{'header':'ValidateBaseline'}],
 'data':[[r'C:\RetailDataRescue\raw\Online Retail.xlsx',True]]})
settings.write('A9','Baseline TRUE = published UCI source; FALSE = labelled append test/new source.',fmt['wrap'])
settings.write('A11','This is a regenerated Python reference workbook. The main workbook contains four embedded Excel audit queries.',fmt['status'])
settings.set_row(10,35)
settings.write('A14','See docs/KURULUM_YENILEME_TR.md for precise installation and table destinations.',fmt['wrap'])
settings.write('A17','Four-decimal fixed-scale arithmetic; reconciliation tolerance GBP 0.0001.',fmt['wrap'])

load_order=['KPI','Monthly','Country','Product','Reconciliation','Bridge','Quality','Controls','DuplicateScenario',
 'BeforeAfter','ProductDictionary','SpecialOperations','Customers','ReviewExamples','DashboardData']
for name in load_order:
 table=tables[name].copy(); ws=wb.add_worksheet(name);ws.freeze_panes(1,0);ws.hide_gridlines(2)
 ws.set_column(0,len(table.columns)-1,18)
 for i,col in enumerate(table.columns):
  if col in ['Description_Raw','DescriptionTrim','RawDescriptions','TrimmedDescriptions','ReviewReasons','Definition','Control','Step','Scenario','Metric','SafeDisplayLabel']:
   ws.set_column(i,i,45)
  elif col=='RecordID':ws.set_column(i,i,58)
  elif col in ['Country','TransactionClass','Flag','Label','ExampleReason']:ws.set_column(i,i,29)
 columns=[]
 for col in table.columns:
  f=None
  if 'GBP' in col or col=='LineAmount':f=fmt['precise']
  elif col in ['MonthDate','FirstObservedDate','LastObservedDate']:f=fmt['date']
  elif col.startswith('InvoiceDate'):f=fmt['datetime']
  columns.append({'header':col,**({'format':f} if f else {})})
 data=[]
 for row in table.itertuples(index=False,name=None):
  data.append([None if (v is None or (isinstance(v,float) and math.isnan(v))) else v.to_pydatetime() if isinstance(v,pd.Timestamp) else v.item() if hasattr(v,'item') else v for v in row])
 ws.add_table(0,0,len(data),len(columns)-1,{'name':'tbl'+name,'style':'Table Style Medium 2','columns':columns,'data':data})
 ws.set_row(0,28);ws.autofit(max_width=350);ws.set_landscape();ws.fit_to_pages(1,0)

refresh=wb.add_worksheet('RefreshInfo');refresh.add_table('A1:B2',{'name':'tblRefreshInfo','columns':[{'header':'ExecutedAtUTC'},{'header':'Mode'}],
 'data':[[summary['built_at_utc'],'Python reference snapshot — Excel refresh pending']]});refresh.set_column('A:B',65)
lists=wb.add_worksheet('Lists');countries=['All']+tables['Country'].Country.tolist();months=sorted(df.Month.unique().tolist())
lists.write(0,0,'All')
for i in range(300):
 lists.write_formula(i+1,0,f'=IF(Country!A{i+2}="","",Country!A{i+2})',None,countries[i+1] if i+1<len(countries) else '')
for i in range(1000):
 lists.write_formula(i,1,f'=IF(Monthly!A{i+2}="","",Monthly!A{i+2})',None,months[i] if i<len(months) else '')
wb.define_name('rdrCountryOptions','=Lists!$A$1:INDEX(Lists!$A:$A,COUNTA(Country!$A:$A))')
wb.define_name('rdrMonthOptions','=Lists!$B$1:INDEX(Lists!$B:$B,COUNTA(Monthly!$A:$A)-1)')
lists.hide()
dash.data_validation('B5',{'validate':'list','source':'=rdrCountryOptions'})
dash.data_validation('F5',{'validate':'list','source':'=rdrMonthOptions'})
dash.data_validation('J5',{'validate':'list','source':'=rdrMonthOptions'})
chartdata=wb.add_worksheet('ChartData');monthly=tables['Monthly']
for i in range(1000):
 r=i+2;row=i+1
 cachedmonth=monthly.Month.iloc[i] if i<len(monthly) else ''
 chartdata.write_formula(row,0,f'=IF(Monthly!A{r}="","",Monthly!A{r})',None,cachedmonth)
 criteria=f'DashboardData!$A:$A,IF(Dashboard!$B$5="All","*",Dashboard!$B$5),DashboardData!$B:$B,A{r}'
 inside=f'AND(A{r}>=Dashboard!$F$5,A{r}<=Dashboard!$J$5,A{r}<>"")'
 for col,cubecol,sign in [(1,'D',1),(2,'E',-1)]:
  value=(float(monthly.SalesGBP.iloc[i]) if col==1 else -float(monthly.CancellationGBP.iloc[i])) if i<len(monthly) else '#N/A'
  chartdata.write_formula(row,col,f'=IF({inside},{"-" if sign<0 else ""}SUMIFS(DashboardData!${cubecol}:${cubecol},{criteria}),NA())',fmt['money'],value)
chartdata.write_row(0,0,['Month','Sales','Cancellations (negative)']);chartdata.hide()
for label,col in [('rdrMonths','A'),('rdrSales','B'),('rdrCancels','C')]:
 wb.define_name('Dashboard!'+label,f'=ChartData!${col}$2:INDEX(ChartData!${col}:${col},COUNTA(Monthly!$A:$A))')
chart=wb.add_chart({'type':'column'})
for name,values,color in [('Merchandise sales','rdrSales',BLUE),('Merchandise cancellations','rdrCancels',ORANGE)]:
 chart.add_series({'name':name,'categories':'=Dashboard!rdrMonths','values':'=Dashboard!'+values,
    'categories_data':monthly.Month.tolist(),
    'values_data':monthly.SalesGBP.tolist() if values=='rdrSales' else (-monthly.CancellationGBP).tolist(),
    'fill':{'color':color},'border':{'none':True}})
chart.set_title({'name':'Monthly merchandise activity — selected filters'})
chart.set_y_axis({'name':'GBP (signed)','num_format':'#,##0','major_gridlines':{'visible':True,'line':{'color':'#E2E8F0'}}})
chart.set_legend({'position':'bottom'});chart.set_chartarea({'border':{'none':True}});chart.set_size({'width':930,'height':310})
dash.insert_chart('B21',chart)
for anchor,name,label,key in [('B38','Country','Country sales — full source','Country'),('H38','Product','Product code sales — full source','StockCode')]:
 c=wb.add_chart({'type':'bar'});n=5
 valuecol=list(tables[name].columns).index('SalesGBP')
 c.add_series({'name':label,'categories':[name,1,0,n,0],'values':[name,1,valuecol,n,valuecol],
   'fill':{'color':BLUE if name=='Country' else TEAL},'border':{'none':True}})
 c.set_title({'name':label});c.set_legend({'none':True});c.set_y_axis({'reverse':True});c.set_x_axis({'num_format':'#,##0','name':'GBP'})
 c.set_chartarea({'border':{'none':True}});c.set_size({'width':450,'height':280});dash.insert_chart(anchor,c)
dash.merge_range('B54:M55','Scope: positive-priced merchandise sales; eligible merchandise cancellations shown separately. Special operations excluded. Exact duplicates retained.',fmt['wrap'])
dash.merge_range('B56:M57','Review flags overlap reporting scope. Merchandise balance is not accounting revenue. Cancellation does not establish a physical return.',fmt['wrap'])
dash.set_landscape();dash.set_paper(9);dash.fit_to_pages(1,1);dash.print_area('B1:M57');dash.set_margins(.25,.25,.3,.3)
wb.close()

# Actual-output figures: generated from the computed reference tables, never simulated screenshots.
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':10,'axes.spines.top':False,'axes.spines.right':False})
fig=plt.figure(figsize=(15,10),facecolor='#F8FAFC');gs=fig.add_gridspec(3,2,height_ratios=[.85,1.8,1.5],hspace=.48,wspace=.35)
ax=fig.add_subplot(gs[0,:]);ax.axis('off')
ax.text(0,1,'RETAIL DATA RESCUE',fontsize=26,fontweight='bold',color=NAVY,transform=ax.transAxes)
ax.text(0,.78,'GBP • 01 Dec 2010–09 Dec 2011 • All countries • duplicates retained',color=GRAY,transform=ax.transAxes)
for x,label,val in [(0,'REPORTABLE SALES',f"£{totals['ReportableSalesGBP']:,.2f}"),(.34,'MERCHANDISE CANCELLATIONS',f"£{totals['MerchandiseCancellationGBP']:,.2f}"),(.72,'CUSTOMER ROW COVERAGE',f"{totals['KnownCustomerRowCoverage']:.1%}")]:
 ax.text(x,.40,label,fontsize=10,color=GRAY,transform=ax.transAxes);ax.text(x,.09,val,fontsize=24,fontweight='bold',color=NAVY,transform=ax.transAxes)
ax=fig.add_subplot(gs[1,:]);x=np.arange(len(monthly));ax.bar(x,monthly.SalesGBP/1e6,color=BLUE,label='Merchandise sales')
ax.bar(x,-monthly.CancellationGBP/1e6,color=ORANGE,label='Merchandise cancellations');ax.axhline(0,color=GRAY,lw=.8)
ax.set_xticks(x,monthly.Month,rotation=30,ha='right');ax.set_ylabel('GBP millions (signed)');ax.set_title('Monthly activity • edge months are observed windows, not comparable full months',loc='left',fontweight='bold');ax.legend(frameon=False,ncol=2)
for axis,table,key,title,color in [(fig.add_subplot(gs[2,0]),tables['Country'],'Country','Country sales',BLUE),(fig.add_subplot(gs[2,1]),tables['Product'],'StockCode','Product code sales',TEAL)]:
 t=table.head(5).iloc[::-1];axis.barh(t[key].astype(str),t.SalesGBP,color=color);axis.xaxis.set_major_formatter(FuncFormatter(lambda n,pos:f'£{n/1000:,.0f}k'));axis.set_title(title,loc='left',fontweight='bold')
fig.text(.07,.025,'PYTHON REFERENCE OUTPUT • Excel Power Query execution/refresh pending. Not accounting revenue or physical returns.',color=GRAY,fontsize=10)
fig.savefig(IMG/'dashboard_reference.png',dpi=160,bbox_inches='tight');plt.close(fig)
before=tables['BeforeAfter']; rows=[12,624,1445,487,541,420393]
fig,ax=plt.subplots(figsize=(16,4.5));ax.axis('off')
data=[]
for row in rows:
 r=before.loc[before.SourceRow==row].iloc[0]
 raw='<blank>' if pd.isna(r.Description_Raw) else repr(r.Description_Raw)
 trim='<null>' if pd.isna(r.DescriptionTrim) else repr(r.DescriptionTrim)
 data.append([row,textwrap.shorten(raw,width=42,placeholder='…'),textwrap.shorten(trim,width=42,placeholder='…'),r.TransactionClass,textwrap.shorten(r.ReviewReasons,width=36,placeholder='…')])
tbl=ax.table(cellText=data,colLabels=['Excel row','Original description','Trimmed description','Class (record retained)','Review reason'],loc='center',cellLoc='left',colWidths=[.07,.23,.23,.20,.27]);tbl.auto_set_font_size(False);tbl.set_fontsize(9);tbl.scale(1,2)
for (r,c),cell in tbl.get_celld().items():
 if r==0:cell.set_facecolor(NAVY);cell.set_text_props(color='white',weight='bold')
 else:cell.set_facecolor('#F1F5F9' if r%2 else 'white')
ax.set_title('BEFORE / AFTER • same source rows • only documented transformations',loc='left',color=NAVY,fontweight='bold',pad=20)
fig.savefig(IMG/'before_after_reference.png',dpi=160,bbox_inches='tight');plt.close(fig)

# One-page client summary, truthful about runtime and portfolio status.
c=canvas.Canvas(str(ROOT/'Client_Handover.pdf'),pagesize=A4);w,h=A4
c.setFillColor(colors.HexColor(NAVY));c.rect(0,h-116,w,116,fill=1,stroke=0);c.setFillColor(colors.white)
c.setFont('Helvetica-Bold',24);c.drawString(40,h-52,'RETAIL DATA RESCUE')
c.setFont('Helvetica',11);c.drawString(40,h-78,'Excel & Power Query | Independent portfolio by Hakan')
c.drawString(40,h-98,'Data quality case study | Reference reporting with Excel audit queries')
y=h-145
def pdfblock(title,body):
 global y
 c.setFillColor(colors.HexColor(NAVY));c.setFont('Helvetica-Bold',12);c.drawString(40,y,title);y-=19
 c.setFillColor(colors.HexColor('#334155'));c.setFont('Helvetica',10)
 for line in textwrap.wrap(body,96):c.drawString(40,y,line);y-=14
 y-=14
pdfblock('Business problem','A retail transaction export mixes merchandise sales, cancellations, fees, adjustments and ambiguous records. The process makes rules and source-to-report reconciliations explicit.')
pdfblock('Prepared outputs','Excel reference dashboard with filters; complete M sources; immutable raw workbook; complete review register and reproducible classified export; decision log; quality audit; reconciliation and duplicate sensitivity; reproducible Python controls.')
pdfblock('Verified reference results',f"541,909 rows explained. Merchandise sales GBP {totals['ReportableSalesGBP']:,.3f}; merchandise cancellations GBP {totals['MerchandiseCancellationGBP']:,.2f}; signed merchandise balance GBP {totals['MerchandiseBalanceGBP']:,.3f}. Missing customer IDs do not exclude eligible sales.")
pdfblock('Key decisions','Exact duplicates remain in the main result. The first-record-only alternative removes 5,268 rows and GBP 21,740.98 from raw signed transactions. Special codes are separated by explicit rules. C means cancellation, not confirmed physical return.')
pdfblock('Validation and acceptance','Python source checks, reconciliation and repeat/append tests passed. Microsoft M parser syntax checks passed. Four audit queries are embedded in the main workbook; source and missing-value refreshes were checked in Excel 2021. Full-model M sources are supplied separately; their end-to-end Excel refresh acceptance is pending.')
pdfblock('Limits','Internal source reconciliation only; no external accounting validation, cost/profit or tax conclusion. December 2010 observed dates end on 23 Dec; December 2011 ends on 9 Dec. Rankings and repeat scenarios need their stated scopes.')
c.setFont('Helvetica',8);c.setFillColor(colors.HexColor(GRAY));c.drawString(40,45,'Source: Chen, D. (2015), Online Retail, UCI. DOI 10.24432/C5BW33. CC BY 4.0.')
c.drawString(40,32,'See docs/KURULUM_YENILEME_TR.md and validation/EXCEL_ACCEPTANCE_CHECKLIST.csv.');c.showPage();c.save()

# A separate dashboard reference PDF makes visualization provenance explicit.
c=canvas.Canvas(str(ROOT/'Dashboard_Reference.pdf'),pagesize=landscape(A4));pw,ph=landscape(A4)
c.drawImage(str(IMG/'dashboard_reference.png'),15,15,width=pw-30,height=ph-30,preserveAspectRatio=True,anchor='c');c.showPage();c.save()
print('Created reference workbook, figures, dashboard PDF and one-page client handover')
