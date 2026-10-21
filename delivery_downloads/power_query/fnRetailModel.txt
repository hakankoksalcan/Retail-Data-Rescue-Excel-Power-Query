// Complete Power Query function. Excel execution is pending; syntax is validated separately.
// All row identities are added before grouping, joins, filtering, or sorting.
(SourcePath as text, optional ValidateBaseline as nullable logical) as record =>
let
    Baseline = if ValidateBaseline = null then true else ValidateBaseline,
    Fields = {"InvoiceNo","StockCode","Description","Quantity","InvoiceDate","UnitPrice","CustomerID","Country"},
    Scale = 10000,
    LargeAmountGBP = 10000,
    SourceFile = List.Last(Text.Split(Text.Replace(SourcePath,"\","/"),"/")),
    Workbook = Excel.Workbook(File.Contents(SourcePath),null,true),
    Sheet = Workbook{[Item="Online Retail",Kind="Sheet"]}[Data],
    Headers = Table.PromoteHeaders(Sheet,[PromoteAllScalars=true,Culture="en-GB"]),
    Schema = if List.Sort(Table.ColumnNames(Headers)) <> List.Sort(Fields) then
        error Error.Record("SchemaMismatch","Exactly the eight documented fields are required",Table.ColumnNames(Headers)) else Headers,
    Invalid = Table.SelectRows(Schema, each
        [InvoiceNo] = null or [StockCode] = null or [Country] = null or
        not Value.Is([Quantity],type number) or not Value.Is([UnitPrice],type number) or
        not (Value.Is([InvoiceDate],type datetime) or Value.Is([InvoiceDate],type date)) or
        (if Value.Is([Quantity],type number) then Number.Mod([Quantity],1) <> 0 else true) or
        (if Value.Is([UnitPrice],type number) then
            Number.Abs([UnitPrice]-Number.Round([UnitPrice]*Scale,0)/Scale)>0.0000000001 else true)),
    Validated = if Table.RowCount(Invalid)>0 then
        error Error.Record("UnsupportedSource","Invalid key, native date/number, quantity or price precision; source was not changed",Table.FirstN(Invalid,5)) else Schema,
    Indexed = Table.AddIndexColumn(Validated,"SourceRow",2,1,Int64.Type),
    FileAdded = Table.AddColumn(Indexed,"SourceFile",each SourceFile,type text),
    SheetAdded = Table.AddColumn(FileAdded,"SourceSheet",each "Online Retail",type text),
    IDAdded = Table.AddColumn(SheetAdded,"RecordID",each [SourceFile]&"|"&[SourceSheet]&"|"&Text.PadStart(Text.From([SourceRow]),9,"0"),type text),
    Raw = Table.RenameColumns(IDAdded,List.Transform(Fields,each {_,_&"_Raw"})),
    RawFields = List.Transform(Fields,each _&"_Raw"),
    AsID = (v as any) as nullable text => if v=null then null else Text.From(v,"en-GB"),
    WithCleanRecord = Table.AddColumn(Raw,"Clean",each [
        InvoiceNo=AsID([InvoiceNo_Raw]),StockCode=AsID([StockCode_Raw]),
        CustomerID=AsID([CustomerID_Raw]),
        DescriptionTrim=if Value.Is([Description_Raw],type text) then Text.Trim([Description_Raw]) else null,
        Quantity=Int64.From([Quantity_Raw]),UnitPrice=Number.From([UnitPrice_Raw]),
        InvoiceDate=DateTime.From([InvoiceDate_Raw]),Country=Text.From([Country_Raw]),
        PriceUnits4=Int64.From(Number.Round([UnitPrice_Raw]*Scale,0))]),
    Clean = Table.ExpandRecordColumn(WithCleanRecord,"Clean",{"InvoiceNo","StockCode","CustomerID","DescriptionTrim","Quantity","UnitPrice","InvoiceDate","Country","PriceUnits4"}),
    WithAmount = Table.AddColumn(Clean,"AmountUnits4",each Int64.From([Quantity]*[PriceUnits4]),Int64.Type),
    WithDisplay = Table.AddColumn(WithAmount,"LineAmount",each [AmountUnits4]/Scale,type number),
    Stage = Table.AddColumn(WithDisplay,"Month",each Date.ToText(Date.From([InvoiceDate]),"yyyy-MM","en-GB"),type text),
    OriginalRows = Table.RowCount(Stage),
    Sum0 = (xs as list) as number => if List.IsEmpty(xs) then 0 else List.Sum(xs),
    Amount = (t as table) as number => Sum0(t[AmountUnits4]),
    OriginalAmount = Amount(Stage),
    ExactGroups = Table.Group(Stage,RawFields,{
        {"ExactGroupRows",each Table.RowCount(_),Int64.Type},
        {"FirstSourceRow",each List.Min([SourceRow]),Int64.Type}}),
    ExactJoin = Table.NestedJoin(Stage,RawFields,ExactGroups,RawFields,"Exact",JoinKind.LeftOuter),
    ExactExpanded = Table.ExpandTableColumn(ExactJoin,"Exact",{"ExactGroupRows","FirstSourceRow"}),
    VariantKeys = List.RemoveItems(RawFields,{"Description_Raw"}),
    VariantGroups = Table.Group(Stage,VariantKeys,{
        {"SameKeyDescriptionCount",each List.Count(List.Distinct([Description_Raw])),Int64.Type}}),
    VariantJoin = Table.NestedJoin(ExactExpanded,VariantKeys,VariantGroups,VariantKeys,"Variants",JoinKind.LeftOuter),
    VariantExpanded = Table.ExpandTableColumn(VariantJoin,"Variants",{"SameKeyDescriptionCount"}),
    TextDescriptions = (xs as list) as list => List.Sort(List.Distinct(List.Select(xs,each Value.Is(_,type text) and Text.Trim(_)<>""))),
    Normalize = (s as text) as text => Text.Upper(Text.Combine(List.Select(Text.SplitAny(Text.Trim(s)," #(tab)#(cr)#(lf)"),each _<>"")," "),"en-GB"),
    DictionaryRecord = Table.Group(Stage,{"StockCode"},{
        {"Dictionary",each let
            ds=TextDescriptions([Description_Raw]),ts=List.Sort(List.Distinct(List.Transform(ds,each Text.Trim(_)))),
            ns=List.Distinct(List.Transform(ds,each Normalize(_)))
        in [DescriptionVariantCount=List.Count(ds),TrimmedVariantCount=List.Count(ts),
            NormalizedVariantCount=List.Count(ns),RawDescriptions=Text.Combine(ds," | "),
            TrimmedDescriptions=Text.Combine(ts," | "),
            SafeDisplayLabel=if List.Count(ts)=1 then ts{0} else "[multiple/absent descriptions; use code]",
            SourceRows=Table.RowCount(_)]}}),
    ProductDictionary = Table.ExpandRecordColumn(DictionaryRecord,"Dictionary",{"DescriptionVariantCount","TrimmedVariantCount","NormalizedVariantCount","RawDescriptions","TrimmedDescriptions","SafeDisplayLabel","SourceRows"}),
    DictionaryJoin = Table.NestedJoin(VariantExpanded,{"StockCode"},ProductDictionary,{"StockCode"},"Dictionary",JoinKind.LeftOuter),
    DictionaryExpanded = Table.ExpandTableColumn(DictionaryJoin,"Dictionary",{"DescriptionVariantCount"}),
    // Explicit descriptor-supported rules; nonstandard format alone is not a service rule.
    CodeRules = #table(type table [StockCode=text,CodeCategory=text],{{"POST","Logistics"},{"C2","Logistics"},{"DOT","Logistics"},{"D","Discount"},{"M","Manual"},{"m","Manual"},{"BANK CHARGES","BankFee"},{"S","Sample"},{"AMAZONFEE","PlatformFee"},{"B","AccountingAdjustment"},{"CRUK","Commission"},{"gift_0001_10","Voucher"},{"gift_0001_20","Voucher"},{"gift_0001_30","Voucher"},{"gift_0001_40","Voucher"},{"gift_0001_50","Voucher"},{"PADS","MerchandiseException"},{"DCGS0076","MerchandiseException"},{"DCGS0003","MerchandiseException"},{"DCGS0070","MerchandiseException"},{"DCGS0055","MerchandiseException"},{"DCGS0072","MerchandiseException"},{"DCGS0074","MerchandiseException"},{"DCGS0069","MerchandiseException"},{"DCGS0057","MerchandiseException"},{"DCGSSBOY","MerchandiseException"},{"DCGSSGIRL","MerchandiseException"},{"DCGS0004","MerchandiseException"},{"DCGS0073","MerchandiseException"},{"DCGS0071","MerchandiseException"},{"DCGS0068","MerchandiseException"},{"DCGS0067","MerchandiseException"},{"DCGS0066P","MerchandiseException"}}),
    RuleJoin = Table.NestedJoin(DictionaryExpanded,{"StockCode"},CodeRules,{"StockCode"},"Rule",JoinKind.LeftOuter),
    RuleExpanded = Table.ExpandTableColumn(RuleJoin,"Rule",{"CodeCategory"},{"MappedCategory"}),
    DigitsOnly = (s as text) as logical => Text.Length(s)>0 and Text.Select(s,{"0".."9"})=s,
    CommonCode = (s as text) as logical => Text.Length(s)>=5 and DigitsOnly(Text.Start(s,5)) and
        Text.Select(Text.Range(s,5),{"A".."Z","a".."z"})=Text.Range(s,5),
    Categorized = Table.AddColumn(RuleExpanded,"CodeCategory",each if [MappedCategory]<>null then [MappedCategory]
        else if CommonCode([StockCode]) then "Merchandise" else "UnmappedCode",type text),
    FlagRecord = Table.AddColumn(Categorized,"Flags",each let
        c=Text.StartsWith([InvoiceNo],"C"),inv=if c then Text.Range([InvoiceNo],1) else [InvoiceNo]
        in [HasMissingCustomer=[CustomerID]=null or [CustomerID]="",
        HasMissingDescription=[Description_Raw]=null or [DescriptionTrim]="",
        HasNonTextDescription=[Description_Raw]<>null and not Value.Is([Description_Raw],type text),
        HasOuterWhitespace=if Value.Is([Description_Raw],type text) then [Description_Raw]<>[DescriptionTrim] else false,
        IsExactDuplicateMember=[ExactGroupRows]>1,IsDuplicateBeyondFirst=[SourceRow]>[FirstSourceRow],
        IsDescriptionVariantDuplicate=[SameKeyDescriptionCount]>1,
        HasMultipleDescriptions=[DescriptionVariantCount]>1,
        IsCancellationCode=c,IsInvoicePatternException=not (Text.Length(inv)=6 and DigitsOnly(inv)),
        IsNonCNegative=[Quantity]<0 and not c,IsZeroPrice=[PriceUnits4]=0,
        IsNegativePrice=[PriceUnits4]<0,IsLargeAmount=Number.Abs([AmountUnits4])>=LargeAmountGBP*Scale,
        IsUnmappedCode=[CodeCategory]="UnmappedCode"]),
    FlagNames = {"HasMissingCustomer","HasMissingDescription","HasNonTextDescription","HasOuterWhitespace", "IsExactDuplicateMember","IsDuplicateBeyondFirst","IsDescriptionVariantDuplicate","HasMultipleDescriptions", "IsCancellationCode","IsInvoicePatternException","IsNonCNegative","IsZeroPrice","IsNegativePrice","IsLargeAmount","IsUnmappedCode"},
    Flagged = Table.ExpandRecordColumn(FlagRecord,"Flags",FlagNames),
    Classified = Table.AddColumn(Flagged,"TransactionClass",each
        if List.Contains({"Logistics","Discount","Manual","BankFee","Sample","PlatformFee","AccountingAdjustment","Commission","Voucher"},[CodeCategory]) then "SpecialOperation"
        else if [HasNonTextDescription] or [IsUnmappedCode] or [IsInvoicePatternException] then "UnresolvedReview"
        else if [IsCancellationCode] and [Quantity]<0 and [PriceUnits4]>0 then "MerchandiseCancellation"
        else if [IsNonCNegative] and [IsZeroPrice] then "ZeroPriceAdjustment"
        else if not [IsCancellationCode] and [Quantity]>0 and [PriceUnits4]>0 then "MerchandiseSale"
        else "UnresolvedReview",type text),
    ReviewFlags = {"HasMissingCustomer","HasMissingDescription","HasNonTextDescription","IsExactDuplicateMember","IsDescriptionVariantDuplicate","HasMultipleDescriptions","IsNonCNegative","IsZeroPrice","IsNegativePrice","IsLargeAmount","IsUnmappedCode"},
    ReviewLabels = {"Missing customer ID","Missing description","Non-text description","Exact duplicate group member","Description-only duplicate candidate","Code with multiple descriptions","Negative quantity without C","Zero price","Negative price","Large absolute line amount","Unmapped nonstandard code"},
    ScopeRecord = Table.AddColumn(Classified,"Scope",each let r=_ in [
        ReportScope=List.Contains({"MerchandiseSale","MerchandiseCancellation"},[TransactionClass]),
        CustomerAnalysisEligible=[TransactionClass]="MerchandiseSale" and not [HasMissingCustomer],
        NeedsReview=List.AnyTrue(List.Transform(ReviewFlags,(f)=>Record.Field(r,f))) or [TransactionClass]="UnresolvedReview",
        ReviewReasons=Text.Combine(List.Combine({List.Transform(List.Select(List.Positions(ReviewFlags),(i)=>Record.Field(r,ReviewFlags{i})),(i)=>ReviewLabels{i}),
            if [TransactionClass]="UnresolvedReview" then {"Unresolved transaction class"} else {}}),"; ")]),
    ExpandedScope = Table.ExpandRecordColumn(ScopeRecord,"Scope",{"ReportScope","CustomerAnalysisEligible","NeedsReview","ReviewReasons"}),
    RemovedHelpers = Table.RemoveColumns(ExpandedScope,{"MappedCategory","ExactGroupRows","FirstSourceRow","SameKeyDescriptionCount"}),
    Facts = if Table.RowCount(RemovedHelpers)<>OriginalRows or Amount(RemovedHelpers)<>OriginalAmount then
        error Error.Record("JoinInflation","Dictionary or grouping joins changed row count or amount",null) else RemovedHelpers,
    ClassNames = {"MerchandiseSale","MerchandiseCancellation","SpecialOperation","ZeroPriceAdjustment","UnresolvedReview"},
    ClassRows = (t as table,c as text) as table => Table.SelectRows(t,each [TransactionClass]=c),
    InvoiceCount = (t as table) as number => List.Count(List.Distinct(t[InvoiceNo])),
    Sales = ClassRows(Facts,"MerchandiseSale"),
    Cancellations = ClassRows(Facts,"MerchandiseCancellation"),
    ReportTransactions = Table.SelectRows(Facts,each [ReportScope]),
    CustomerTransactions = Table.SelectRows(Facts,each [CustomerAnalysisEligible]),
    ReviewRecords = Table.SelectRows(Facts,each [NeedsReview]),
    Reconciliation = Table.FromRecords(List.Transform(ClassNames,(c)=>let t=ClassRows(Facts,c) in
        [TransactionClass=c,Rows=Table.RowCount(t),InvoiceIDs=InvoiceCount(t),AmountUnits4=Amount(t),
            SignedAmountGBP=Amount(t)/Scale,ReviewFlagRows=Table.RowCount(Table.SelectRows(t,each [NeedsReview]))])),
    Kpis = (t as table) as table => let s=ClassRows(t,"MerchandiseSale"),c=ClassRows(t,"MerchandiseCancellation"),
        known=Table.SelectRows(s,each not [HasMissingCustomer]),review=Table.SelectRows(t,each [NeedsReview]),
        extras=Table.SelectRows(t,each [IsDuplicateBeyondFirst])
        in #table(type table[Metric=text,Value=number],{
        {"SourceRows",Table.RowCount(t)},{"RawInvoiceIDs",InvoiceCount(t)},{"RawSignedAmountGBP",Amount(t)/Scale},
        {"ReportableSalesGBP",Amount(s)/Scale},{"MerchandiseCancellationGBP",-Amount(c)/Scale},
        {"MerchandiseBalanceGBP",(Amount(s)+Amount(c))/Scale},{"EligibleSalesInvoiceIDs",InvoiceCount(s)},
        {"SalesRows",Table.RowCount(s)},{"KnownCustomerSalesRows",Table.RowCount(known)},
        {"KnownCustomerRowCoverage",if Table.RowCount(s)=0 then 0 else Table.RowCount(known)/Table.RowCount(s)},
        {"KnownCustomerAmountCoverage",if Amount(s)=0 then 0 else Amount(known)/Amount(s)},
        {"KnownCustomerIDs",List.Count(List.Distinct(known[CustomerID]))},{"ReviewRows",Table.RowCount(review)},
        {"ReviewSignedAmountGBP",Amount(review)/Scale},{"DuplicateExtraRows",Table.RowCount(extras)},
        {"DuplicateExtraSignedAmountGBP",Amount(extras)/Scale}}),
    KPI = Kpis(Facts),
    Monthly = Table.Sort(Table.Group(Facts,{"Month"},{
        {"SourceRows",each Table.RowCount(_),Int64.Type},
        {"SalesRows",each Table.RowCount(ClassRows(_,"MerchandiseSale")),Int64.Type},
        {"SalesInvoiceIDs",each InvoiceCount(ClassRows(_,"MerchandiseSale")),Int64.Type},
        {"SalesGBP",each Amount(ClassRows(_,"MerchandiseSale"))/Scale,type number},
        {"CancellationGBP",each -Amount(ClassRows(_,"MerchandiseCancellation"))/Scale,type number},
        {"BalanceGBP",each (Amount(ClassRows(_,"MerchandiseSale"))+Amount(ClassRows(_,"MerchandiseCancellation")))/Scale,type number},
        {"ReviewRows",each Table.RowCount(Table.SelectRows(_,each [NeedsReview])),Int64.Type},
        {"FirstObservedDate",each Date.From(List.Min([InvoiceDate])),type date},
        {"LastObservedDate",each Date.From(List.Max([InvoiceDate])),type date}}),{{"Month",Order.Ascending}}),
    Aggregate = (t as table,key as text) as table => Table.Sort(Table.Group(t,{key},{
        {"SalesGBP",each Amount(ClassRows(_,"MerchandiseSale"))/Scale,type number},
        {"CancellationGBP",each -Amount(ClassRows(_,"MerchandiseCancellation"))/Scale,type number},
        {"BalanceGBP",each Amount(_)/Scale,type number},
        {"SalesRows",each Table.RowCount(ClassRows(_,"MerchandiseSale")),Int64.Type},
        {"SalesInvoiceIDs",each InvoiceCount(ClassRows(_,"MerchandiseSale")),Int64.Type}}),{{"SalesGBP",Order.Descending}}),
    Country = Aggregate(ReportTransactions,"Country"),
    Product = Aggregate(ReportTransactions,"StockCode"),
    Customers = Table.Group(CustomerTransactions,{"CustomerID"},{
        {"SalesRows",each Table.RowCount(_),Int64.Type},{"SalesGBP",each Amount(_)/Scale,type number},
        {"SalesInvoiceIDs",each InvoiceCount(_),Int64.Type}}),
    DashboardData = Table.Group(Facts,{"Country","Month"},{
        {"MonthDate",each Date.FromText(List.First([Month])&"-01"),type date},
        {"SalesGBP",each Amount(ClassRows(_,"MerchandiseSale"))/Scale,type number},
        {"CancellationGBP",each -Amount(ClassRows(_,"MerchandiseCancellation"))/Scale,type number},
        {"SalesRows",each Table.RowCount(ClassRows(_,"MerchandiseSale")),Int64.Type},
        {"KnownCustomerSalesRows",each Table.RowCount(Table.SelectRows(_,each [CustomerAnalysisEligible])),Int64.Type},
        {"SalesInvoiceIDs",each InvoiceCount(ClassRows(_,"MerchandiseSale")),Int64.Type},
        {"ReviewRows",each Table.RowCount(Table.SelectRows(_,each [NeedsReview])),Int64.Type}}),
    Alternative = Table.SelectRows(Facts,each not [IsDuplicateBeyondFirst]),
    ScenarioRow = (t as table,label as text) as record => let s=ClassRows(t,"MerchandiseSale"),c=ClassRows(t,"MerchandiseCancellation") in
        [Scenario=label,SourceRows=Table.RowCount(t),RawAmountUnits4=Amount(t),RawSignedGBP=Amount(t)/Scale,
        SalesGBP=Amount(s)/Scale,CancellationGBP=-Amount(c)/Scale,BalanceGBP=(Amount(s)+Amount(c))/Scale,
        SalesInvoiceIDs=InvoiceCount(s),CustomerSalesIDs=List.Count(List.Distinct(List.RemoveNulls(s[CustomerID]))),
        ReviewRows=Table.RowCount(Table.SelectRows(t,each [NeedsReview]))],
    DuplicateScenario = Table.FromRecords({ScenarioRow(Facts,"Main: retain all rows"),ScenarioRow(Alternative,"Alternative: first exact record only")}),
    SpecialAmount = Amount(ClassRows(Facts,"SpecialOperation")),
    AdjustmentAmount = Amount(ClassRows(Facts,"ZeroPriceAdjustment")),
    UnresolvedAmount = Amount(ClassRows(Facts,"UnresolvedReview")),
    Bridge = Table.FromRecords({
        [Step="Raw signed transactions",ChangeGBP=OriginalAmount/Scale,BalanceGBP=OriginalAmount/Scale],
        [Step="Separate SpecialOperation",ChangeGBP=-SpecialAmount/Scale,BalanceGBP=(OriginalAmount-SpecialAmount)/Scale],
        [Step="Separate ZeroPriceAdjustment",ChangeGBP=-AdjustmentAmount/Scale,BalanceGBP=(OriginalAmount-SpecialAmount-AdjustmentAmount)/Scale],
        [Step="Separate UnresolvedReview",ChangeGBP=-UnresolvedAmount/Scale,BalanceGBP=(OriginalAmount-SpecialAmount-AdjustmentAmount-UnresolvedAmount)/Scale],
        [Step="Merchandise signed balance",ChangeGBP=0,BalanceGBP=Amount(ReportTransactions)/Scale]}),
    Quality = Table.FromRecords(List.Transform(FlagNames,(f)=>let t=Table.SelectRows(Facts,(r)=>Record.Field(r,f)) in
        [Flag=f,Rows=Table.RowCount(t),SourceRowExample=if Table.IsEmpty(t) then null else List.Min(t[SourceRow])])),
    SpecialOperations = Table.Group(ClassRows(Facts,"SpecialOperation"),{"StockCode","CodeCategory"},{
        {"Rows",each Table.RowCount(_),Int64.Type},{"SignedGBP",each Amount(_)/Scale,type number},
        {"SourceRowExample",each List.Min([SourceRow]),Int64.Type}}),
    Check = (name as text,expected as any,actual as any) as record => [Control=name,Expected=expected,Actual=actual,
        Status=if Value.Is(expected,type number) and Value.Is(actual,type number) then
            (if Number.Abs(expected-actual)<=0.0001 then "PASS" else "FAIL") else (if expected=actual then "PASS" else "FAIL"),Scope="Excel Power Query"],
    AlwaysChecks = {
        Check("Unique source RecordID",OriginalRows,List.Count(List.Distinct(Facts[RecordID]))),
        Check("Source rows preserved",OriginalRows,Table.RowCount(Facts)),
        Check("Exactly one class per row",OriginalRows,Sum0(Reconciliation[Rows])),
        Check("Class amounts reconcile (integer units)",OriginalAmount,Sum0(Reconciliation[AmountUnits4])),
        Check("Monthly source rows reconcile",OriginalRows,Sum0(Monthly[SourceRows])),
        Check("Monthly sales reconcile",Amount(Sales)/Scale,Sum0(Monthly[SalesGBP])),
        Check("Country sales reconcile",Amount(Sales)/Scale,Sum0(Country[SalesGBP])),
        Check("Product sales reconcile",Amount(Sales)/Scale,Sum0(Product[SalesGBP])),
        Check("Bridge closes to reporting scope",Amount(ReportTransactions)/Scale,List.Last(Bridge[BalanceGBP])),
        Check("Duplicate scenario rows removed",Table.RowCount(Table.SelectRows(Facts,each [IsDuplicateBeyondFirst])),Table.RowCount(Facts)-Table.RowCount(Alternative)),
        Check("Duplicate scenario amount delta (integer units)",Amount(Table.SelectRows(Facts,each [IsDuplicateBeyondFirst])),OriginalAmount-Amount(Alternative)),
        Check("No join inflation rows",OriginalRows,Table.RowCount(Facts)),
        Check("No join inflation amount",OriginalAmount,Amount(Facts))},
    BaselineChecks = if not Baseline then {} else {
        Check("Baseline rows",541909,OriginalRows),Check("Baseline invoice IDs",25900,InvoiceCount(Facts)),
        Check("Baseline missing CustomerID",135080,Table.RowCount(Table.SelectRows(Facts,each [HasMissingCustomer]))),
        Check("Baseline missing Description",1454,Table.RowCount(Table.SelectRows(Facts,each [Description_Raw]=null))),
        Check("Baseline duplicate groups",4879,Table.RowCount(Table.SelectRows(ExactGroups,each [ExactGroupRows]>1))),
        Check("Baseline duplicate members",10147,Table.RowCount(Table.SelectRows(Facts,each [IsExactDuplicateMember]))),
        Check("Baseline duplicate extras",5268,Table.RowCount(Table.SelectRows(Facts,each [IsDuplicateBeyondFirst]))),
        Check("Baseline C-prefix rows",9288,Table.RowCount(Table.SelectRows(Facts,each [IsCancellationCode]))),
        Check("Baseline non-C negative",1336,Table.RowCount(Table.SelectRows(Facts,each [IsNonCNegative]))),
        Check("Baseline zero prices",2515,Table.RowCount(Table.SelectRows(Facts,each [IsZeroPrice]))),
        Check("Baseline negative prices",2,Table.RowCount(Table.SelectRows(Facts,each [IsNegativePrice]))),
        Check("Baseline multidescription codes",650,Table.RowCount(Table.SelectRows(ProductDictionary,each [DescriptionVariantCount]>1))),
        Check("Baseline outer whitespace",113452,Table.RowCount(Table.SelectRows(Facts,each [HasOuterWhitespace]))),
        Check("Baseline raw amount units",97477479340,OriginalAmount)},
    Controls = Table.FromRecords(List.Combine({AlwaysChecks,BaselineChecks,{
        Check("Sales invoice cube partition",InvoiceCount(Sales),Sum0(DashboardData[SalesInvoiceIDs]))}})),
    BeforeAfter = Table.SelectColumns(Table.SelectRows(Facts,each List.Contains({12,487,541,624,1445,9304,199126,199146,299985,420393,540423,540424},[SourceRow])),
        {"RecordID","SourceRow","InvoiceNo_Raw","StockCode_Raw","Description_Raw","DescriptionTrim","CustomerID_Raw","CustomerID","TransactionClass","ReviewReasons","LineAmount"}),
    Result = [Facts=Facts,ProductDictionary=ProductDictionary,CodeRules=CodeRules,KPI=KPI,Monthly=Monthly,
        Country=Country,Product=Product,Customers=Customers,DashboardData=DashboardData,Reconciliation=Reconciliation,Bridge=Bridge,
        Quality=Quality,ReviewRecords=ReviewRecords,ReviewExamples=Table.FirstN(Table.Sort(ReviewRecords,{{"SourceRow",Order.Ascending}}),200),
        DuplicateScenario=DuplicateScenario,Controls=Controls,BeforeAfter=BeforeAfter,
        ReportTransactions=ReportTransactions,CustomerTransactions=CustomerTransactions,SpecialOperations=SpecialOperations]
in
    Result
