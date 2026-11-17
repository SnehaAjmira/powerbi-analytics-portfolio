// =====================================================================================
// Project Saguaro (fictional battery cell gigafactory): Power Query for all sources
//
// One block per query. In Power BI Desktop: Home > Transform data > New Source >
// Blank Query > Advanced Editor, paste the code between "let" and "in ...", and rename
// the query to the name in the header comment. Table names must match exactly because
// the DAX measures use them.
//
// Sources
//   SQL Server          dim_discipline, dim_area, dim_company, dim_system, dim_subsystem,
//                       project_info, company_erp_xref, fact_document_revisions, fact_rfi,
//                       fact_submittals
//   ERP export (Excel)  fact_procurement        (vendor numbers mapped to company IDs)
//   P6 export (Excel)   dim_cwp, fact_progress_weekly
//   Completions (CSV)   fact_itr (folder of extracts), fact_tag_register, fact_punch
//   HSE API (JSON)      fact_safety_weekly
//   SharePoint (Excel)  fact_ncr
// =====================================================================================


// ===== PARAMETERS (Home > Manage Parameters > New Parameter, Type = Text)
// SqlServer    e.g. localhost\SQLEXPRESS   (or the project's SQL Server name)
// SqlDatabase  SaguaroPIIM
// SourceRoot   folder that holds the other source folders, ending with a backslash,
//              e.g. C:\PowerBI\gigafactory\


// =====================================================================================
// 1. SQL SERVER (engineering database). Navigation steps fold back to SQL Server.
// =====================================================================================

// ----- dim_discipline
let
    Source = Sql.Database(SqlServer, SqlDatabase),
    Data = Source{[Schema = "ref", Item = "discipline"]}[Data]
in
    Data

// ----- dim_area
let
    Source = Sql.Database(SqlServer, SqlDatabase),
    Data = Source{[Schema = "ref", Item = "area"]}[Data]
in
    Data

// ----- dim_company
let
    Source = Sql.Database(SqlServer, SqlDatabase),
    Data = Source{[Schema = "ref", Item = "company"]}[Data]
in
    Data

// ----- company_erp_xref   (helper: right-click > uncheck "Enable load")
let
    Source = Sql.Database(SqlServer, SqlDatabase),
    Data = Source{[Schema = "ref", Item = "company_erp_xref"]}[Data]
in
    Data

// ----- dim_system
let
    Source = Sql.Database(SqlServer, SqlDatabase),
    Data = Source{[Schema = "ref", Item = "facility_system"]}[Data]
in
    Data

// ----- dim_subsystem
let
    Source = Sql.Database(SqlServer, SqlDatabase),
    Data = Source{[Schema = "ref", Item = "subsystem"]}[Data]
in
    Data

// ----- project_info
let
    Source = Sql.Database(SqlServer, SqlDatabase),
    Data = Source{[Schema = "ref", Item = "project_info"]}[Data]
in
    Data

// ----- fact_document_revisions
let
    Source = Sql.Database(SqlServer, SqlDatabase),
    Data = Source{[Schema = "eng", Item = "document_revision"]}[Data]
in
    Data

// ----- fact_rfi
let
    Source = Sql.Database(SqlServer, SqlDatabase),
    Data = Source{[Schema = "eng", Item = "rfi"]}[Data]
in
    Data

// ----- fact_submittals
let
    Source = Sql.Database(SqlServer, SqlDatabase),
    Data = Source{[Schema = "eng", Item = "submittal"]}[Data]
in
    Data


// =====================================================================================
// 2. ERP PURCHASE ORDER EXPORT (Excel). ERP column names and vendor numbers are
//    translated to the project model.
// =====================================================================================

// ----- fact_procurement
let
    Source = Excel.Workbook(File.Contents(SourceRoot & "erp_procurement\PO_Lines_Export.xlsx"), null, true),
    Sheet = Source{[Item = "PO_Lines", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"PO Number", type text}, {"Item", Int64.Type}, {"Vendor No", type text}, {"Material Description", type text},
        {"UoM", type text}, {"Order Qty", Int64.Type}, {"Net Value (USD)", type number}, {"Long Lead Flag", type text},
        {"PO Date", type date}, {"Required On Site", type date}, {"Vendor Promised Date", type date},
        {"Expedited Forecast Date", type date}, {"Goods Receipt Date", type date}, {"Discipline", type text},
        {"Delivery Area", type text}, {"Expediting Status", type text}}),
    Renamed = Table.RenameColumns(Typed, {
        {"PO Number", "PO_No"}, {"Vendor No", "Vendor_No"}, {"Material Description", "Commodity"}, {"UoM", "Unit"},
        {"Order Qty", "Quantity"}, {"Net Value (USD)", "Line_Value_USD"}, {"PO Date", "Order_Date"},
        {"Required On Site", "ROS_Date"}, {"Vendor Promised Date", "Promised_Date"},
        {"Expedited Forecast Date", "Forecast_Delivery_Date"}, {"Goods Receipt Date", "Actual_Delivery_Date"},
        {"Discipline", "Discipline_Code"}, {"Delivery Area", "Area_Code"}, {"Expediting Status", "Expediting_Status"}}),
    // ERP line items count in steps of 10 (10, 20, 30...)
    AddLine = Table.AddColumn(Renamed, "PO_Line", each Number.IntegerDivide([Item], 10), Int64.Type),
    AddLongLead = Table.AddColumn(AddLine, "Long_Lead", each if [Long Lead Flag] = "X" then "Y" else "N", type text),
    // map ERP vendor numbers to project company IDs through the SQL cross-reference table
    Merged = Table.NestedJoin(AddLongLead, {"Vendor_No"}, company_erp_xref, {"ERP_Vendor_No"}, "xref", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, "xref", {"Company_ID"}, {"Supplier_ID"}),
    Result = Table.RemoveColumns(Expanded, {"Item", "Long Lead Flag"})
in
    Result


// =====================================================================================
// 3. P6 SCHEDULE / PROGRESS EXPORT (Excel)
// =====================================================================================

// ----- dim_cwp
let
    Source = Excel.Workbook(File.Contents(SourceRoot & "p6_schedule\Construction_Progress_Export.xlsx"), null, true),
    Sheet = Source{[Item = "WBS_CWP", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Renamed = Table.RenameColumns(Promoted, {
        {"CWP ID", "CWP_ID"}, {"Area", "Area_Code"}, {"Discipline", "Discipline_Code"}, {"Contractor", "Company_ID"},
        {"Planned Start", "Planned_Start"}, {"Planned Finish", "Planned_Finish"}, {"Budgeted Units (hrs)", "Budget_Hours"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"CWP_ID", type text}, {"Area_Code", type text}, {"Discipline_Code", type text}, {"Company_ID", type text},
        {"Planned_Start", type date}, {"Planned_Finish", type date}, {"Budget_Hours", type number}})
in
    Typed

// ----- fact_progress_weekly
let
    Source = Excel.Workbook(File.Contents(SourceRoot & "p6_schedule\Construction_Progress_Export.xlsx"), null, true),
    Sheet = Source{[Item = "Weekly_Progress", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Renamed = Table.RenameColumns(Promoted, {
        {"Week Ending", "Week_Ending"}, {"CWP ID", "CWP_ID"}, {"Planned Units", "Planned_Hours"},
        {"Earned Units", "Earned_Hours"}, {"Actual Units", "Actual_Hours"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Week_Ending", type date}, {"CWP_ID", type text}, {"Planned_Hours", type number},
        {"Earned_Hours", type number}, {"Actual_Hours", type number}})
in
    Typed


// =====================================================================================
// 4. COMPLETIONS SYSTEM (CSV extracts). New monthly files dropped in the folder are
//    picked up on the next refresh.
// =====================================================================================

// ----- fact_itr
let
    Source = Folder.Files(SourceRoot & "completions_system\itr_extracts"),
    Extracts = Table.SelectRows(Source, each Text.StartsWith([Name], "itr_extract_") and [Extension] = ".csv"),
    Parsed = Table.AddColumn(Extracts, "Data", each
        Table.PromoteHeaders(Csv.Document([Content], [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
            [PromoteAllScalars = true])),
    Kept = Table.SelectColumns(Parsed, {"Name", "Data"}),
    Expanded = Table.ExpandTableColumn(Kept, "Data",
        {"ITR_ID", "Tag_No", "Subsystem_Code", "Discipline_Code", "ITR_Form", "ITR_Description", "Stage",
         "Planned_Date", "Completed_Date", "Status", "Completed_By_ID"}),
    Renamed = Table.RenameColumns(Expanded, {{"Name", "Source_File"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Planned_Date", type date}, {"Completed_Date", type date}, {"ITR_ID", type text}, {"Tag_No", type text},
        {"Subsystem_Code", type text}, {"Discipline_Code", type text}, {"ITR_Form", type text},
        {"ITR_Description", type text}, {"Stage", type text}, {"Status", type text}, {"Completed_By_ID", type text}}, "en-US"),
    // extracts can overlap; keep one row per ITR
    Deduped = Table.Distinct(Typed, {"ITR_ID"})
in
    Deduped

// ----- fact_tag_register
let
    Source = Csv.Document(File.Contents(SourceRoot & "completions_system\tag_register.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"Tag_No", type text}, {"Tag_Type", type text}, {"Discipline_Code", type text}, {"System_Code", type text},
        {"Subsystem_Code", type text}, {"Area_Code", type text}, {"Install_Status", type text},
        {"Install_Date", type date}}, "en-US")
in
    Typed

// ----- fact_punch
let
    Source = Csv.Document(File.Contents(SourceRoot & "completions_system\punch_list.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"Punch_No", type text}, {"Subsystem_Code", type text}, {"Tag_No", type text}, {"Area_Code", type text},
        {"Discipline_Code", type text}, {"Category", type text}, {"Punch_Type", type text}, {"Date_Raised", type date},
        {"Target_Date", type date}, {"Date_Cleared", type date}, {"Raised_By_ID", type text},
        {"Responsible_ID", type text}, {"Status", type text}}, "en-US")
in
    Typed


// =====================================================================================
// 5. HSE REST API (JSON, OData style)
// =====================================================================================

// ----- fact_safety_weekly
let
    // Offline copy of the API response. For the live API replace the next line with:
    //   Source = Json.Document(Web.Contents("https://hse.example.com",
    //       [RelativePath = "api/v1/WeeklyStatistics", Headers = [Accept = "application/json"]])),
    // Keep the base URL fixed and use RelativePath so scheduled refresh works in the Power BI service.
    Source = Json.Document(File.Contents(SourceRoot & "hse_api\safety_weekly.json")),
    Rows = Table.FromRecords(Source[value]),
    Renamed = Table.RenameColumns(Rows, {
        {"weekEnding", "Week_Ending"}, {"companyId", "Company_ID"}, {"headcount", "Headcount"},
        {"manhours", "Manhours"}, {"recordables", "Recordables"}, {"lostTimeInjuries", "Lost_Time_Injuries"},
        {"firstAids", "First_Aids"}, {"nearMisses", "Near_Misses"}, {"observations", "Safety_Observations"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Week_Ending", type date}, {"Company_ID", type text}, {"Headcount", Int64.Type}, {"Manhours", Int64.Type},
        {"Recordables", Int64.Type}, {"Lost_Time_Injuries", Int64.Type}, {"First_Aids", Int64.Type},
        {"Near_Misses", Int64.Type}, {"Safety_Observations", Int64.Type}}, "en-US")
in
    Typed


// =====================================================================================
// 6. SHAREPOINT LIST (NCR log)
// =====================================================================================

// ----- fact_ncr
let
    // Offline copy exported from the list. For the live list replace the next three lines with:
    //   Source = SharePoint.Tables("https://<tenant>.sharepoint.com/sites/ProjectSaguaro", [Implementation = "2.0"]),
    //   Promoted = Source{[Title = "NCR Log"]}[Items],
    Source = Excel.Workbook(File.Contents(SourceRoot & "quality_sharepoint\NCR_Log.xlsx"), null, true),
    Sheet = Source{[Item = "NCR Log", Kind = "Sheet"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    // drop SharePoint system columns
    NoMeta = Table.RemoveColumns(Promoted, {"ID", "Created", "Modified", "Item Type", "Path"}, MissingField.Ignore),
    Renamed = Table.RenameColumns(NoMeta, {
        {"Title", "NCR_No"}, {"Discipline", "Discipline_Code"}, {"Contractor", "Company_ID"}, {"Area", "Area_Code"},
        {"Root Cause", "Root_Cause"}, {"Date Opened", "Date_Opened"}, {"Date Closed", "Date_Closed"},
        {"NCR Status", "Status"}}),
    Typed = Table.TransformColumnTypes(Renamed, {
        {"NCR_No", type text}, {"Discipline_Code", type text}, {"Company_ID", type text}, {"Area_Code", type text},
        {"Category", type text}, {"Severity", type text}, {"Root_Cause", type text}, {"Date_Opened", type date},
        {"Date_Closed", type date}, {"Status", type text}})
in
    Typed
