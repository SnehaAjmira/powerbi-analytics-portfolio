// Power Query (M). Create a text parameter named DataFolder that points to the data folder
// (ending with a backslash), e.g. C:\PowerBI\lab\

// ===== RAW_2023
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
    RAW_2023_Table = Source{[Item="RAW_2023",Kind="Table"]}[Data]
in
    RAW_2023_Table

// ===== RAW_2024
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
    RAW_2024_Table = Source{[Item="RAW_2024",Kind="Table"]}[Data]
in
    RAW_2024_Table

// ===== RAW_2025
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
    RAW_2025_Table = Source{[Item="RAW_2025",Kind="Table"]}[Data]
in
    RAW_2025_Table

// ===== Fact Table (Services)
let
    Source = Table.Combine({RAW_2023, RAW_2024, RAW_2025, RAW_2026, Canada_2026}),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Submission_Release_Date", type date}, {"Submission_Date", type date}, {"Estimated_Completion_Date", type date}, {"Completion_Date", type date}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Type", each "Lab Serivce"),
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom",{"Test_Point"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns",{{"Submission_Date", Order.Descending}}),
    #"Filtered Rows" = Table.SelectRows(#"Sorted Rows", each true)
in
    #"Filtered Rows"

// ===== RAW_2026
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
    RAW_2026_Table = Source{[Item="RAW_2026",Kind="Table"]}[Data]
in
    RAW_2026_Table

// ===== Therapeutics_Table
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
    Therapeutics_Table_Table = Source{[Item="Therapeutics_Table",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Therapeutics_Table_Table,{{"Sample name", type text}, {"Therapy Type", type text}, {"Umbrella", type text}})
in
    #"Changed Type"

// ===== Client Orders
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Product Test.xlsx"), null, true),
    #"Client Orders_Sheet" = Source{[Item="Client Orders",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Client Orders_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Order_Number", type text}, {"Client_Facility", type text}, {"Account_Number", type text}, {"Web_Order_Number", Int64.Type}, {"PO_Number", type text}, {"Order_Date", type date}, {"Status", type text}, {"Order_Placed_By", type text}, {"Phone", type text}, {"Fax", type text}, {"Email", type text}, {"Promo_Code", type text}, {"Ship_To_Address", type text}, {"Bill_To_Address", type text}, {"Line_Count", Int64.Type}, {"Line_Total", Int64.Type}, {"Shipping_And_Handling", type number}, {"Tax", Int64.Type}, {"Discount", Int64.Type}, {"Order_Total", type number}})
in
    #"Changed Type"

// ===== Order Details
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Product Test.xlsx"), null, true),
    #"Order Details_Sheet" = Source{[Item="Order Details",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Order Details_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Order_Number", type text}, {"Item", type text}, {"Packaging", type text}, {"Ordered", Int64.Type}, {"Shipped", Int64.Type}, {"Unit_Price", Int64.Type}, {"Ext_Price", Int64.Type}}),
    #"Merged Queries" = Table.FuzzyNestedJoin(#"Changed Type", {"Order_Number"}, #"Client Orders", {"Order_Number"}, "Client Orders", JoinKind.LeftOuter, [IgnoreCase=true, IgnoreSpace=true])
in
    #"Merged Queries"

// ===== Fact Table (Products)
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Product Test.xlsx"), null, true),
    #"Order Details_Sheet" = Source{[Item="Order Details",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Order Details_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Order_Number", type text}, {"Item", type text}, {"Packaging", type text}, {"Ordered", Int64.Type}, {"Shipped", Int64.Type}, {"Unit_Price", Int64.Type}, {"Ext_Price", Int64.Type}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"Order_Number"}, #"Client Orders", {"Order_Number"}, "Client Orders", JoinKind.LeftOuter),
    #"Renamed Columns" = Table.RenameColumns(#"Merged Queries",{{"Client Orders", "Client Orders Header"}}),
    #"Expanded Client Orders Header" = Table.ExpandTableColumn(#"Renamed Columns", "Client Orders Header", {"Client_Facility", "Order_Date", "Line_Count", "Line_Total", "Shipping_And_Handling", "Tax", "Discount", "Order_Total"}, {"Client_Facility", "Order_Date", "Line_Count", "Line_Total", "Shipping_And_Handling", "Tax", "Discount", "Order_Total"})
in
    #"Expanded Client Orders Header"

// ===== Consulting Activities
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Consulting Activities.xlsx"), null, true),
    #"Consulting Activities_Sheet" = Source{[Item="Consulting Activities",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Consulting Activities_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Activity_Id", type text}, {"Client_Facility_Name", type text}, {"Project_Name", type text}, {"PO_Number", type text}, {"Quote_Number", type text}, {"Invoice_Number", Int64.Type}, {"Billing_Codes", type text}, {"Billing_Domain", type text}, {"Estimated_Completion_Date", type date}, {"Completed", type text}, {"Completion_Date", type date}, {"Invoiced", type text}, {"Invoice_Submitted_Date", type date}, {"Price", type number}, {"Description", type text}, {"Expenses_Price", type number}, {"Expenses_Description", type text}, {"Entered_By", type text}})
in
    #"Changed Type"

// ===== Client_Facility
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
  Navigation = Source{[Item = "Table13", Kind = "Table"]}[Data]
in
    Navigation

// ===== Canada_2026
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw Data Canada.xlsx"), null, true),
    Table1_Table = Source{[Item="Table1",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table1_Table,{{"Submission_Number", type text}, {"Submission_Date", type date}, {"Submission_Status", type text}, {"Client_Facility", type text}, {"Select", type any}, {"Facility_Type", type text}, {"Sample_Name", type text}, {"Dosage_Form", type text}, {"Lot_Number", type text}, {"BUD_Test_Points", type any}, {"Test_Point", type any}, {"Test_Name", type text}, {"Analyte", type text}, {"Test_Id", type text}, {"Submission_Release_Date", type date}, {"Estimated_Completion_Date", type date}, {"Completion_Date", type date}, {"Rush", type text}, {"Status", type text}, {"Analyst", type text}, {"Controlled_Substance_Class", type text}, {"Hazardous_Drug", type text}, {"Price", type number}, {"Test_Facility", type text}})
in
    #"Changed Type"

// ===== Canada Client Facility
let
  Source = Excel.Workbook(File.Contents(DataFolder & "Raw Data Canada.xlsx"), null, true),
  #"Navigation 1" = Source{[Item = "Table2", Kind = "Table"]}[Data],
  #"Changed column type" = Table.TransformColumnTypes(#"Navigation 1", {{"ClientFacilityName", type text}, {"BillTo_City", type text}, {"BillTo_StateProvince", type text}, {"BillTo_Country", type text}, {"Active Client", type text}, {"Test Facility", type text}})
in
  #"Changed column type"

// ===== All Clients
let
    Source = Table.Combine({#"Canada Client Facility", Client_Facility})
in
    Source

// ===== _Measures
let
  Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i44FAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t])
in
  Source
