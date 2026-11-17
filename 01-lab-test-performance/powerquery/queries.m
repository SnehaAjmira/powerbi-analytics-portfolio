// Power Query (M). Create a text parameter named DataFolder that points to the data folder
// (ending with a backslash), e.g. C:\PowerBI\lab\

// ===== Tests Complete
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Tests Completed.xlsx"), null, true),
    Tests_Sheet = Source{[Item="Tests",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Tests_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Submission_Number", type text}, {"Submission_Date", type date}, {"Submission_Status", type text}, {"Client_Facility", type text}, {"Select", type text}, {"Facility_Type", type text}, {"Sample_Name", type text}, {"Dosage_Form", type text}, {"Lot_Number", type text}, {"BUD_Test_Points", type text}, {"Test_Point", type text}, {"Test_Name", type text}, {"Analyte", type text}, {"Test_Id", type text}, {"Submission_Release_Date", type date}, {"Estimated_Completion_Date", type date}, {"Completion_Date", type date}, {"Rush", type text}, {"Status", type text}, {"Analyst", type text}, {"Controlled_Substance_Class", type text}, {"Hazardous_Drug", type text}, {"Price", type number}})
in
    #"Changed Type"

// ===== Tests Release
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Submission Release.xlsx"), null, true),
    Tests_Sheet = Source{[Item="Tests",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Tests_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Submission_Number", type text}, {"Submission_Date", type date}, {"Submission_Status", type text}, {"Client_Facility", type text}, {"Select", type text}, {"Facility_Type", type text}, {"Sample_Name", type text}, {"Dosage_Form", type text}, {"Lot_Number", type text}, {"BUD_Test_Points", type text}, {"Test_Point", type text}, {"Test_Name", type text}, {"Analyte", type text}, {"Test_Id", type text}, {"Submission_Release_Date", type date}, {"Estimated_Completion_Date", type date}, {"Completion_Date", type date}, {"Rush", type text}, {"Status", type text}, {"Analyst", type text}, {"Controlled_Substance_Class", type text}, {"Hazardous_Drug", type text}, {"Price", Int64.Type}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Submission_Date", Order.Descending}})
in
    #"Sorted Rows"

// ===== Departments
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
    Table2_Table = Source{[Item="Table2",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Table2_Table,{{"Department", type text}, {"Sub - Department", type text}, {"Test_Name", type text}})
in
    #"Changed Type"

// ===== _Measures
let
  Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i44FAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t]),
    #"Removed Columns" = Table.RemoveColumns(Source,{"Column1"})
in
  #"Removed Columns"

// ===== Quality Data
let
  Source = Excel.Workbook(File.Contents(DataFolder & "Quality Data.xlsx"), null, true),
  #"Navigation 1" = Source{[Item = "Quality Data", Kind = "Sheet"]}[Data],
  #"Promoted headers" = Table.PromoteHeaders(#"Navigation 1", [PromoteAllScalars = true]),
  #"Changed column type" = Table.TransformColumnTypes(#"Promoted headers", {{"Submission", type text}, {"Test_Name", type text}, {"Department", type text}, {"Dosage_Form", type text}, {"Analyst", type text}, {"Client_Name", type text}, {"Facility_Type", type text}, {"Time_In_Sample_Prep", type number}, {"Time_In_Sample_Analysis", type number}, {"Time_In_Data_Processing", type number}, {"Time_In_Data_Review", type number}, {"Time_In_On_Hold", type number}, {"Time_In_Approved", Int64.Type}, {"Time_In_Qa_Review", type number}, {"Time_In_Qa_Hold", type number}, {"Time_In_Qa_Approved", Int64.Type}, {"Time_To_Complete", type number}, {"Completed", type datetime}, {"Approved_By", type text}})
in
  #"Changed column type"

// ===== OosByClient
let
  Source = Excel.Workbook(File.Contents(DataFolder & "OosByClient.xlsx"), null, true),
  #"Navigation 1" = Source{[Item = "OosByClient", Kind = "Sheet"]}[Data],
  #"Promoted headers" = Table.PromoteHeaders(#"Navigation 1", [PromoteAllScalars = true]),
  #"Added custom" = Table.AddColumn(#"Promoted headers", "Analyte", each try
  let
    s = Text.Replace(Text.Trim(Text.From([TestType])), "  ", " "),
    a = if Text.Contains(s, ") (") then Text.AfterDelimiter(s, ") (")
        else if Text.Contains(s, "Samples (") then Text.AfterDelimiter(s, "Samples (")
        else null
  in
    if a = null then null else Text.Start(a, Text.Length(a) - 1)
otherwise "CHECK")
in
  #"Added custom"

// ===== OosByTechnician
let
  Source = Excel.Workbook(File.Contents(DataFolder & "OosByTechnician.xlsx"), null, true),
  #"Navigation 1" = Source{[Item = "OosByTechnician", Kind = "Sheet"]}[Data],
  #"Promoted headers" = Table.PromoteHeaders(#"Navigation 1", [PromoteAllScalars = true]),
  #"Changed column type" = Table.TransformColumnTypes(#"Promoted headers", {{"Date", type date}, {"Analyst", type text}, {"Department", type text}, {"TestType", type text}, {"TestsPerformed", Int64.Type}, {"NumberOfOOS", Int64.Type}})
in
  #"Changed column type"
