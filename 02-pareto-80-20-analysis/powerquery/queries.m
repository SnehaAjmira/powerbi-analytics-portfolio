// Power Query (M). Create a text parameter named DataFolder that points to the data folder
// (ending with a backslash), e.g. C:\PowerBI\lab\

// ===== RAW_2023
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
    RAW_2023_Table = Source{[Item="RAW_2023",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(RAW_2023_Table,{{"Submission_Number", type text}, {"Submission_Date", type date}, {"Submission_Status", type text}, {"Client_Facility", type text}, {"Select", type any}, {"Facility_Type", type text}, {"Sample_Name", type text}, {"Dosage_Form", type text}, {"Lot_Number", type any}, {"Test_Point", type text}, {"Test_Name", type text}, {"Analyte", type text}, {"Test_Id", type text}, {"Submission_Release_Date", type date}, {"Estimated_Completion_Date", type date}, {"Completion_Date", type date}, {"Rush", type text}, {"Status", type text}, {"Analyst", type text}, {"Controlled_Substance_Class", type text}, {"Hazardous_Drug", type text}, {"Price", type number} })
in
    #"Changed Type"

// ===== RAW_2024
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
    RAW_2024_Table = Source{[Item="RAW_2024",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(RAW_2024_Table,{{"Submission_Number", type text}, {"Submission_Date", type date}, {"Submission_Status", type text}, {"Client_Facility", type text}, {"Select", type text}, {"Facility_Type", type text}, {"Sample_Name", type text}, {"Dosage_Form", type text}, {"Lot_Number", type any}, {"Test_Point", type text}, {"Test_Name", type text}, {"Analyte", type text}, {"Test_Id", type text}, {"Submission_Release_Date", type date}, {"Estimated_Completion_Date", type date}, {"Completion_Date", type date}, {"Rush", type text}, {"Status", type text}, {"Analyst", type text}, {"Controlled_Substance_Class", type text}, {"Hazardous_Drug", type text}, {"Price", type number}})
in
    #"Changed Type"

// ===== RAW_2025
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
    RAW_2025_Table = Source{[Item="RAW_2025",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(RAW_2025_Table,{{"Submission_Number", type text}, {"Submission_Date", type date}, {"Submission_Status", type text}, {"Client_Facility", type text}, {"Select", type text}, {"Facility_Type", type text}, {"Sample_Name", type text}, {"Dosage_Form", type text}, {"Lot_Number", type text}, {"Test_Point", type text}, {"Test_Name", type text}, {"Analyte", type text}, {"Test_Id", type text}, {"Submission_Release_Date", type date}, {"Estimated_Completion_Date", type date}, {"Completion_Date", type date}, {"Rush", type text}, {"Status", type text}, {"Analyst", type text}, {"Controlled_Substance_Class", type text}, {"Hazardous_Drug", type text}, {"Price", type number}})
in
    #"Changed Type"

// ===== AllData
let
    Source = Table.Combine({RAW_2025, RAW_2023, RAW_2024, RAW_2026}),
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Price", Int64.Type}}),
    #"Sorted Rows" = Table.Sort(#"Changed Type",{{"Submission_Date", Order.Descending}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Sorted Rows",{{"Submission_Release_Date", type date}, {"Estimated_Completion_Date", type date}, {"Completion_Date", type date},  {"Price", type number}})
in
    #"Changed Type1"

// ===== RAW_2026
let
    Source = Excel.Workbook(File.Contents(DataFolder & "Raw DATA.xlsx"), null, true),
    RAW_2026_Table = Source{[Item="RAW_2026",Kind="Table"]}[Data]
in
    RAW_2026_Table
