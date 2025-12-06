let
    // --- 1. HISTORICAL (MASTER) DATA PREPARATION ---
    // Get historical (Master) File
    Master_Source = Excel.Workbook(File.Contents("C:\Users\doria\PowerQuery\master_file_template_DT_TEXT.xlsx"), null, true),
    Sample_Master_Sheet = Master_Source{[Item="Employee_Master",Kind="Sheet"]}[Data],
    HistoricalSCD = Table.PromoteHeaders(Sample_Master_Sheet, [PromoteAllScalars=true]),
    // Note: Effective_Start_Date and Effective_End_Date columns are retained but not used for the SCD Type 2 logic below.
    Master_Changed_Type = Table.TransformColumnTypes(HistoricalSCD,{{"Employee ID", type number},{"Data_End_Date", type date}, 
    {"Data_Start_Date", type date},{"Hire Date", type date}}),
    //Buffer_Sorted = Table.Buffer(Master_Changed_Type),
    Sorted_Master = Table.Sort(Master_Changed_Type,{{"Employee ID", Order.Ascending}}),

    // --- 2. CURRENT (HR REPORT) DATA PREPARATION ---
    HR_Source = Excel.Workbook(File.Contents("C:\Users\doria\PowerQuery\HR Standard Report.xlsx"), null, true),
    HR_Standard_Sheet = HR_Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    H_Removed_Top_Rows = Table.Skip(HR_Standard_Sheet,2),
    H_Promoted_Headers1 = Table.PromoteHeaders(H_Removed_Top_Rows, [PromoteAllScalars=true]),
    H_Filtered_Rows = Table.SelectRows(H_Promoted_Headers1, 
        each [#"Worker Type - Text"] = "Employee" 
            and Text.Contains([Cost Center], "Sales") 
            and [Country] = "United States of America"
    ),
    //H_Rename_EID =   Table.RenameColumns(H_Filtered_Rows, {{"Employee ID","HR Employee ID"}}),
    H_Remove_Cols = Table.RemoveColumns(H_Filtered_Rows  ,{"Worker Type - Text", "Employee Type", "Fixed-Term End Date", "Contingent Worker Type", "Exempt / Non-Exempt", "Contract End Date", "Time Type", "Source", "First Name", "Last Name", "Position ID", "Job Code", "Job Profile", "Job Title", "Company", "Supervisory Organization", "Supervisory Organization - Hierarchy", "Functional Organizational Hierarchy", "Functional Org Level 1", "Functional Org Level 2", "Functional Org Level 3", "Functional Org Level 4", "Functional Org Level 5", "Cost Center", "Cost Center Hierarchy", "Job Family", "Work Arrangement", "Current Work Arrangement", "Location", "Scheduled Weekly Hours", "Country", "City", "State", "Continuous Service Date", "Company Service Date", "Termination Date", "Length of Service", "Time in Job Profile", "Work Shift", "Is Manager", "Management Level", "HR Partner", "User Name", "Prior Work Email", "Number of Direct Reports", "Management Chain - Level 09", "Management Chain - Level 08", "Management Chain - Level 07", "Management Chain - Level 06", "Management Chain - Level 05", "Management Chain - Level 04", "Management Chain - Level 03", "Management Chain - Level 02", "Management Chain - Level 01"}),
    H_EID_Num = Table.TransformColumnTypes(H_Remove_Cols,{{"Employee ID", type number}}),
    //H_Buffer_Table = Table.Buffer(H_EID_Num),
    H_Sorted_Rows = Table.Sort(H_EID_Num,{{"Employee ID", Order.Ascending}}),
    
    // Define the key column
    //KeyColumn = "Employee ID",
    // --- 3. SCD TYPE 2 IMPLEMENTATION LOGIC ---
    // Configuration Variables
    // Define the columns used for comparison (the 'data' columns)
   // ComparisonColumns = {"Worker", "Business Title", "Current Work Email", "Worker's Manager(s)"},
    InactiveMaster_final = Table.SelectRows(Sorted_Master, each [Data_End_Date] <> null and [Employee ID] <> null),
    // Get only the currently ACTIVE records from the Master Data
    ActiveMaster = Table.SelectRows(Sorted_Master, each [Data_End_Date] = null and [Employee ID] <> null),

    // Perform a Right Anti-Join to find NEW records in the Current Snapshot (H_Sorted_Rows) 
    // that DO NOT exist in the Master (ActiveMaster).
    NewInserts = Table.NestedJoin(
        ActiveMaster, // Left table (Master)
        "Employee ID",
        H_Sorted_Rows, // Right table (Current)
       "Employee ID",
       "NTC",
        JoinKind.RightAnti
    ),

    NewInsertsExpanded = Table.ExpandTableColumn(
        NewInserts,
        "NTC",
        {"Employee ID","Worker", "Hire Date",  "Business Title", "Current Work Email", "Worker's Manager(s)"},
         {"HR Employee ID","HR Worker", "HR Hire Date",  "HR Business Title", "HR Current Work Email", "HR Workers Managers"}
            ),
   Removed_Columns = Table.RemoveColumns(NewInsertsExpanded,
  {"Employee ID","Worker", "Hire Date",  "Business Title", "Current Work Email", "Worker's Manager(s)",
  "Data_Start_Date","Note"}
)//"Effective_End_Date",
,

    RenameNewInsertColumns =  Table.RenameColumns(Removed_Columns,
       {{"HR Employee ID","Employee ID"},{"HR Worker","Worker"},{"HR Hire Date","Hire Date"},
        {"HR Business Title","Business Title"},
       {"HR Current Work Email","Current Work Email"},{ "HR Workers Managers","Worker's Manager(s)"}}), 


    UpdateDataStartDate = Table.AddColumn(RenameNewInsertColumns, "Data_Start_Date", each DateTime.Date(DateTime.LocalNow()), type date),
    
    //UpdateEndDate = Table.AddColumn(UpdateStartDate, "Effective_End_Date", each DateTime.Date(DateTime.LocalNow()), type date),

     NewInserts_Final = Table.AddColumn(UpdateDataStartDate, "Note", each "Hired", type text ),

// LeftOuter	All rows from the left table, plus matching rows from the right table; unmatched right-side values become null
// RightOuter	All rows from the right table, plus matching rows from the left table; unmatched left-side values become null
// FullOuter	All rows from both tables; unmatched rows on either side show nulls for missing fields
// Inner	Only rows where the key exists in both tables; no unmatched rows
// LeftAnti	Rows from the left table that have no match in the right table
// RightAnti	Rows from the right table that have no match in the left table

    // Perform a Left Outer Join and Expand to align historical and current records for comparison.
    JoinedRecords = Table.NestedJoin(
         ActiveMaster, // Left table (Master)
        {"Employee ID"},
        H_Sorted_Rows, // Right table (Current)
       {"Employee ID"},
        "CurrentRecord",
        JoinKind.Inner
    ),
    
    // Expand the current record columns and rename them to avoid conflict.
    // The comparison columns are {"Worker", "Business Title", "Current Work Email", "Worker's Manager(s)"}
    ExpandedRecords = Table.ExpandTableColumn(
        JoinedRecords, 
        "CurrentRecord", 
        {"Worker",  "Business Title", "Current Work Email", "Worker's Manager(s)"},
         {"HR Worker",  "HR Business Title", "HR Current Work Email", "HR Workers Managers"}
    ),

    // Identify records that have CHANGED and need to be expired (Type 2 change)
    // The logic is now a direct comparison between the historical column and the new expanded column.
    RecordsToExpire = Table.SelectRows(ExpandedRecords, each
        // Skip records that did not match a current record (shouldn't happen with LeftOuter, but good check)
        [HR Worker] <> null and 
        // Check for any difference in the comparison fields (OR condition)
        (
            [Worker] <> [HR Worker] or 
            [Business Title] <> [HR Business Title] or 
            [Current Work Email] <> [HR Current Work Email] or 
            [#"Worker's Manager(s)"] <> [HR Workers Managers]
        )
    ),


RecordsWithDifferences =
    Table.AddColumn(
        RecordsToExpire,
        "Note_updated",
        each
            let
                newNote =
                    Text.Combine(
                        List.RemoveNulls({
                            if not Value.NullableEquals([Worker], [HR Worker]) then "Worker" else null,
                            if not Value.NullableEquals([Business Title], [HR Business Title]) then "Business Title" else null,
                            if not Value.NullableEquals([Current Work Email], [HR Current Work Email]) then "Current Work Email" else null,
                            if not Value.NullableEquals([#"Worker's Manager(s)"], [HR Workers Managers]) then "Worker's Manager(s)" else null
                        }),
                        ", "
                    )
            in

            if Text.Length(newNote) > 0 then
              if [Note] = null then  newNote
              else if Text.Length([Note]) = 0 then newNote
              else [Note] & "; "& newNote
            else [Note],
   type text ),




// remove the new columns from HR since these are the originals
    Expired_Remove_Columns = Table.RemoveColumns(RecordsWithDifferences, {"HR Worker",  "HR Business Title", "HR Current Work Email", "HR Workers Managers","Data_End_Date"}),

    Expired_versions_final = Table.AddColumn(Expired_Remove_Columns, "Data_End_Date", each Date.AddDays(Date.From(DateTime.LocalNow()), -1), type date),


    // Expired_versions_final = Table.AddColumn(Expired_UpdateDataEndDate, "Note", each "Changed", type text ),

    
    // Select the original historical columns only
    //Expired_Final = "= Table.SelectColumns(RecordsToExpire, {"Employee ID", "Worker", "Business Title", "Current Work Email", "Worker's Manager(s)", "Effective_Start_Date", "Effective_End_Date", "Data_Start_Date"}) & Table.RenameColumns(Table.SelectColumns(RecordsToExpire, {"New_Data_End_Date"}), {{"New_Data_End_Date", "Data_End_Date"}})

  // now take the same rows and replace the changed fields with the HR versions 
    NewVersions = RecordsWithDifferences, 

   NewVersions_Removed_Columns = Table.RemoveColumns(NewVersions,
     { "Data_Start_Date","Note","Worker","Business Title","Current Work Email","Worker's Manager(s)"}),


       NewVersions_Rename_Columns =  Table.RenameColumns(NewVersions_Removed_Columns,
       {{"HR Worker","Worker"},
        {"HR Business Title","Business Title"},{"Note_updated","Note"},
       {"HR Current Work Email","Current Work Email"},{ "HR Workers Managers","Worker's Manager(s)"}}), 


    NewVersions_final = Table.AddColumn(NewVersions_Rename_Columns, "Data_Start_Date", each DateTime.Date(DateTime.LocalNow()), type date),

 
 
 
    // 2. Prepare New Active Records for the Changed Employees
    // Filter the Current Snapshot to get the rows corresponding to the changed employees
   ChangedIDs = RecordsToExpire[Employee ID],
    //ChangedInserts = Table.SelectRows(H_Sorted_Rows, each List.Contains(ChangedIDs, [Employee ID])),
    ChangedIDsTable = Table.FromList(ChangedIDs, Splitter.SplitByNothing(), {"Employee ID"}),


    // 3. Identify Unchanged Records (Records that stay active)
    // The logic is now a direct comparison between the historical column and the new expanded column.
    UnchangedRecords = Table.SelectRows(ExpandedRecords, each
        // Record must have matched a current record
        [HR Worker] <> null and 
        // Check for NO difference in the comparison fields (AND condition)
        (
            [Worker] = [HR Worker] and 
            [Business Title] = [HR Business Title] and 
            [Current Work Email] = [HR Current Work Email] and 
            [#"Worker's Manager(s)"] = [HR Workers Managers]
        )
    ),

    Unchanged_RemoveHR_Columns = Table.RemoveColumns(UnchangedRecords, {"HR Worker",  "HR Business Title", "HR Current Work Email", "HR Workers Managers","Data_End_Date","Note"}),

    // Select the existing historical columns only
    Unchanged_Final = Unchanged_RemoveHR_Columns, //Table.SelectColumns(RecordsToKeep, {"Employee ID", "Worker", "Business Title", "Current Work Email", "Worker's Manager(s)", "Effective_Start_Date", "Effective_End_Date", "Data_Start_Date", "Data_End_Date"}),

    // ---  (TERM REPORT) DATA PREPARATION ---
    Term_Source = Excel.Workbook(File.Contents("C:\Users\doria\PowerQuery\HR Standard Report - Terminations.xlsx"), null, true),
    Term_Standard_Sheet = Term_Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    Term_Removed_Top_Rows = Table.Skip(Term_Standard_Sheet,2),
    Term_Promoted_Headers = Table.PromoteHeaders(Term_Removed_Top_Rows, [PromoteAllScalars=true]),
    Term_Filtered_Rows = Table.SelectRows(Term_Promoted_Headers, 
        each [#"Worker Type"] = "Employee" 
            and Text.Contains([Cost Center], "Sales") 
            and [#"Location Address - Country"] = "United States of America"
    ),
    //H_Rename_EID =   Table.RenameColumns(H_Filtered_Rows, {{"Employee ID","HR Employee ID"}}),
    Term_Remove_Cols = Table.RemoveColumns(Term_Filtered_Rows  ,{"Legal Name - Last Name", "Legal Name - First Name", "Worker", "Email", "Worker Type", "Job Profile", "Supervisory Organization", "Cost Center", "Location Name", "Location Address - Country", "Hire Date"}),
    Term_EID_Num = Table.TransformColumnTypes(Term_Remove_Cols,{{"Employee ID", type number}}),
    //H_Buffer_Table = Table.Buffer(H_EID_Num),
    Term_Sorted_Rows = Table.Sort(Term_EID_Num,{{"Employee ID", Order.Ascending}}),


    TermExp_JoinedRecords = Table.NestedJoin(
         ActiveMaster, // Left table (Master)
        {"Employee ID"},
        Term_Sorted_Rows, // Right table (Current)
       {"Employee ID"},
        "TermRecs",    
        JoinKind.Inner
    ),
    TermExp_Result = Table.RemoveColumns(TermExp_JoinedRecords, {"TermRecs"}),


   TermExp_Remove_Columns = Table.RemoveColumns(TermExp_Result, {"Term","Data_End_Date","Note"}),

     TermExp_UpdateDataEndDate = Table.AddColumn(TermExp_Remove_Columns, "Data_End_Date", each Date.AddDays(Date.From(DateTime.LocalNow()), -1), type date),
     TermExp_Final = TermExp_UpdateDataEndDate,

//  Yesterday = Date.AddDays(Date.From(DateTime.LocalNow()), -1)

   // Perform a INNER Join and Expand to align historical and current records for comparison.
    Term_JoinedRecords = Table.NestedJoin(
         ActiveMaster, // Left table (Master)
        {"Employee ID"},
        Term_Sorted_Rows, // Right table (Current)
       {"Employee ID"},
        "TermRecs",    
        JoinKind.Inner
    ),

    Term_Expanded = Table.ExpandTableColumn(
        Term_JoinedRecords,
        "TermRecs",
        {"Termination Date"},
         {"Termination Date"}
    ),


    TermExp_AppendToNote = Table.AddColumn(Term_Expanded, "Term_Note_update",
        each
        if [Note] = null then  "Termed"
        else if Text.Length([Note]) = 0 then "Termed"
        else [Note] & "; Termed"
      ),
       Term_Remove_Columns = Table.RemoveColumns(TermExp_AppendToNote, {"Term","Data_Start_Date","Data_End_Date","Note"}),

     Term_UpdateDataEndDate = Table.AddColumn(Term_Remove_Columns, "Data_End_Date", each  DateTime.Date(DateTime.LocalNow()),  type date),
     Term_UpdateDataStartDate = Table.AddColumn(Term_UpdateDataEndDate, "Data_Start_Date", each DateTime.Date(DateTime.LocalNow()), type date),
 
     Term_final =  Table.RenameColumns(Term_UpdateDataStartDate,
       {{"Termination Date","Term"},{"Term_Note_update","Note"}}),


    // 4. Combine all the new/modified data streams
    FinalCombine = Table.Combine({ 
        InactiveMaster_final,
        Unchanged_Final,          // Old records with a new end date
        Expired_versions_final,             // Unchanged active records
       NewVersions_final,         // New employees being inserted
        NewInserts_Final,      // New active records for changed employees
        TermExp_Final,
       Term_final 
    }),

//Final_Change_Types = Table.TransformColumnTypes(FinalCombine,{{"Hire Date", type date},{"Effective_Start_Date", type date},{"Effective_End_Date", type date}}),
Final_Change_Types = Table.TransformColumnTypes(FinalCombine,{{"Hire Date", type date}}),

FinalSCDTable =  Table.ReorderColumns(
    Final_Change_Types,
    {"Employee ID", "Data_Start_Date", "Data_End_Date", "Note", "Worker", "Business Title", "Current Work Email", "Email HR Data", "Current_Role", "Hire Date", "Effective_Start_Date", "Effective_End_Date", "Term", "Status", "LOA_Flag", "LOA_Dates", "Commission_Plan_Type", "Region", "District", "Territory", "Worker's Manager(s)", "MNGR_EID", "ORG_CODE", "Salesforce_Code", "Salesforce_Name", "Organization_Code", "Organization_Name", "Parent_Org_Code", "Parent_Org_Name", "Org_Type", "Parent_Org_Type", "Default_Type", "Commission Assignment"} )
in
   FinalSCDTable
