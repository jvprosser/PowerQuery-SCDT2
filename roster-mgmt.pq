let
    // --- 1. HISTORICAL (MASTER) DATA PREPARATION ---
    // Get historical (Master) File


    BasePath = "C:\Users\doria\PowerQuery\",
    WeekPath = BasePath & pWeek & "\",
    
    HR_Source = Excel.Workbook(File.Contents(WeekPath & "HR Standard Report.xlsx"), null, true),
    Term_Source = Excel.Workbook(File.Contents(WeekPath & "HR Standard Report - Terminations.xlsx"), null, true),    
    L_Source = Excel.Workbook(File.Contents(WeekPath & "HR LOA Report.xlsx"), null, true),
    Master_Source = Excel.Workbook(File.Contents(BasePath & "master_file_template_DT_TEXT.xlsx"), null, true),
    
    Sample_Master_Sheet = Master_Source{[Item="Employee_Master",Kind="Sheet"]}[Data],
    HistoricalSCD = Table.PromoteHeaders(Sample_Master_Sheet, [PromoteAllScalars=true]),
    // Note: Effective_Start_Date and Effective_End_Date columns are retained but not used for the SCD Type 2 logic below.
    Master_Changed_Type = Table.TransformColumnTypes(HistoricalSCD,{{"Employee ID", type number},{"Data_End_Date", type date}, 
    {"Data_Start_Date", type date},{"Hire Date", type date}}),
    //Buffer_Sorted = Table.Buffer(Master_Changed_Type),
    Sorted_Master = Table.Sort(Master_Changed_Type,{{"Employee ID", Order.Ascending}}),

 // --- 2. CURRENT (HR REPORT) DATA PREPARATION ---

    L_Standard_Sheet = L_Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    L_Removed_Top_Rows = Table.Skip(L_Standard_Sheet,5),
    L_Promoted_Headers = Table.PromoteHeaders(L_Removed_Top_Rows, [PromoteAllScalars=true]),


    L_Remove_Cols = Table.RemoveColumns(L_Promoted_Headers  ,{"Hire Date", "Supervisory Organization", "Location", "Manager(s)", "HR Business Partner", "HR Partner", "Cost Centers", "Effect - Payroll", "Effect - Absence Accrual", "Effect - Stock Vesting", "Effect - Benefits", "Effect - Continuous Service Accrual"}),

    L_EID_Num = Table.TransformColumnTypes(L_Remove_Cols, {{"Employee ID", type number}}),

    L_RenameColumns =  Table.RenameColumns(L_EID_Num, {{"Employee ID","L Employee ID"},{"Worker","L Worker"}}),
    L_Sorted_Rows = Table.Sort(L_RenameColumns,{{"L Employee ID", Order.Ascending}}),

    // --- 2. CURRENT (HR REPORT) DATA PREPARATION ---

    HR_Standard_Sheet = HR_Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    H_Removed_Top_Rows = Table.Skip(HR_Standard_Sheet,2),
    H_Promoted_Headers1 = Table.PromoteHeaders(H_Removed_Top_Rows, [PromoteAllScalars=true]),
    H_Filtered_Rows = Table.SelectRows(H_Promoted_Headers1, 
        each [#"Worker Type - Text"] = "Employee" 
            and Text.Contains([Cost Center], "Sales") 
            and [Country] = "United States of America"
    ),
    //H_Rename_EID =   Table.RenameColumns(H_Filtered_Rows, {{"Employee ID","HR Employee ID"}}),
    H_Remove_Cols = Table.RemoveColumns(H_Filtered_Rows  ,{"Worker Type - Text", "Employee Type", "Fixed-Term End Date", "Contingent Worker Type", "Exempt / Non-Exempt", "Contract End Date", "Time Type", "Source", "First Name", "Last Name",  "Job Code", "Job Profile", "Business Title", "Company", "Supervisory Organization", "Supervisory Organization - Hierarchy", "Functional Organizational Hierarchy", "Functional Org Level 1", "Functional Org Level 2", "Functional Org Level 3", "Functional Org Level 4", "Functional Org Level 5", "Cost Center", "Cost Center Hierarchy", "Job Family", "Work Arrangement", "Current Work Arrangement", "Location", "Scheduled Weekly Hours", "Country", "City", "State", "Continuous Service Date", "Company Service Date", "Termination Date", "Length of Service", "Time in Job Profile", "Work Shift", "Is Manager", "Management Level", "HR Partner", "User Name", "Prior Work Email", "Number of Direct Reports", "Management Chain - Level 09", "Management Chain - Level 08", "Management Chain - Level 07", "Management Chain - Level 06", "Management Chain - Level 05", "Management Chain - Level 04", "Management Chain - Level 03", "Management Chain - Level 02", "Management Chain - Level 01"}),
    
    H_EID_Num = Table.TransformColumnTypes(H_Remove_Cols,{{"Employee ID", type number}}),
    //H_Buffer_Table = Table.Buffer(H_EID_Num),

    H_RenameColumns =  Table.RenameColumns(H_EID_Num,
       {{"Employee ID","HR Employee ID"},{"Worker","HR Worker"},{"Hire Date","HR Hire Date"},
        {"Job Title","HR Job Title"},{"Position ID", "HR Position ID"},
       {"Current Work Email","HR Current Work Email"},{ "Worker's Manager(s)","HR Worker's Manager(s)"}}), 


    H_Sorted_Rows = Table.Sort(H_RenameColumns,{{"HR Employee ID", Order.Ascending}}),
    
    // --- 3. SCD TYPE 2 IMPLEMENTATION LOGIC ---
    // get the expire records so we can add them back in 
    InactiveMaster_final = Table.SelectRows(Sorted_Master, each [Data_End_Date] <> null and [Employee ID] <> null),
    
    // Get only the currently ACTIVE records from the Master Data
    ActiveMaster = Table.SelectRows(Sorted_Master, each [Data_End_Date] = null and [Employee ID] <> null),
///// LOA 
 
// "Last Date for Which Paid",
// Join ActiveMaster to LOA table
L_NewLOA =
    Table.NestedJoin(
        ActiveMaster,              // Left table
        "Employee ID",
        L_Sorted_Rows,              // Right table
        "L Employee ID",
        "LOA",
        JoinKind.Inner
    ),

// Expand LOA records
L_NewLOAExpanded =
    Table.ExpandTableColumn(
        L_NewLOA,
        "LOA",
        {
            "L Employee ID",
            "L Worker",
            "Leave Type (Including Family)",
            "Last Day of Work",
            "First Day",
            "Estimated Last Day",
            "Actual Last Day",
            "Total Days",
            "Units Requested",
            "Unit of Time"
        }
    ),

// Ensure correct data types
L_ChangedTypes =
    Table.TransformColumnTypes(
        L_NewLOAExpanded,
        {
            {"Last Day of Work", type date},
            {"Estimated Last Day", type date},
            {"Actual Last Day", type date},
            {"Total Days", type text}
        }
    ),


// Group by Employee ID and calculate required values
// Group and create a single note column
L_GroupedLOA =
    Table.Group(
    L_ChangedTypes,
    {"L Employee ID"},
    {
         {"Earliest Last Day of Work", each List.Min([Last Day of Work]), type date},
            {"Latest Estimated Last Day", each List.Max([Estimated Last Day]), type date},
            {"Latest Actual Last Day", each List.Max([Actual Last Day]), type date},
{
            "L_LOA_Flag",
            each
                if List.Contains([Actual Last Day], null)
                then "LOA Started"
                else "LOA Ended",
            type text
        },
            {
            "L_LOA_Dates",
            each
                let
                    StartDate =
                        List.Min([Last Day of Work]),
                    EndDate =
                        List.Max(List.RemoveNulls([Estimated Last Day])),
                    DaysTotal =
                        Text.Combine(
                            List.RemoveNulls([Total Days]),
                            ", "
                        )
                in
                    ""
                    & Date.ToText(StartDate, "MM/dd/yyyy")
                    & " - "
                    & Date.ToText(EndDate, "MM/dd/yyyy")
                    & "(est) " & DaysTotal & "",
            type text
        }
    }
),

    CutoffDate = Date.AddDays(Date.From(pWeek), -7),
    TodayDate = Date.From(pWeek),

  L_OnLeaveNow = Table.SelectRows(L_GroupedLOA, each
        // Record must have matched a current record
        [L_LOA_Flag] = "LOA Started" and 
         Date.From([Earliest Last Day of Work]) >= CutoffDate   
         or ([L_LOA_Flag] = "LOA Ended" and  [Latest Actual Last Day] >= CutoffDate and  [Latest Actual Last Day] < TodayDate)
    ),


L_MergedBackExpire =
Table.Join(
        ActiveMaster,              // Left table
        "Employee ID",
    L_OnLeaveNow,           // grouped summary
    "L Employee ID",
    JoinKind.Inner
),

L_MergedBackExpireRemove_Cols = Table.RemoveColumns(L_MergedBackExpanded  ,{"Data_End_Date","L_LOA_Flag","L_LOA_Dates"}),
L_MergedBackExpire_Final = Table.AddColumn(L_MergedBackExpireRemove_Cols, "Data_End_Date", each Date.AddDays(Date.From(pWeek), -1), type date),
 

L_MergedBack =
Table.NestedJoin(
        ActiveMaster,              // Left table
        "Employee ID",
    L_OnLeaveNow,           // grouped summary
    "L Employee ID",
    "LOA",
    JoinKind.Inner
),

L_MergedBackExpanded =
Table.ExpandTableColumn(
    L_MergedBack,
    "LOA",
    {"L_LOA_Flag", "L_LOA_Dates"}
),

L_MergedRemove_Cols = Table.RemoveColumns(L_MergedBackExpanded  ,{"LOA_Flag","LOA_Dates","Data_Start_Date","Note"}),
L_MergeUpdateDataStartDate = Table.AddColumn(L_MergedRemove_Cols, "Data_Start_Date", each DateTime.Date(pWeek), type date),
L_MergeUpdateNote = Table.DuplicateColumn(L_MergeUpdateDataStartDate,"L_LOA_Flag", "Note"), 

L_MergedRenamedColumns_final = Table.RenameColumns(
    L_MergeUpdateNote,
    {
        {"L_LOA_Flag", "LOA_Flag"},
        {"L_LOA_Dates", "LOA_Dates"}
    }
),



 




// Perform a Right Anti-Join to find NEW records in the Current Snapshot (H_Sorted_Rows) 
    // that DO NOT exist in the Master (ActiveMaster).
    NewInserts = Table.NestedJoin(
        ActiveMaster, // Left table (Master)
        "Employee ID",
        H_Sorted_Rows, // Right table (Current)
       "HR Employee ID",
       "NTC",
        JoinKind.RightAnti
    ),

    NewInsertsExpanded = Table.ExpandTableColumn(
        NewInserts,
        "NTC",
          {"HR Employee ID","HR Worker", "HR Hire Date",  "HR Job Title", "HR Current Work Email", "HR Workers Managers", "HR Position ID"}
    ),
   Removed_Columns = Table.RemoveColumns(NewInsertsExpanded,
  {"Employee ID","Worker", "Hire Date",  "Job Title", "Current Work Email", "Worker's Manager(s)","Position ID",  "Data_Start_Date","Note"}
)//"Effective_End_Date",
,

    RenameNewInsertColumns =  Table.RenameColumns(Removed_Columns,
       {{"HR Employee ID","Employee ID"}, {"HR Worker","Worker"}, {"HR Hire Date","Hire Date"}, {"HR Job Title","Job Title"},{"HR Position ID","Position ID"},
       {"HR Current Work Email","Current Work Email"}, { "HR Workers Managers","Worker's Manager(s)"}}), 


    UpdateDataStartDate = Table.AddColumn(RenameNewInsertColumns, "Data_Start_Date", each DateTime.Date(pWeek), type date),
    
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
       {"HR Employee ID"},
        "CurrentRecord",
        JoinKind.Inner
    ),
    
    // Expand the current record columns and rename them to avoid conflict.
    // The comparison columns are {"Worker", "Job Title", "Current Work Email", "Worker's Manager(s)", "Position ID"}
    ExpandedRecords = Table.ExpandTableColumn(
        JoinedRecords, 
        "CurrentRecord", 
         {"HR Worker",  "HR Job Title", "HR Current Work Email", "HR Worker's Manager(s)", "HR Position ID"}
    ),

    // Identify records that have CHANGED and need to be expired (Type 2 change)
    // The logic is now a direct comparison between the historical column and the new expanded column.
    RecordsToExpire = Table.SelectRows(ExpandedRecords, each
        // Skip records that did not match a current record (shouldn't happen with LeftOuter, but good check)
        [HR Worker] <> null and 
        // Check for any difference in the comparison fields (OR condition)
        (
            [Worker] <> [HR Worker] or 
            [Job Title] <> [HR Job Title] or 
            [Current Work Email] <> [HR Current Work Email] or 
            [#"Worker's Manager(s)"] <> [#"HR Worker's Manager(s)"]
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
			       if not Value.NullableEquals([Worker], [HR Worker]) then "Worker Name Change" else null,
			       if not Value.NullableEquals([Job Title], [HR Job Title]) then "Job Title Change" else null,
			       if not Value.NullableEquals([Current Work Email], [HR Current Work Email]) then "Current Work Email Change" else null,
			       if not Value.NullableEquals([#"Worker's Manager(s)"], [#"HR Worker's Manager(s)"]) then "Worker's Manager(s) Change" else null
			   }),
			   ", "
		       )
	       in
           newNote,
//	       if Text.Length(newNote) > 0 then
//		 if [Note] = null then  newNote
//		 else if Text.Length([Note]) = 0 then newNote
//		 else [Note] & "; "& newNote
//	       else [Note],
      type text ),


// remove the new columns from HR since these are the originals and we keep thier values
    Expired_Remove_Columns = Table.RemoveColumns(RecordsWithDifferences, {"HR Worker",  "HR Job Title", "HR Current Work Email", "HR Worker's Manager(s)","Data_End_Date","Note_updated", "HR Position ID"}),

// Add a data end date of yesterday to expire the record
    Expired_versions_final = Table.AddColumn(Expired_Remove_Columns, "Data_End_Date", each Date.AddDays(Date.From(pWeek), -1), type date),

  // now make a copy of the same records but we remove the original values and put the new ones from HR in
   NewVersions = RecordsWithDifferences,
 
   NewVersions_Removed_Columns = Table.RemoveColumns(NewVersions,
     { "Data_Start_Date","Note","Worker","Job Title","Current Work Email","Worker's Manager(s)","Status","Effective_Start_Date", "Effective_End_Date", "Position ID"}),

    NewVersions_Rename_Columns =  Table.RenameColumns(NewVersions_Removed_Columns,
       {{"HR Worker","Worker"}, {"HR Job Title","Job Title"},{"Note_updated","Note"},{"HR Position ID", "Position ID"},
       {"HR Current Work Email","Current Work Email"},{ "HR Worker's Manager(s)","Worker's Manager(s)"}}), 

    NewVersions_status = Table.AddColumn(NewVersions_Rename_Columns, "Status", each "Active", type text),
    NewVersions_final = Table.AddColumn(NewVersions_status, "Data_Start_Date", each DateTime.Date(pWeek), type date),

    // 2. Prepare New Active Records for the Changed Employees
    // Filter the Current Snapshot to get the rows corresponding to the changed employees
    ChangedIDs = RecordsToExpire[Employee ID],

    ChangedIDsTable = Table.FromList(ChangedIDs, Splitter.SplitByNothing(), {"Employee ID"}), 

    // 3. Identify Unchanged Records (Records that stay active)
    // The logic is now a direct comparison between the historical column and the new expanded column.
    UnchangedRecords = Table.SelectRows(ExpandedRecords, each
        // Record must have matched a current record
        [HR Worker] <> null and 
        // Check for NO difference in the comparison fields (AND condition)
        (
            [Worker] = [HR Worker] and 
            [Job Title] = [HR Job Title] and 
            [Current Work Email] = [HR Current Work Email] and 
            [#"Worker's Manager(s)"] = [#"HR Worker's Manager(s)"]
        )
    ),

    Unchanged_RemoveHR_Columns = Table.RemoveColumns(UnchangedRecords, {"HR Worker",  "HR Job Title", "HR Current Work Email", "HR Worker's Manager(s)","Data_End_Date", "HR Position ID"}),

    // Select the existing historical columns only
    Unchanged_Final = Unchanged_RemoveHR_Columns, //Table.SelectColumns(RecordsToKeep, {"Employee ID", "Worker", "Job Title", "Current Work Email", "Worker's Manager(s)", "Effective_Start_Date", "Effective_End_Date", "Data_Start_Date", "Data_End_Date"}),

    // ---  (TERM REPORT) DATA PREPARATION ---

    Term_Standard_Sheet = Term_Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    Term_Removed_Top_Rows = Table.Skip(Term_Standard_Sheet,2),
    Term_Promoted_Headers = Table.PromoteHeaders(Term_Removed_Top_Rows, [PromoteAllScalars=true]),
    Term_Filtered_Rows = Table.SelectRows(Term_Promoted_Headers, 
        each [#"Worker Type"] = "Employee" 
            and Text.Contains([Cost Center], "Sales") 
            and [#"Location Address - Country"] = "United States of America"
    ),

    Term_Remove_Cols = Table.RemoveColumns(Term_Filtered_Rows  ,{"Legal Name - Last Name", "Legal Name - First Name", "Worker", "Email", "Worker Type", "Job Profile", "Supervisory Organization", "Cost Center", "Location Name", "Location Address - Country", "Hire Date"}),
    Term_EID_Num = Table.TransformColumnTypes(Term_Remove_Cols,{{"Employee ID", type number}}),

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


   TermExp_Remove_Columns = Table.RemoveColumns(TermExp_Result, {"Term","Data_End_Date"}),

     TermExp_UpdateDataEndDate = Table.AddColumn(TermExp_Remove_Columns, "Data_End_Date", each Date.AddDays(Date.From(pWeek), -1), type date),

   TermExp_Final = TermExp_UpdateDataEndDate,

//  Yesterday = Date.AddDays(Date.From(pWeek), -1)

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
        else "Termed"
        //else [Note] & "; Termed" // We are not appending. the original note is in the expired record.
      ),
   Term_Remove_Columns = Table.RemoveColumns(TermExp_AppendToNote, {"Term","Data_Start_Date","Data_End_Date","Note", "Effective_Start_Date", "Effective_End_Date"}),

     Term_UpdateDataEndDate = Table.AddColumn(Term_Remove_Columns, "Data_End_Date", each  DateTime.Date(pWeek),  type date),
     Term_UpdateDataStartDate = Table.AddColumn(Term_UpdateDataEndDate, "Data_Start_Date", each DateTime.Date(pWeek), type date),
     Term_dup_ESD = Table.DuplicateColumn(Term_UpdateDataStartDate,"Termination Date", "Effective_Start_Date"),
     Term_dup_EED = Table.DuplicateColumn(Term_dup_ESD,"Termination Date", "Effective_End_Date"), 
     Term_final =  Table.RenameColumns(Term_dup_EED, {{"Termination Date","Term"},{"Term_Note_update","Note"}}),


    // 4. Combine all the new/modified data streams
    FinalCombine = Table.Combine({
       L_MergedRenamedColumns_final ,
       L_MergedBackExpire_Final,
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
    {"Employee ID", "Data_Start_Date", "Data_End_Date", "Note", "Worker", "Job Title", "Current Work Email", "Email HR Data", "Current_Role", "Hire Date", "Effective_Start_Date", "Effective_End_Date", "Term", "Status", "LOA_Flag", "LOA_Dates", "Commission_Plan_Type", "Region", "District", "Territory", "Worker's Manager(s)", "MNGR_EID", "ORG_CODE", "Salesforce_Code", "Salesforce_Name", "Organization_Code", "Organization_Name", "Parent_Org_Code", "Parent_Org_Name", "Org_Type", "Parent_Org_Type", "Default_Type", "Commission Assignment"} )
in
   FinalSCDTable
