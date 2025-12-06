let
    // --- 1. HISTORICAL (MASTER) DATA PREPARATION ---
    // The Master file serves as the historical source (the existing SCD
    // Type 2 table).
    Master_Source = Excel.Workbook(
        File.Contents("C:\Users\doria\PowerQuery\master_file_template_DT_TEXT.xlsx"),
        null,
        true
    ),
    Sample_Master_Sheet = Master_Source{[Item="Employee_Master",Kind="Sheet"]}[Data],
    HistoricalSCD = Table.PromoteHeaders(Sample_Master_Sheet, [PromoteAllScalars=true]),
    // Convert key columns to their correct types. Data_End_Date is critical
    // for identifying active records.
    Master_Changed_Type = Table.TransformColumnTypes(HistoricalSCD,
        {
            {"Employee ID", type number},
            {"Data_End_Date", type date},
            {"Data_Start_Date", type date},
            {"Hire Date", type date}
        }
    ),
    Sorted_Master = Table.Sort(Master_Changed_Type,{{"Employee ID", Order.Ascending}}),

    // Separate the historical data into records that are CURRENTLY ACTIVE
    // and those that are EXPIRED.
    // Active records have Data_End_Date = null.
    ActiveMaster = Table.SelectRows(Sorted_Master,
        each [Data_End_Date] = null and [Employee ID] <> null
    ),
    InactiveMaster_final = Table.SelectRows(Sorted_Master,
        each [Data_End_Date] <> null and [Employee ID] <> null
    ),

    // --- 2. CURRENT (HR REPORT) DATA PREPARATION ---
    // This is the source snapshot that represents the CURRENT state of
    // employees.
    HR_Source = Excel.Workbook(
        File.Contents("C:\Users\doria\PowerQuery\HR Standard Report.xlsx"),
        null,
        true
    ),
    HR_Standard_Sheet = HR_Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    H_Removed_Top_Rows = Table.Skip(HR_Standard_Sheet,2),
    H_Promoted_Headers = Table.PromoteHeaders(H_Removed_Top_Rows,
        [PromoteAllScalars=true]
    ),

    // Apply the necessary filters for the scope of the dimension (e.g., US
    // Sales Employees).
    H_Filtered_Rows = Table.SelectRows(H_Promoted_Headers,
        each [#"Worker Type - Text"] = "Employee"
            and Text.Contains([Cost Center], "Sales")
            and [Country] = "United States of America"
    ),

    // Select and rename ONLY the columns needed for comparison and the final
    // dimension table.
    // Standardize column names for the join and comparison logic.
    H_Select_Rename = Table.SelectColumns(H_Filtered_Rows,
        {"Employee ID", "Worker", "Hire Date", "Business Title",
         "Current Work Email", "Worker's Manager(s)"}
    ),
    H_RenameColumns = Table.RenameColumns(H_Select_Rename,
        {
            {"Employee ID", "Current_Employee ID"},
            {"Worker", "Current_Worker"},
            {"Hire Date", "Current_Hire Date"},
            {"Business Title", "Current_Business Title"},
            {"Current Work Email", "Current_Current Work Email"},
            {"Worker's Manager(s)", "Current_Worker's Manager(s)"}
        }
    ),

    H_EID_Num = Table.TransformColumnTypes(H_RenameColumns,
        {{"Current_Employee ID", type number}}
    ),
    CurrentSnapshot = Table.Sort(H_EID_Num,
        {{"Current_Employee ID", Order.Ascending}}
    ),


    // --- 3. CURRENT (TERMINATION REPORT) DATA PREPARATION ---
    // Source snapshot for employees who have been terminated since the last
    // run.
    Term_Source = Excel.Workbook(
        File.Contents("C:\Users\doria\PowerQuery\HR Standard Report - Terminations.xlsx"),
        null,
        true
    ),
    Term_Standard_Sheet = Term_Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    Term_Removed_Top_Rows = Table.Skip(Term_Standard_Sheet,2),
    Term_Promoted_Headers = Table.PromoteHeaders(Term_Removed_Top_Rows,
        [PromoteAllScalars=true]
    ),
    Term_Filtered_Rows = Table.SelectRows(Term_Promoted_Headers,
        each [#"Worker Type"] = "Employee"
            and Text.Contains([Cost Center], "Sales")
            and [#"Location Address - Country"] = "United States of America"
    ),
    Term_Select_Cols = Table.SelectColumns(Term_Filtered_Rows,
        {"Employee ID", "Termination Date"}
    ),
    Term_EID_Num = Table.TransformColumnTypes(Term_Select_Cols,
        {{"Employee ID", type number}, {"Termination Date", type date}}
    ),
    TermSnapshot = Term_EID_Num,


    // --- 4. SCD TYPE 2 IMPLEMENTATION LOGIC ---

    // A. IDENTIFY TERMINATIONS: Find currently active employees in the Master
    // who are in the TermSnapshot.
    TerminatedRecs_Joined = Table.NestedJoin(
        ActiveMaster, "Employee ID",
        TermSnapshot, "Employee ID",
        "TermData", JoinKind.Inner
    ),
    // Expire the historical record. Set Data_End_Date to YESTERDAY.
    TerminatedRecs_Expired = Table.AddColumn(TerminatedRecs_Joined,
        "Data_End_Date_New",
        each Date.AddDays(Date.From(DateTime.LocalNow()), -1),
        type date
    ),
    TerminatedRecs_Note = Table.AddColumn(TerminatedRecs_Expired, "Note_Term",
        each if [Note] = null then "Terminated" else [Note] & "; Terminated",
        type text
    ),
    TerminatedRecs_Final = Table.SelectColumns(TerminatedRecs_Note,
        Table.ColumnNames(ActiveMaster) & {"Data_End_Date_New", "Note_Term"}
    ),
    TerminatedRecs_Final = Table.RenameColumns(TerminatedRecs_Final,
        {{"Data_End_Date", "Data_End_Date_Original"},
         {"Data_End_Date_New", "Data_End_Date"},
         {"Note", "Note_Original"},
         {"Note_Term", "Note"}}
    ),
    TerminatedRecs_Final = Table.RemoveColumns(TerminatedRecs_Final,
        {"Data_End_Date_Original", "Note_Original"}
    ),


    // B. IDENTIFY NEW INSERTS: Find records in CurrentSnapshot that DO NOT
    // exist in ActiveMaster. (Right Anti-Join)
    NewInserts = Table.NestedJoin(
        ActiveMaster, "Employee ID",
        CurrentSnapshot, "Current_Employee ID",
        "NewRec", JoinKind.RightAnti
    ),
    NewInserts_Expanded = Table.ExpandTableColumn(
        NewInserts, "NewRec",
        Table.ColumnNames(CurrentSnapshot)
    ),
    // Prepare New Inserts for final table. Data_Start_Date = TODAY,
    // Data_End_Date = null.
    NewInserts_Final = Table.SelectColumns(NewInserts_Expanded,
        Table.ColumnNames(CurrentSnapshot)
    ),
    NewInserts_Final = Table.RenameColumns(NewInserts_Final,
        {
            {"Current_Employee ID", "Employee ID"},
            {"Current_Worker", "Worker"},
            {"Current_Hire Date", "Hire Date"},
            {"Current_Business Title", "Business Title"},
            {"Current_Current Work Email", "Current Work Email"},
            {"Current_Worker's Manager(s)", "Worker's Manager(s)"}
        }
    ),
    NewInserts_Final = Table.AddColumn(NewInserts_Final,
        "Data_Start_Date",
        each DateTime.Date(DateTime.LocalNow()),
        type date
    ),
    NewInserts_Final = Table.AddColumn(NewInserts_Final,
        "Data_End_Date",
        each null,
        type date
    ),
    NewInserts_Final = Table.AddColumn(NewInserts_Final,
        "Note",
        each "New Hire",
        type text
    ),


    // C. IDENTIFY CHANGES: Find records in ActiveMaster that also exist in
    // CurrentSnapshot (Inner Join).
    JoinedRecords = Table.NestedJoin(
        ActiveMaster, "Employee ID",
        CurrentSnapshot, "Current_Employee ID",
        "CurrentRecord", JoinKind.Inner
    ),
    // Expand the current record columns to be side-by-side with historical
    // data for comparison.
    ExpandedRecords = Table.ExpandTableColumn(
        JoinedRecords, "CurrentRecord",
        Table.ColumnNames(CurrentSnapshot)
    ),

    // Identify records where at least one comparison field has changed.
    RecordsToExpire = Table.SelectRows(ExpandedRecords, each
        (
            [Worker] <> [Current_Worker] or
            [Business Title] <> [Current_Business Title] or
            [Current Work Email] <> [Current_Current Work Email] or
            [#"Worker's Manager(s)"] <> [#"Current_Worker's Manager(s)"]
        )
    ),

    // C.1. EXPIRE OLD VERSION: Expire the historical record. Set
    // Data_End_Date to YESTERDAY.
    Expired_versions = Table.AddColumn(RecordsToExpire,
        "Data_End_Date_New",
        each Date.AddDays(Date.From(DateTime.LocalNow()), -1),
        type date
    ),

    // Generate a list of changed fields for the 'Note' column.
    RecordsWithDifferences = Table.AddColumn(Expired_versions,
        "Note_updated",
        each Text.Combine(
            List.RemoveNulls(
                {
                    if not Value.NullableEquals([Worker], [Current_Worker])
                        then "Worker" else null,
                    if not Value.NullableEquals([Business Title], [Current_Business Title])
                        then "Business Title" else null,
                    if not Value.NullableEquals([Current Work Email], [Current_Current Work Email])
                        then "Current Work Email" else null,
                    if not Value.NullableEquals([#"Worker's Manager(s)"], [#"Current_Worker's Manager(s)"])
                        then "Worker's Manager(s)" else null
                }
            ), ", "
        ), type text
    ),
    Expired_versions_final = Table.SelectColumns(RecordsWithDifferences,
        Table.ColumnNames(ActiveMaster) & {"Data_End_Date_New", "Note_updated"}
    ),
    // Replace old date/note columns with the new calculated ones.
    Expired_versions_final = Table.RenameColumns(Expired_versions_final,
        {{"Data_End_Date", "Data_End_Date_Original"},
         {"Data_End_Date_New", "Data_End_Date"},
         {"Note", "Note_Original"},
         {"Note_updated", "Note"}}
    ),
    Expired_versions_final = Table.RemoveColumns(Expired_versions_final,
        {"Data_End_Date_Original", "Note_Original"}
    ),


    // C.2. CREATE NEW ACTIVE VERSION: Create the new active record with the
    // current values.
    NewVersions_Active = Table.SelectColumns(RecordsWithDifferences,
        {"Current_Employee ID", "Current_Worker", "Current_Hire Date",
         "Current_Business Title", "Current_Current Work Email",
         "Current_Worker's Manager(s)", "Note_updated"}
    ),
    // Rename columns to match the final dimension structure.
    NewVersions_Final = Table.RenameColumns(NewVersions_Active,
        {
            {"Current_Employee ID", "Employee ID"},
            {"Current_Worker", "Worker"},
            {"Current_Hire Date", "Hire Date"},
            {"Current_Business Title", "Business Title"},
            {"Current_Current Work Email", "Current Work Email"},
            {"Current_Worker's Manager(s)", "Worker's Manager(s)"},
            {"Note_updated", "Note"}
        }
    ),
    // Set Data_Start_Date = TODAY, Data_End_Date = null.
    NewVersions_Final = Table.AddColumn(NewVersions_Final,
        "Data_Start_Date",
        each DateTime.Date(DateTime.LocalNow()),
        type date
    ),
    NewVersions_Final = Table.AddColumn(NewVersions_Final,
        "Data_End_Date",
        each null,
        type date
    ),


    // D. IDENTIFY UNCHANGED: Find records that are active in Master and have
    // NO differences in the comparison fields.
    UnchangedRecords = Table.SelectRows(ExpandedRecords, each
        (
            [Worker] = [Current_Worker] and
            [Business Title] = [Current_Business Title] and
            [Current Work Email] = [Current_Current Work Email] and
            [#"Worker's Manager(s)"] = [#"Current_Worker's Manager(s)"]
        )
    ),
    // Keep only the original historical columns, as these records remain
    // active.
    Unchanged_Final = Table.SelectColumns(UnchangedRecords,
        Table.ColumnNames(ActiveMaster)
    ),


    // E. COMBINE ALL STREAMS
    // The final dimension table is the union of all historical and newly
    // processed records.
    FinalCombine = Table.Combine(
        {
            InactiveMaster_final,     // 1. Fully expired historical records
            TerminatedRecs_Final,     // 2. Historical records expired (Term)
            Expired_versions_final,   // 3. Historical records expired (Change)
            Unchanged_Final,          // 4. Historical records that remain active
            NewInserts_Final,         // 5. Brand new employees
            NewVersions_Final         // 6. New active records for changed EIDs
        }
    ),

    // The rest of the original query involved columns that were removed
    // earlier, so they must be added back with nulls to match the final
    // schema. (Analyst must confirm all required columns are present.)
    Final_AddMissingCols = Table.AddColumn(FinalCombine,
        "Email HR Data", each null, type text
    ),
    Final_AddMissingCols = Table.AddColumn(Final_AddMissingCols,
        "Current_Role", each null, type text
    ),
    Final_AddMissingCols = Table.AddColumn(Final_AddMissingCols,
        "Effective_Start_Date", each null, type date
    ),
    Final_AddMissingCols = Table.AddColumn(Final_AddMissingCols,
        "Effective_End_Date", each null, type date
    ),
    Final_AddMissingCols = Table.AddColumn(Final_AddMissingCols,
        "Term", each null, type text
    ),
    Final_AddMissingCols = Table.AddColumn(Final_AddMissingCols,
        "Status", each null, type text
    ),
    // ... (Add all other necessary columns from the final reorder step) ...

    FinalSCDTable = Table.ReorderColumns(
        Final_AddMissingCols,
        {
            "Employee ID", "Data_Start_Date", "Data_End_Date", "Note",
            "Worker", "Business Title", "Current Work Email", "Email HR Data",
            "Current_Role", "Hire Date", "Effective_Start_Date",
            "Effective_End_Date", "Term", "Status", "LOA_Flag", "LOA_Dates",
            "Commission_Plan_Type", "Region", "District", "Territory",
            "Worker's Manager(s)", "MNGR_EID", "ORG_CODE", "Salesforce_Code",
            "Salesforce_Name", "Organization_Code", "Organization_Name",
            "Parent_Org_Code", "Parent_Org_Name", "Org_Type", "Parent_Org_Type",
            "Default_Type", "Commission Assignment"
        }
    )
in
    FinalSCDTable
