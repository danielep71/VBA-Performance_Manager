Attribute VB_Name = "M_cPM_Test"

Option Explicit     'Force explicit declaration of all variables

'
'==============================================================================
'
'                           M_cPM_RegressionTests
'
'==============================================================================
' PURPOSE
'   Full regression test suite for cPerformanceManager and its required
'   companion module M_cPM_TimeWasters
'
' WHY THIS EXISTS
'   A simple smoke test is useful, but it is not sufficient for a component
'   whose behavior depends on:
'
'     - multiple timing backends
'     - session-bound validation rules
'     - strict vs non-strict behavior
'     - diagnostic / reporting helpers
'     - explicit environment cleanup
'     - shared Excel Application-state coordination
'
'   This module therefore provides:
'
'     - isolated test cases
'     - reusable assertion helpers
'     - suite-level counters and summary output
'     - a dedicated worksheet log for durable regression evidence
'     - validation of both success paths and error / fallback paths
'
' ENTRY POINT
'   Run:
'
'     Run_cPerformanceManager_RegressionSuite
'
' DEPENDENCIES
'   - cPerformanceManager
'   - M_cPM_TimeWasters
'   - Excel Application object model
'
' NOTES
'   - Place this code in a STANDARD MODULE
'   - The class name must be exactly: cPerformanceManager
'   - The companion TW manager module must already be imported and compiled
'   - Results are written to a dedicated worksheet log and summarized in the
'     Immediate Window
'   - This suite assumes the current class surface where:
'       * MethodName(2) returns "GetTickCount / GetTickCount64"
'       * ElapsedTime supports formatting an already measured elapsed value
'       * OverheadMeasurement_Text supports an optional Iterations argument
'
' VERSION
'   1.0.3
'
' UPDATED
'   2026-04-15
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const cPM_SHEET_LOG    As String = "REGRESSION_cPM"
    
'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
    'Snapshot of Excel Application state used by TW regression tests
    Private Type T_AppState
        ScreenUpdating  As Boolean   'Application.ScreenUpdating
        EnableEvents    As Boolean   'Application.EnableEvents
        DisplayAlerts   As Boolean   'Application.DisplayAlerts
        Calculation     As Long      'Application.Calculation
        Cursor          As Long      'Application.Cursor
    End Type

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private m_TotalCases            As Long      'Number of executed test cases
    Private m_TotalAssertions       As Long      'Number of executed assertions
    Private m_TotalFailures         As Long      'Number of failed assertions
    Private m_CurrentCaseName       As String    'Current case name

    Private m_CaseAssertions_Begin  As Long      'Assertion count at case start
    Private m_CaseFailures_Begin    As Long      'Failure count at case start

    Private m_RunTimestamp          As String    'Timestamp shared by the whole run
    Private m_SummaryNextRow        As Long      'Next writable row in the case-summary section
    Private m_DetailNextRow         As Long      'Next writable row in the failure-detail section

'
'==============================================================================
'
'                               SUITE RUNNER
'
'==============================================================================

Public Sub Run_cPerformanceManager_RegressionSuite()
'
'==============================================================================
'                  RUN CPERFORMANCEMANAGER REGRESSION SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the full regression suite for cPerformanceManager
'
' WHY THIS EXISTS
'   This is the single public entry point for the regression module. It prepares
'   the dedicated regression worksheet, resets suite counters, runs every test
'   case in a deterministic order, and then prints / writes the final summary
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Forces a clean TW baseline before the suite begins
'   - Resets suite counters and run-level state
'   - Builds or rebuilds the dedicated regression worksheet
'   - Initializes the worksheet log and suite header
'   - Executes all regression cases in deterministic order
'   - Forces a clean TW baseline again at the end of the suite
'   - Prints the suite footer even when the runner fails after initialization
'   - Restores Application state before re-raising any unexpected runner error
'
' ERROR POLICY
'   Individual cases handle their own unexpected errors and record failures
'   This runner restores Application state and then re-raises unexpected
'   runner-level errors normally
'
' DEPENDENCIES
'   - PM_TW_EndAllSessions
'   - Suite_ResetCounters
'   - DEMO_Begin_FastMode
'   - DEMO_End_FastMode
'   - DEMO_Build_DemoTemplate
'   - Suite_InitLogSheet
'   - Suite_PrintHeader
'   - Suite_PrintFooter
'   - All private Test_* procedures in this module
'
' NOTES
'   The suite is intentionally ordered from simpler / core behaviors toward
'   shared-state and cleanup behaviors
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WB                  As Workbook             'Target workbook
    Dim WS_Test             As Worksheet            'Dedicated regression worksheet
    
    Dim FastModeState       As tDemoFastModeState   'Saved Application-state snapshot
    Dim FastModeOn          As Boolean              'TRUE when fast mode was entered
    Dim FooterCanPrint      As Boolean              'TRUE when footer/log infrastructure is ready
    
    Dim SavedErrNum         As Long                 'Captured error number
    Dim SavedErrSrc         As String               'Captured error source
    Dim SavedErrDesc        As String               'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Simulate a pressed button when the routine was launched by a shape
        Btn_Click
    'Force the shared TW manager to a clean baseline before the suite begins
        PM_TW_EndAllSessions
    'Reset suite-level counters and current-case tracking
        Suite_ResetCounters
    'Resolve the workbook that contains the regression module
        Set WB = ThisWorkbook
    'Capture and apply fast-mode Application settings
        DEMO_Begin_FastMode FastModeState
        FastModeOn = True
    'Show a wait cursor while the regression environment is being prepared
        Application.Cursor = xlWait

'------------------------------------------------------------------------------
' PREPARE REGRESSION SHEET
'------------------------------------------------------------------------------
    'Build or rebuild the dedicated regression worksheet
        DEMO_Build_DemoTemplate _
            cPM_SHEET_LOG, _
            "CLASS PERFORMANCE MANAGER", _
            "REGRESSION TESTS", , , , , , , , , , , , , , , "Q", 33
    'Resolve the prepared regression worksheet
        Set WS_Test = WB.Worksheets(cPM_SHEET_LOG)
    'Initialize the dedicated worksheet log
        Suite_InitLogSheet
    'Print the suite header
        Suite_PrintHeader
    'Mark that the footer can safely be printed during centralized cleanup
        FooterCanPrint = True

'------------------------------------------------------------------------------
' RUN REGRESSION CASES
'------------------------------------------------------------------------------
    'Validate constructor/default state
        Test_DefaultState
    'Validate valid MethodName mappings
        Test_MethodName_ValidIndices
    'Validate invalid MethodName behavior
        Test_MethodName_InvalidIndices
    'Validate StartTimer session-state transitions for all methods
        Test_StartTimer_SetsSessionState_AllMethods
    'Validate numeric elapsed-time reads for all methods
        Test_ElapsedSeconds_AllMethods
    'Validate formatted elapsed-time reads for all methods
        Test_ElapsedTime_AllMethods
    'Validate formatting of an already measured elapsed-seconds value
        Test_ElapsedTime_FormatExistingSeconds
    'Validate aligned-start timing for all methods
        Test_AlignedStart_AllMethods
    'Validate raw accessor behavior after a QPC measurement
        Test_Accessors_QPC
    'Validate strict-mode behavior when elapsed time is requested before StartTimer
        Test_StrictMode_ElapsedBeforeStart
    'Validate strict-mode behavior for explicit method mismatch
        Test_StrictMode_MethodMismatch
    'Validate strict-mode behavior for invalid start-method input
        Test_StrictMode_InvalidStartMethod
    'Validate non-strict fallback for invalid start-method input
        Test_NonStrictMode_InvalidStartFallback
    'Validate non-strict fallback for explicit elapsed-method mismatch
        Test_NonStrictMode_MethodMismatchFallback
    'Validate numeric overhead measurement helpers
        Test_OverheadMeasurement_Seconds
    'Validate formatted overhead measurement helpers
        Test_OverheadMeasurement_Text
    'Validate diagnostic / informational properties
        Test_Diagnostics_Properties
    'Validate Pause method 1
        Test_Pause_Method1
    'Validate Pause method 2
        Test_Pause_Method2
    'Validate Pause method 3
        Test_Pause_Method3
    'Validate Pause method 4
        Test_Pause_Method4
    'Validate blank-key behavior in the shared TW manager
        Test_TW_BlankKeyValidation
    'Validate single-instance TW lifecycle
        Test_TW_SingleInstance
    'Validate overlapping multi-instance TW lifecycle
        Test_TW_OverlappingInstances
    'Validate ResetEnvironment idempotence and cleanup behavior
        Test_ResetEnvironment_Idempotent

CleanExit:
'------------------------------------------------------------------------------
' FINALIZE
'------------------------------------------------------------------------------
    'Force the shared TW manager to a clean baseline after the suite ends
        PM_TW_EndAllSessions
    'Print the final suite summary when the logging infrastructure is available
        If FooterCanPrint Then Suite_PrintFooter
    'Bring the regression worksheet to the foreground for the user
        If Not WS_Test Is Nothing Then
            WS_Test.Activate
            WS_Test.Range("A1").Select
        End If
    'Restore the normal cursor
        Application.Cursor = xlDefault
    'Restore the original Excel Application state only when fast mode was entered
        If FastModeOn Then DEMO_End_FastMode FastModeState
    'Re-raise the original runner-level error after cleanup when needed
        If SavedErrNum <> 0 Then Err.Raise SavedErrNum, SavedErrSrc, SavedErrDesc
    
    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before centralized cleanup
        SavedErrNum = Err.Number
        SavedErrSrc = Err.Source
        SavedErrDesc = Err.Description
    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub


'
'==============================================================================
'
'                           SUITE / ASSERT HELPERS
'
'==============================================================================

Private Sub Suite_ResetCounters()
'
'==============================================================================
'                           SUITE RESET COUNTERS
'------------------------------------------------------------------------------
' PURPOSE
'   Resets all suite-level counters and current-case tracking
'
' WHY THIS EXISTS
'   The regression runner needs deterministic counters for every run so that the
'   final summary is meaningful and unaffected by any previous execution
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Clears:
'     - total case count
'     - total assertion count
'     - total failure count
'     - current-case name
'     - per-case start markers
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' RESET STATE
'------------------------------------------------------------------------------
    'Reset the executed-case counter
        m_TotalCases = 0

    'Reset the assertion counter
        m_TotalAssertions = 0

    'Reset the failure counter
        m_TotalFailures = 0

    'Clear the current-case name
        m_CurrentCaseName = vbNullString

    'Reset per-case start markers
        m_CaseAssertions_Begin = 0
        m_CaseFailures_Begin = 0

    'Clear run-level worksheet-log state
        m_RunTimestamp = vbNullString
        m_SummaryNextRow = 0
        m_DetailNextRow = 0
End Sub

Private Sub Suite_InitLogSheet()
'
'==============================================================================
'                           SUITE INIT LOG SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Creates or resets the dedicated worksheet log used by the regression suite
'
' WHY THIS EXISTS
'   Debug.Print alone is not a reliable medium for a large suite because older
'   lines become difficult to inspect. A dedicated worksheet preserves the full
'   run output in a structured, reviewable format
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Gets or creates the regression-log worksheet
'   - Stamps one run timestamp for the whole suite
'   - Clears the target log layout area
'   - Rebuilds the case-summary section
'   - Rebuilds the failure-details section
'   - Applies deterministic widths, alignments, headers, and borders
'   - Initializes row pointers for case summaries and failure details
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Suite_GetOrCreateLogSheet
'   - DEMO_Write_BandHeader
'   - DEMO_Set_RangeBorder
'
' NOTES
'   This routine is intentionally idempotent:
'     - it does not insert rows or columns
'     - it rebuilds the same fixed layout every time
'
' UPDATED
'   2026-04-17
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet    'Regression-log worksheet
    Dim SummaryHeaders      As Variant      'Case-summary column headers
    Dim DetailHeaders       As Variant      'Failure-detail column headers
    Dim i                   As Long         'Header loop index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Stamp the current run timestamp once for the whole suite
        m_RunTimestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    'Resolve the regression-log worksheet
        Set WS = Suite_GetOrCreateLogSheet()
    'Insert two columns
        WS.Columns("L:M").Insert _
            Shift:=xlToRight, _
            CopyOrigin:=xlFormatFromLeftOrAbove

'------------------------------------------------------------------------------
' DEFINE HEADER SETS
'------------------------------------------------------------------------------
    'Define the case-summary headers
        SummaryHeaders = Array( _
            "RunTimestamp", _
            "CaseNo", _
            "CaseName", _
            "Assertions", _
            "Failures", _
            "Result")
    'Define the failure-detail headers
        DetailHeaders = Array( _
            "RunTimestamp", _
            "CaseNo", _
            "CaseName", _
            "FailureType", _
            "Message", _
            "ErrNumber", _
            "ErrDescription")

'------------------------------------------------------------------------------
' APPLY COLUMN LAYOUT
'------------------------------------------------------------------------------
    'Apply the width layout for the case-summary section
        WS.Columns("C").ColumnWidth = 15
        WS.Columns("D").ColumnWidth = 10
        WS.Columns("E").ColumnWidth = 40
        WS.Columns("F:H").ColumnWidth = 10
    'Apply the width layout for the failure-detail section
        WS.Columns("I").ColumnWidth = 15
        WS.Columns("J").ColumnWidth = 10
        WS.Columns("K").ColumnWidth = 40
        WS.Columns("L").ColumnWidth = 10
        WS.Columns("M").ColumnWidth = 40
        WS.Columns("N").ColumnWidth = 10
        WS.Columns("O").ColumnWidth = 40

'------------------------------------------------------------------------------
' APPLY BASE ALIGNMENT
'------------------------------------------------------------------------------
    'Apply center alignment to the structured log block by default
        With WS.Range("C6:O30")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 9
        End With
    'Restore left alignment for text-heavy columns
        WS.Columns("E").HorizontalAlignment = xlLeft
        WS.Columns("K").HorizontalAlignment = xlLeft
        WS.Columns("M").HorizontalAlignment = xlLeft
        WS.Columns("O").HorizontalAlignment = xlLeft

'------------------------------------------------------------------------------
' BUILD CASE SUMMARY AREA
'------------------------------------------------------------------------------
    '--------------------------------------------------------------------------
    ' WRITE SECTION HEADER
    '--------------------------------------------------------------------------
        'Write the case-summary section header band
            DEMO_Write_BandHeader WS.Range("C4:H4"), "CASE SUMMARY"

    '--------------------------------------------------------------------------
    ' WRITE COLUMN CAPTIONS
    '--------------------------------------------------------------------------
        'Write the case-summary headers
            For i = LBound(SummaryHeaders) To UBound(SummaryHeaders)
                WS.Cells(5, "C").Offset(0, i).Value = SummaryHeaders(i)
            Next i

    '--------------------------------------------------------------------------
    ' FORMAT HEADER ROW
    '--------------------------------------------------------------------------
        'Apply standard header styling to the case-summary header row
            With WS.Range("C5:H5")
                .Interior.Color = COLOR_SUBHEADER
                .Font.Bold = True
                .Font.Color = vbWhite
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

    '--------------------------------------------------------------------------
    ' APPLY VISUAL FRAME
    '--------------------------------------------------------------------------
        'Apply a border around the visible case-summary block
            DEMO_Set_RangeBorder WS.Range("C4:H30")

'------------------------------------------------------------------------------
' BUILD FAILURE DETAILS AREA
'------------------------------------------------------------------------------
    '--------------------------------------------------------------------------
    ' WRITE SECTION HEADER
    '--------------------------------------------------------------------------
        'Write the failure-details section header band
            DEMO_Write_BandHeader WS.Range("I4:O4"), "FAILURE DETAILS"

    '--------------------------------------------------------------------------
    ' WRITE COLUMN CAPTIONS
    '--------------------------------------------------------------------------
        'Write the failure-detail headers
            For i = LBound(DetailHeaders) To UBound(DetailHeaders)
                WS.Cells(5, "I").Offset(0, i).Value = DetailHeaders(i)
            Next i

    '--------------------------------------------------------------------------
    ' FORMAT HEADER ROW
    '--------------------------------------------------------------------------
        'Apply standard header styling to the failure-detail header row
            With WS.Range("I5:O5")
                .Interior.Color = COLOR_SUBHEADER
                .Font.Bold = True
                .Font.Color = vbWhite
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

    '--------------------------------------------------------------------------
    ' APPLY VISUAL FRAME
    '--------------------------------------------------------------------------
        'Apply a border around the visible failure-detail block
            DEMO_Set_RangeBorder WS.Range("I4:O30")

'------------------------------------------------------------------------------
' INITIALIZE ROW POINTERS
'------------------------------------------------------------------------------
    'Initialize the first writable summary row
        m_SummaryNextRow = 6
    'Initialize the first writable detail row
        m_DetailNextRow = 6

'------------------------------------------------------------------------------
' POSITION CURSOR
'------------------------------------------------------------------------------
    'Park the cursor near the top-left of the visible log layout
        WS.Range("C4").Select

End Sub
Private Function Suite_GetOrCreateLogSheet() As Worksheet
'
'==============================================================================
'                        SUITE GET OR CREATE LOG SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the dedicated regression worksheet, creating it if missing
'
' WHY THIS EXISTS
'   The regression module should remain self-contained and should not depend on
'   external builder modules merely to create or resolve its output worksheet
'
' INPUTS
'   None
'
' RETURNS
'   Worksheet
'     Existing or newly created regression worksheet
'
' BEHAVIOR
'   - Searches ThisWorkbook for the dedicated regression sheet name
'   - Returns the existing worksheet when found
'   - Creates the worksheet at the end of the workbook when missing
'   - Assigns the standard regression sheet name to the new worksheet
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - ThisWorkbook
'   - cPM_SHEET_LOG
'
' NOTES
'   This helper intentionally uses the regression module's own worksheet-name
'   constant so the suite remains self-contained
'
' UPDATED
'   2026-04-17
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WB                  As Workbook     'Target workbook
    Dim WS                  As Worksheet    'Worksheet iterator / result

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a blank regression-sheet name constant
        If Len(Trim$(cPM_SHEET_LOG)) = 0 Then
            Err.Raise vbObjectError + 2400, _
                      "M_cPM_RegressionTests.Suite_GetOrCreateLogSheet", _
                      "Regression sheet name cannot be blank."
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the workbook that contains the regression module
        Set WB = ThisWorkbook

'------------------------------------------------------------------------------
' SEARCH EXISTING SHEETS
'------------------------------------------------------------------------------
    'Search for an existing worksheet with the requested regression-sheet name
        For Each WS In WB.Worksheets
            If StrComp(WS.Name, cPM_SHEET_LOG, vbTextCompare) = 0 Then
                Set Suite_GetOrCreateLogSheet = WS
                Exit Function
            End If
        Next WS

'------------------------------------------------------------------------------
' CREATE SHEET
'------------------------------------------------------------------------------
    'Create a new worksheet because the requested one does not yet exist
        Set WS = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
    'Assign the requested regression-sheet name to the new worksheet
        WS.Name = cPM_SHEET_LOG

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the existing or newly created regression worksheet
        Set Suite_GetOrCreateLogSheet = WS

End Function
Private Sub Suite_PrintHeader()
'
'==============================================================================
'                             SUITE PRINT HEADER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints the suite-level header to the Immediate Window
'
' WHY THIS EXISTS
'   The header makes a regression run easy to identify and separates one run
'   from prior Immediate Window output
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Prints:
'     - a delimiter line
'     - suite title
'     - timestamp
'     - worksheet-log target
'     - a second delimiter line
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' PRINT HEADER
'------------------------------------------------------------------------------
    'Print the opening delimiter
        Debug.Print String$(100, "=")
    'Print the suite title
        Debug.Print "REGRESSION SUITE START: cPerformanceManager"
    'Print the suite timestamp
        Debug.Print "Timestamp: " & m_RunTimestamp
    'Print the worksheet-log target
        Debug.Print "Worksheet log: " & cPM_SHEET_LOG
    'Print the closing delimiter for the header block
        Debug.Print String$(100, "=")
End Sub

Private Sub Suite_PrintFooter()
'
'==============================================================================
'                             SUITE PRINT FOOTER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints the suite-level summary to the Immediate Window
'
' WHY THIS EXISTS
'   A regression run should end with a concise summary that reports how many
'   cases and assertions ran and whether any failures were recorded
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Prints:
'     - total cases
'     - total assertions
'     - total failures
'     - overall pass / fail status
'     - the worksheet-log sheet name
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SUMMARY
'------------------------------------------------------------------------------
    'Print a delimiter line before the summary block
        Debug.Print String$(100, "-")
    'Print the total number of executed cases
        Debug.Print "Total cases      : " & m_TotalCases
    'Print the total number of executed assertions
        Debug.Print "Total assertions : " & m_TotalAssertions
    'Print the total number of recorded failures
        Debug.Print "Total failures   : " & m_TotalFailures
    'Print the overall suite status
        If m_TotalFailures = 0 Then
            Debug.Print "OVERALL RESULT   : PASS"
        Else
            Debug.Print "OVERALL RESULT   : FAIL"
        End If
    'Print the worksheet-log target
        Debug.Print "Worksheet log    : " & cPM_SHEET_LOG
    'Print a final delimiter line
        Debug.Print String$(100, "=")
End Sub

Private Sub Case_Begin( _
    ByVal CaseName As String)
'
'==============================================================================
'                                CASE BEGIN
'------------------------------------------------------------------------------
' PURPOSE
'   Marks the start of one regression case
'
' WHY THIS EXISTS
'   This helper centralizes case counting, current-case naming, and per-case
'   counter baselining while keeping Immediate Window output compact
'
' INPUTS
'   CaseName
'     Human-readable regression case name
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Increments the total case counter
'   - Stores the active case name
'   - Captures assertion/failure counters at case start
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' UPDATE STATE
'------------------------------------------------------------------------------
    'Increment the total number of executed cases
        m_TotalCases = m_TotalCases + 1

    'Store the current case name
        m_CurrentCaseName = CaseName

    'Capture per-case starting counters
        m_CaseAssertions_Begin = m_TotalAssertions
        m_CaseFailures_Begin = m_TotalFailures
End Sub

Private Sub Case_Finalize()
'
'==============================================================================
'                              CASE FINALIZE
'------------------------------------------------------------------------------
' PURPOSE
'   Finalizes one regression case by writing a compact case summary to the
'   worksheet log and to the Immediate Window
'
' WHY THIS EXISTS
'   Printing one compact line per case keeps the Immediate Window readable,
'   while the worksheet preserves the full structured run output
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Resolves the dedicated regression-log worksheet
'   - Computes per-case assertion count
'   - Computes per-case failure count
'   - Resolves the per-case PASS / FAIL result
'   - Writes one compact summary row to the worksheet log
'   - Prints one compact PASS / FAIL line to the Immediate Window
'   - Advances the summary-row pointer
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Suite_GetOrCreateLogSheet
'
' UPDATED
'   2026-04-17
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet    'Regression-log worksheet
    Dim CaseAssertions      As Long         'Assertions executed in this case
    Dim CaseFailures        As Long         'Failures recorded in this case
    Dim ResultText          As String       'Per-case result text

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the regression-log worksheet
        Set WS = Suite_GetOrCreateLogSheet()
    'Compute assertions executed in this case
        CaseAssertions = m_TotalAssertions - m_CaseAssertions_Begin
    'Compute failures recorded in this case
        CaseFailures = m_TotalFailures - m_CaseFailures_Begin
    'Resolve the per-case result text
        If CaseFailures = 0 Then
            ResultText = "PASS"
        Else
            ResultText = "FAIL"
        End If

'------------------------------------------------------------------------------
' WRITE WORKSHEET SUMMARY ROW
'------------------------------------------------------------------------------
    'Write the run timestamp
        WS.Cells(m_SummaryNextRow, "C").Value = m_RunTimestamp
    'Write the case number
        WS.Cells(m_SummaryNextRow, "D").Value = m_TotalCases
    'Write the case name
        WS.Cells(m_SummaryNextRow, "E").Value = m_CurrentCaseName
    'Write the assertion count
        WS.Cells(m_SummaryNextRow, "F").Value = CaseAssertions
    'Write the failure count
        WS.Cells(m_SummaryNextRow, "G").Value = CaseFailures
    'Write the pass / fail result
        WS.Cells(m_SummaryNextRow, "H").Value = ResultText

'------------------------------------------------------------------------------
' PRINT COMPACT CASE SUMMARY
'------------------------------------------------------------------------------
    'Print one compact case summary line
        If CaseFailures = 0 Then
            Debug.Print "CASE " & Format$(m_TotalCases, "00") & " PASS - " & _
                        m_CurrentCaseName & " | assertions=" & CStr(CaseAssertions)
        Else
            Debug.Print "CASE " & Format$(m_TotalCases, "00") & " FAIL - " & _
                        m_CurrentCaseName & " | assertions=" & CStr(CaseAssertions) & _
                        " | failures=" & CStr(CaseFailures)
        End If

'------------------------------------------------------------------------------
' ADVANCE POINTER
'------------------------------------------------------------------------------
    'Advance the summary-row pointer
        m_SummaryNextRow = m_SummaryNextRow + 1

End Sub

Private Sub AssertTrue( _
    ByVal Condition As Boolean, _
    ByVal MessageText As String)
'
'==============================================================================
'                                ASSERT TRUE
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion based on a Boolean condition
'
' WHY THIS EXISTS
'   Most regression checks naturally reduce to a Boolean predicate. This helper
'   centralizes assertion counting and failure logging while avoiding excessive
'   Immediate Window noise from passing assertions
'
' INPUTS
'   Condition
'     Boolean result to evaluate
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Increments the assertion counter
'   - Does not print passing assertions
'   - Records failing assertions in the worksheet log
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' UPDATE ASSERTION COUNT
'------------------------------------------------------------------------------
    'Increment the total assertion count
        m_TotalAssertions = m_TotalAssertions + 1

'------------------------------------------------------------------------------
' RECORD FAILURE ONLY
'------------------------------------------------------------------------------
    'Exit quietly for passing assertions
        If Condition Then Exit Sub
    'Count one failing assertion
        m_TotalFailures = m_TotalFailures + 1
    'Write the failure detail to the worksheet log
        LogFailureDetail "ASSERT", MessageText
End Sub

Private Sub AssertEqualLong( _
    ByVal Expected As Long, _
    ByVal Actual As Long, _
    ByVal MessageText As String)
'
'==============================================================================
'                             ASSERT EQUAL LONG
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion for Long equality
'
' WHY THIS EXISTS
'   Many regression checks compare numeric IDs, counters, and enumerated Excel
'   Application values that are naturally represented as Long
'
' INPUTS
'   Expected
'     Expected Long value
'
'   Actual
'     Actual Long value
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' BEHAVIOR
'   Delegates to AssertTrue using a Long equality comparison
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Assert that the actual Long equals the expected Long
        AssertTrue (Actual = Expected), _
                   MessageText & " | expected=" & CStr(Expected) & _
                   " actual=" & CStr(Actual)
End Sub

Private Sub AssertEqualBoolean( _
    ByVal Expected As Boolean, _
    ByVal Actual As Boolean, _
    ByVal MessageText As String)
'
'==============================================================================
'                           ASSERT EQUAL BOOLEAN
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion for Boolean equality
'
' WHY THIS EXISTS
'   Several class and Application-state checks are Boolean in nature, such as:
'     - StrictMode
'     - HasActiveSession
'     - TW_IsActive
'     - ScreenUpdating
'     - EnableEvents
'     - DisplayAlerts
'
' INPUTS
'   Expected
'     Expected Boolean value
'
'   Actual
'     Actual Boolean value
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' BEHAVIOR
'   Delegates to AssertTrue using a Boolean equality comparison
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Assert that the actual Boolean equals the expected Boolean
        AssertTrue (Actual = Expected), _
                   MessageText & " | expected=" & CStr(Expected) & _
                   " actual=" & CStr(Actual)
End Sub

Private Sub AssertEqualString( _
    ByVal Expected As String, _
    ByVal Actual As String, _
    ByVal MessageText As String)
'
'==============================================================================
'                            ASSERT EQUAL STRING
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion for String equality
'
' WHY THIS EXISTS
'   Method-name and text-reporting checks often require exact String comparison
'
' INPUTS
'   Expected
'     Expected String
'
'   Actual
'     Actual String
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' BEHAVIOR
'   Delegates to AssertTrue using a String equality comparison
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Assert that the actual String equals the expected String
        AssertTrue (Actual = Expected), _
                   MessageText & " | expected=\" & Expected & _
                   "\ actual=\" & Actual & "\"
End Sub

Private Sub AssertContainsString( _
    ByVal SourceText As String, _
    ByVal SubText As String, _
    ByVal MessageText As String)
'
'==============================================================================
'                          ASSERT CONTAINS STRING
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion that one String contains another
'
' WHY THIS EXISTS
'   Several formatted reports are not suitable for exact equality checks, but
'   they should still contain required semantic markers such as:
'     - "milliseconds"
'     - "seconds"
'     - backend labels
'
' INPUTS
'   SourceText
'     Full text to inspect
'
'   SubText
'     Required substring
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' BEHAVIOR
'   Delegates to AssertTrue using a case-insensitive substring search
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT CONTAINS
'------------------------------------------------------------------------------
    'Assert that the source text contains the required substring
        AssertTrue (InStr(1, SourceText, SubText, vbTextCompare) > 0), _
                   MessageText & " | required=\" & SubText & " \ "
End Sub

Private Sub AssertApproxDouble( _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal Tolerance As Double, _
    ByVal MessageText As String)
'
'==============================================================================
'                           ASSERT APPROX DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion for approximate Double equality
'
' WHY THIS EXISTS
'   Floating-point timing values are often best compared within a tolerance
'   rather than by exact equality
'
' INPUTS
'   Expected
'     Expected Double value
'
'   Actual
'     Actual Double value
'
'   Tolerance
'     Maximum allowed absolute difference
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' BEHAVIOR
'   Delegates to AssertTrue using:
'
'     Abs(Actual - Expected) <= Tolerance
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT APPROXIMATE EQUALITY
'------------------------------------------------------------------------------
    'Assert that the absolute difference is within the requested tolerance
        AssertTrue (Abs(Actual - Expected) <= Tolerance), _
                   MessageText & " | expected=" & Format$(Expected, "0.000000000") & _
                   " actual=" & Format$(Actual, "0.000000000") & _
                   " tol=" & Format$(Tolerance, "0.000000000")
End Sub

Private Sub AssertInRangeDouble( _
    ByVal LowerBound As Double, _
    ByVal UpperBound As Double, _
    ByVal Actual As Double, _
    ByVal MessageText As String)
'
'==============================================================================
'                           ASSERT INRANGE DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion that a Double lies within a closed interval
'
' WHY THIS EXISTS
'   Pause and elapsed-time sanity tests are often best validated by acceptable
'   lower/upper bounds rather than exact target values
'
' INPUTS
'   LowerBound
'     Minimum acceptable value
'
'   UpperBound
'     Maximum acceptable value
'
'   Actual
'     Actual measured value
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' BEHAVIOR
'   Delegates to AssertTrue using:
'
'     Actual >= LowerBound And Actual <= UpperBound
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT RANGE
'------------------------------------------------------------------------------
    'Assert that the actual value lies within the closed interval
        AssertTrue ((Actual >= LowerBound) And (Actual <= UpperBound)), _
                   MessageText & " | range=[" & Format$(LowerBound, "0.000000000") & _
                   ", " & Format$(UpperBound, "0.000000000") & _
                   "] actual=" & Format$(Actual, "0.000000000")
End Sub

Private Sub AssertNonNegativeDouble( _
    ByVal Actual As Double, _
    ByVal MessageText As String)
'
'==============================================================================
'                        ASSERT NONNEGATIVE DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion that a Double is nonnegative
'
' WHY THIS EXISTS
'   Many timing and diagnostic values should never be negative, even when they
'   are too coarse or too small to be meaningfully positive
'
' INPUTS
'   Actual
'     Actual Double value
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' BEHAVIOR
'   Delegates to AssertTrue using:
'
'     Actual >= 0#
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT NONNEGATIVE
'------------------------------------------------------------------------------
    'Assert that the actual value is nonnegative
        AssertTrue (Actual >= 0#), _
                   MessageText & " | actual=" & Format$(Actual, "0.000000000")
End Sub

Private Sub LogFailureDetail( _
    ByVal FailureType As String, _
    ByVal MessageText As String, _
    Optional ByVal ErrNumberIn As Variant, _
    Optional ByVal ErrDescriptionIn As Variant)
'
'==============================================================================
'                           LOG FAILURE DETAIL
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one detailed failure row to the worksheet log
'
' WHY THIS EXISTS
'   Failure details should be preserved in a structured, filterable medium
'   rather than relying only on the Immediate Window
'
' INPUTS
'   FailureType
'     Short failure category such as ASSERT or ERROR
'
'   MessageText
'     Human-readable failure message
'
'   ErrNumberIn (optional)
'     Optional error number for unexpected-error cases
'
'   ErrDescriptionIn (optional)
'     Optional error description for unexpected-error cases
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Resolves the dedicated regression-log worksheet
'   - Writes one row into the failure-detail section of the worksheet log
'   - Writes optional error fields only when supplied
'   - Advances the detail-row pointer
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Suite_GetOrCreateLogSheet
'
' UPDATED
'   2026-04-17
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet    'Regression-log worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the regression-log worksheet
        Set WS = Suite_GetOrCreateLogSheet()

'------------------------------------------------------------------------------
' WRITE DETAIL ROW
'------------------------------------------------------------------------------
    'Write the run timestamp
        WS.Cells(m_DetailNextRow, "I").Value = m_RunTimestamp
    'Write the case number
        WS.Cells(m_DetailNextRow, "J").Value = m_TotalCases
    'Write the case name
        WS.Cells(m_DetailNextRow, "K").Value = m_CurrentCaseName
    'Write the failure type
        WS.Cells(m_DetailNextRow, "L").Value = FailureType
    'Write the failure message
        WS.Cells(m_DetailNextRow, "M").Value = MessageText
    'Write the optional error number when supplied
        If Not IsMissing(ErrNumberIn) Then
            If Not IsEmpty(ErrNumberIn) Then
                WS.Cells(m_DetailNextRow, "N").Value = ErrNumberIn
            End If
        End If
    'Write the optional error description when supplied
        If Not IsMissing(ErrDescriptionIn) Then
            If Not IsEmpty(ErrDescriptionIn) Then
                WS.Cells(m_DetailNextRow, "O").Value = ErrDescriptionIn
            End If
        End If

'------------------------------------------------------------------------------
' ADVANCE POINTER
'------------------------------------------------------------------------------
    'Advance the detail-row pointer
        m_DetailNextRow = m_DetailNextRow + 1

End Sub
Private Sub RecordUnexpectedError( _
    ByVal ProcName As String)
'
'==============================================================================
'                          RECORD UNEXPECTED ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Records one unexpected test-case error as a suite failure
'
' WHY THIS EXISTS
'   A regression case may encounter an unexpected runtime error before reaching
'   some or all of its explicit assertions. This helper converts that event into
'   a recorded test failure and stores diagnostic information in the worksheet
'   log
'
' INPUTS
'   ProcName
'     Name of the regression procedure that encountered the error
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Increments the assertion counter by one
'   - Increments the failure counter by one
'   - Stores the procedure name, error number, and error description
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
    'Count one synthetic assertion for the unexpected error event
        m_TotalAssertions = m_TotalAssertions + 1

    'Count one failure for the unexpected error event
        m_TotalFailures = m_TotalFailures + 1

'------------------------------------------------------------------------------
' STORE DIAGNOSTIC
'------------------------------------------------------------------------------
    'Write the unexpected error detail to the worksheet log
        LogFailureDetail "ERROR", _
                         "Unexpected error in " & ProcName, _
                         Err.Number, _
                         Err.Description
End Sub

Private Sub CaptureAppState( _
    ByRef StateOut As T_AppState)
'
'==============================================================================
'                            CAPTURE APP STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Captures the current Excel Application state used by TW regression tests
'
' WHY THIS EXISTS
'   TW tests need a precise before/after baseline for Application properties
'   that are intentionally modified by the shared TW manager
'
' INPUTS
'   StateOut
'     Output structure that receives the Application snapshot
'
' RETURNS
'   None
'
' BEHAVIOR
'   Copies current Application values for:
'     - ScreenUpdating
'     - EnableEvents
'     - DisplayAlerts
'     - Calculation
'     - Cursor
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' CAPTURE STATE
'------------------------------------------------------------------------------
    'Copy the current Excel Application state into the output structure
        With Application
            StateOut.ScreenUpdating = .ScreenUpdating
            StateOut.EnableEvents = .EnableEvents
            StateOut.DisplayAlerts = .DisplayAlerts
            StateOut.Calculation = .Calculation
            StateOut.Cursor = .Cursor
        End With
End Sub

Private Function DelayForTimingMethod( _
    ByVal iMethod As Integer) _
    As Double
'
'==============================================================================
'                         DELAY FOR TIMING METHOD
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a practical per-method delay used by timing regression tests
'
' WHY THIS EXISTS
'   Different timing backends have different practical resolution
'   characteristics. In particular, method 6 (Now() * 86400) is much coarser
'   for test purposes than QPC
'
' INPUTS
'   iMethod
'     Timing backend identifier
'
' RETURNS
'   Double
'     Suggested delay in seconds for regression tests
'
' BEHAVIOR
'   - Returns 1.1 seconds for method 6
'   - Returns 0.05 seconds for all other methods
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Use a longer delay for the coarse wall-clock method
        If iMethod = 6 Then
            DelayForTimingMethod = 1.1
            Exit Function
        End If

    'Use a shorter delay for the remaining methods
        DelayForTimingMethod = 0.05
End Function

'
'==============================================================================
'
'                              REGRESSION CASES
'
'==============================================================================

Private Sub Test_DefaultState()
'
'==============================================================================
'                              TEST DEFAULT STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the constructor/default state of a fresh cPerformanceManager
'   instance
'
' WHY THIS EXISTS
'   A deterministic constructor baseline is essential for predictable timing,
'   validation behavior, and TW lifecycle behavior
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Verifies that a newly created instance starts with:
'     - StrictMode = True
'     - HasActiveSession = False
'     - ActiveMethodID = 0
'     - T1 = 0
'     - T2 = 0
'     - ET = 0
'     - TW_IsActive = False
'     - TW_ActiveSessionCount = 0
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM As cPerformanceManager    'Class under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Default constructor state"
    'Create a fresh class instance
        Set cPM = New cPerformanceManager
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Assert the default strict-mode state
        AssertEqualBoolean True, cPM.StrictMode, "StrictMode defaults to True"
    'Assert that no active timing session exists yet
        AssertEqualBoolean False, cPM.HasActiveSession, "HasActiveSession defaults to False"
    'Assert that no active method is bound yet
        AssertEqualLong 0, cPM.ActiveMethodID, "ActiveMethodID defaults to 0"
    'Assert the default raw/cached timing values
        AssertApproxDouble 0#, cPM.T1, 0#, "T1 defaults to 0"
        AssertApproxDouble 0#, cPM.T2, 0#, "T2 defaults to 0"
        AssertApproxDouble 0#, cPM.ET, 0#, "ET defaults to 0"
    'Assert that no TW session is active for the new instance
        AssertEqualBoolean False, cPM.TW_IsActive, "TW_IsActive defaults to False"
    'Assert that the shared TW manager is currently idle
        AssertEqualLong 0, cPM.TW_ActiveSessionCount, "TW_ActiveSessionCount defaults to 0"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_DefaultState"

    Resume CleanExit
End Sub

Private Sub Test_MethodName_ValidIndices()
'
'==============================================================================
'                        TEST METHODNAME VALID INDICES
'------------------------------------------------------------------------------
' PURPOSE
'   Validates exact MethodName mappings for valid indices 1..6
'
' WHY THIS EXISTS
'   The method-name map is both a public diagnostic surface and a documentation
'   contract for the class
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM As cPerformanceManager    'Class under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "MethodName valid indices"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Assert each documented method label exactly
        AssertEqualString "Timer", cPM.MethodName(1), "MethodName(1)"
        AssertEqualString "GetTickCount / GetTickCount64", cPM.MethodName(2), "MethodName(2)"
        AssertEqualString "timeGetTime", cPM.MethodName(3), "MethodName(3)"
        AssertEqualString "timeGetSystemTime", cPM.MethodName(4), "MethodName(4)"
        AssertEqualString "QPC", cPM.MethodName(5), "MethodName(5)"
        AssertEqualString "Now()", cPM.MethodName(6), "MethodName(6)"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release the class instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_MethodName_ValidIndices"

    Resume CleanExit
End Sub

Private Sub Test_MethodName_InvalidIndices()
'
'==============================================================================
'                       TEST METHODNAME INVALID INDICES
'------------------------------------------------------------------------------
' PURPOSE
'   Validates MethodName behavior for out-of-range indices
'
' WHY THIS EXISTS
'   The class documents that invalid MethodName indices should return
'   vbNullString rather than raising or returning misleading text
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM As cPerformanceManager    'Class under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "MethodName invalid indices"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Assert vbNullString for representative invalid indices
        AssertEqualString vbNullString, cPM.MethodName(0), "MethodName(0)"
        AssertEqualString vbNullString, cPM.MethodName(-1), "MethodName(-1)"
        AssertEqualString vbNullString, cPM.MethodName(7), "MethodName(7)"
        AssertEqualString vbNullString, cPM.MethodName(99), "MethodName(99)"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release the class instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_MethodName_InvalidIndices"

    Resume CleanExit
End Sub

Private Sub Test_StartTimer_SetsSessionState_AllMethods()
'
'==============================================================================
'                 TEST STARTTIMER SETS SESSION STATE ALL METHODS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates StartTimer session-state transitions for all timing methods
'
' WHY THIS EXISTS
'   StartTimer is the root of the session-bound timing model. A regression in
'   this area undermines elapsed-time validity across the whole class
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM     As cPerformanceManager    'Class under test
    Dim iMethod As Integer                'Timing backend index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "StartTimer sets session state for all methods"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6

            'Start a new timing session for the selected backend
                cPM.StartTimer iMethod, False

            'Assert that a session is now active
                AssertEqualBoolean True, cPM.HasActiveSession, _
                                   "HasActiveSession after StartTimer(" & CStr(iMethod) & ")"

            'Assert that the active method ID matches the requested method
                AssertEqualLong iMethod, cPM.ActiveMethodID, _
                                "ActiveMethodID after StartTimer(" & CStr(iMethod) & ")"

            'Assert that the method name is available for the active method
                AssertTrue (Len(cPM.MethodName(cPM.ActiveMethodID)) > 0), _
                           "MethodName available after StartTimer(" & CStr(iMethod) & ")"

            'Assert that the raw start capture is nonnegative
                AssertNonNegativeDouble cPM.T1, _
                                        "T1 nonnegative after StartTimer(" & CStr(iMethod) & ")"
        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_StartTimer_SetsSessionState_AllMethods"

    Resume CleanExit
End Sub

Private Sub Test_ElapsedSeconds_AllMethods()
'
'==============================================================================
'                      TEST ELAPSEDSECONDS ALL METHODS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates basic numeric elapsed-time behavior across all timing methods
'
' WHY THIS EXISTS
'   Numeric elapsed-time retrieval is the central timing output of the class and
'   must behave sensibly across all documented backends
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim iMethod     As Integer                'Timing backend index
    Dim DelayS      As Double                 'Requested delay in seconds
    Dim ElapsedS    As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ElapsedSeconds across all methods"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6

            'Choose a practical per-method delay
                DelayS = DelayForTimingMethod(iMethod)

            'Start a new timing session
                cPM.StartTimer iMethod, False

            'Perform a deliberate pause so the elapsed value should become positive
                cPM.Pause DelayS, 1

            'Read numeric elapsed time
                ElapsedS = cPM.ElapsedSeconds(iMethod)

            'Assert that the numeric elapsed time is nonnegative
                AssertNonNegativeDouble ElapsedS, _
                                        "ElapsedSeconds nonnegative for method " & CStr(iMethod)

            'Assert that the measured value is meaningfully positive relative to the delay
                AssertTrue (ElapsedS >= (DelayS / 4#)), _
                           "ElapsedSeconds meaningfully positive for method " & CStr(iMethod)
        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ElapsedSeconds_AllMethods"

    Resume CleanExit
End Sub

Private Sub Test_ElapsedTime_AllMethods()
'
'==============================================================================
'                        TEST ELAPSEDTIME ALL METHODS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates formatted elapsed-time output across all timing methods
'
' WHY THIS EXISTS
'   ElapsedTime is the public display/reporting companion to ElapsedSeconds and
'   should return readable, semantically complete output for every backend
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim iMethod     As Integer                'Timing backend index
    Dim DelayS      As Double                 'Requested delay in seconds
    Dim TextOut     As String                 'Formatted elapsed-time output

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ElapsedTime across all methods"
    'Create a fresh class instance
        Set cPM = New cPerformanceManager
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6
            'Choose a practical per-method delay
                DelayS = DelayForTimingMethod(iMethod)
            'Start a new timing session
                cPM.StartTimer iMethod, False
            'Perform a deliberate pause
                cPM.Pause DelayS, 1
            'Read formatted elapsed time
                TextOut = cPM.ElapsedTime(iMethod)
            'Assert that the formatted string is non-empty
                AssertTrue (Len(TextOut) > 0), _
                           "ElapsedTime non-empty for method " & CStr(iMethod)
            'Assert that the formatted string contains each documented unit group
                AssertContainsString TextOut, "milliseconds", _
                                     "ElapsedTime contains milliseconds for method " & CStr(iMethod)

                AssertContainsString TextOut, "microseconds", _
                                     "ElapsedTime contains microseconds for method " & CStr(iMethod)

                AssertContainsString TextOut, "nanoseconds", _
                                     "ElapsedTime contains nanoseconds for method " & CStr(iMethod)
        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ElapsedTime_AllMethods"

    Resume CleanExit
End Sub

Private Sub Test_ElapsedTime_FormatExistingSeconds()
'
'==============================================================================
'                TEST ELAPSEDTIME FORMAT EXISTING SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the ElapsedTime path that formats an already measured
'   elapsed-seconds value without taking a second timing sample
'
' WHY THIS EXISTS
'   The class now supports formatting an existing elapsed-seconds value directly.
'   That behavior avoids double measurement in callers that already captured
'   ElapsedSeconds() separately for logging or further numeric processing
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim ElapsedS    As Double                 'Measured elapsed seconds
    Dim TextOut     As String                 'Formatted elapsed-time output
    Dim T2Before    As Double                 'Cached raw end timestamp before formatting
    Dim ETBefore    As Double                 'Cached elapsed seconds before formatting

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ElapsedTime formats an existing elapsed-seconds value"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Start a QPC timing session
        cPM.StartTimer 5, False

    'Perform a short pause
        cPM.Pause 0.05, 1

    'Measure elapsed time numerically once
        ElapsedS = cPM.ElapsedSeconds(5)

    'Capture T2 and ET before formatting-only use
        T2Before = cPM.T2
        ETBefore = cPM.ET

    'Format the existing elapsed-seconds value directly
        TextOut = cPM.ElapsedTime(, ElapsedS)

    'Assert that the formatted string is non-empty
        AssertTrue (Len(TextOut) > 0), "ElapsedTime(, ElapsedSecondsIn) returns non-empty text"

    'Assert that the formatted string contains the documented unit groups
        AssertContainsString TextOut, "milliseconds", "ElapsedTime(, ElapsedSecondsIn) contains milliseconds"
        AssertContainsString TextOut, "microseconds", "ElapsedTime(, ElapsedSecondsIn) contains microseconds"
        AssertContainsString TextOut, "nanoseconds", "ElapsedTime(, ElapsedSecondsIn) contains nanoseconds"

    'Assert that formatting-only use does not take a new timing sample
        AssertApproxDouble T2Before, cPM.T2, 0#, "T2 unchanged when ElapsedTime formats an existing elapsed value"
        AssertApproxDouble ETBefore, cPM.ET, 0#, "ET unchanged when ElapsedTime formats an existing elapsed value"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ElapsedTime_FormatExistingSeconds"

    Resume CleanExit
End Sub

Private Sub Test_AlignedStart_AllMethods()
'
'==============================================================================
'                       TEST ALIGNEDSTART ALL METHODS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates aligned-start timing behavior across all timing methods
'
' WHY THIS EXISTS
'   AlignToNextTick is a specialized benchmark feature and should still behave
'   sanely across all documented backends
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim iMethod     As Integer                'Timing backend index
    Dim DelayS      As Double                 'Requested delay in seconds
    Dim ElapsedS    As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "AlignToNextTick across all methods"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6

            'Choose a practical per-method delay
                DelayS = DelayForTimingMethod(iMethod)

            'Start a new aligned timing session
                cPM.StartTimer iMethod, True

            'Perform a deliberate pause
                cPM.Pause DelayS, 1

            'Read numeric elapsed time
                ElapsedS = cPM.ElapsedSeconds(iMethod)

            'Assert that the aligned elapsed time is nonnegative
                AssertNonNegativeDouble ElapsedS, _
                                        "Aligned ElapsedSeconds nonnegative for method " & CStr(iMethod)

            'Assert that the aligned elapsed time is meaningfully positive
                AssertTrue (ElapsedS >= (DelayS / 4#)), _
                           "Aligned ElapsedSeconds meaningfully positive for method " & CStr(iMethod)
        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_AlignedStart_AllMethods"

    Resume CleanExit
End Sub

Private Sub Test_Accessors_QPC()
'
'==============================================================================
'                           TEST ACCESSORS QPC
'------------------------------------------------------------------------------
' PURPOSE
'   Validates raw/cached accessor behavior after a QPC measurement
'
' WHY THIS EXISTS
'   T1, T2, and ET are explicit inspection/debugging surfaces and should remain
'   coherent with the underlying elapsed measurement
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim ElapsedS    As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Accessors after QPC measurement"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Start a QPC timing session
        cPM.StartTimer 5, False

    'Perform a short pause
        cPM.Pause 0.03, 1

    'Read numeric elapsed time through the public API
        ElapsedS = cPM.ElapsedSeconds(5)

    'Assert that the raw captures are nonnegative
        AssertNonNegativeDouble cPM.T1, "T1 nonnegative after QPC measurement"
        AssertNonNegativeDouble cPM.T2, "T2 nonnegative after QPC measurement"

    'Assert that the raw end capture is not earlier than the raw start capture
        AssertTrue (cPM.T2 >= cPM.T1), "T2 >= T1 after QPC measurement"

    'Assert that ET mirrors the cached elapsed value
        AssertApproxDouble ElapsedS, cPM.ET, 0.000000001, "ET matches ElapsedSeconds after QPC measurement"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Accessors_QPC"

    Resume CleanExit
End Sub

Private Sub Test_StrictMode_ElapsedBeforeStart()
'
'==============================================================================
'                   TEST STRICTMODE ELAPSED BEFORE START
'------------------------------------------------------------------------------
' PURPOSE
'   Validates strict-mode behavior when elapsed time is requested before
'   StartTimer
'
' WHY THIS EXISTS
'   Calling ElapsedSeconds before a timing session exists is a fundamental
'   contract violation that strict mode must reject
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim Dummy       As Double                 'Throwaway target for the failing call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "StrictMode: ElapsedSeconds before StartTimer"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Force strict mode explicitly for clarity
        cPM.StrictMode = True

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next

    'Attempt an invalid elapsed-time read before StartTimer
        Dummy = cPM.ElapsedSeconds

    'Assert that an error was raised
        AssertTrue (Err.Number <> 0), "Strict mode raises when ElapsedSeconds is called before StartTimer"

    'Assert that the error description mentions StartTimer
        AssertContainsString Err.Description, "StartTimer", _
                             "Strict-mode error text mentions StartTimer"

    'Clear the expected error state
        Err.Clear

    'Restore normal error handling for the remainder of the case
        On Error GoTo CleanFail

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_StrictMode_ElapsedBeforeStart"

    Resume CleanExit
End Sub

Private Sub Test_StrictMode_MethodMismatch()
'
'==============================================================================
'                     TEST STRICTMODE METHOD MISMATCH
'------------------------------------------------------------------------------
' PURPOSE
'   Validates strict-mode behavior for explicit elapsed-method mismatch
'
' WHY THIS EXISTS
'   The class is intentionally session-bound. Strict mode must reject attempts
'   to read elapsed time with a method different from the one that started the
'   session
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim Dummy       As Double                 'Throwaway target for the failing call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "StrictMode: explicit elapsed-method mismatch"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Force strict mode explicitly for clarity
        cPM.StrictMode = True

    'Start a session with method 1
        cPM.StartTimer 1, False

    'Perform a short pause so the session is live
        cPM.Pause 0.05, 1

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next

    'Attempt an invalid explicit elapsed read with a mismatched method
        Dummy = cPM.ElapsedSeconds(2)

    'Assert that an error was raised
        AssertTrue (Err.Number <> 0), "Strict mode raises on explicit elapsed-method mismatch"

    'Assert that the error description mentions the method mismatch
        AssertContainsString Err.Description, "does not match", _
                             "Strict-mode mismatch error text is informative"

    'Clear the expected error state
        Err.Clear

    'Restore normal error handling for the remainder of the case
        On Error GoTo CleanFail

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_StrictMode_MethodMismatch"

    Resume CleanExit
End Sub

Private Sub Test_StrictMode_InvalidStartMethod()
'
'==============================================================================
'                  TEST STRICTMODE INVALID START METHOD
'------------------------------------------------------------------------------
' PURPOSE
'   Validates strict-mode behavior for invalid start-method input
'
' WHY THIS EXISTS
'   StartTimer should fail fast in strict mode when the caller passes an invalid
'   method identifier
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM As cPerformanceManager    'Class under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "StrictMode: invalid StartTimer method"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Force strict mode explicitly for clarity
        cPM.StrictMode = True

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next

    'Attempt an invalid StartTimer call
        cPM.StartTimer 99, False

    'Assert that an error was raised
        AssertTrue (Err.Number <> 0), "Strict mode raises on invalid StartTimer method"

    'Assert that the error description mentions invalid timer method
        AssertContainsString Err.Description, "Invalid timer method", _
                             "Strict-mode invalid-start error text is informative"

    'Clear the expected error state
        Err.Clear

    'Restore normal error handling for the remainder of the case
        On Error GoTo CleanFail

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_StrictMode_InvalidStartMethod"

    Resume CleanExit
End Sub

Private Sub Test_NonStrictMode_InvalidStartFallback()
'
'==============================================================================
'               TEST NONSTRICTMODE INVALID START FALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Validates non-strict fallback behavior for invalid start-method input
'
' WHY THIS EXISTS
'   In non-strict mode the class documents that invalid start-method inputs are
'   coerced toward a usable backend rather than immediately rejected
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim DelayS      As Double                 'Requested delay in seconds
    Dim ElapsedS    As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "NonStrictMode: invalid StartTimer fallback"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Force non-strict mode
        cPM.StrictMode = False

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

    'Call StartTimer with an invalid method
        cPM.StartTimer 99, False

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Assert that a session is active after fallback
        AssertEqualBoolean True, cPM.HasActiveSession, _
                           "Non-strict invalid StartTimer still establishes a session"

    'Assert that the resolved active method is valid
        AssertTrue ((cPM.ActiveMethodID = 5) Or (cPM.ActiveMethodID = 2)), _
                   "Non-strict invalid StartTimer resolves to method 5 or 2"

    'Assert that the resolved active method has a non-empty name
        AssertTrue (Len(cPM.MethodName(cPM.ActiveMethodID)) > 0), _
                   "Resolved fallback method has a valid MethodName"

    'Choose a practical delay for the resolved backend
        DelayS = DelayForTimingMethod(cPM.ActiveMethodID)

    'Perform a deliberate pause
        cPM.Pause DelayS, 1

    'Read elapsed time using the active-session path
        ElapsedS = cPM.ElapsedSeconds

    'Assert that the fallback path produces a nonnegative elapsed value
        AssertNonNegativeDouble ElapsedS, _
                                "Non-strict fallback path returns nonnegative elapsed time"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_NonStrictMode_InvalidStartFallback"

    Resume CleanExit
End Sub

Private Sub Test_NonStrictMode_MethodMismatchFallback()
'
'==============================================================================
'            TEST NONSTRICTMODE METHOD MISMATCH FALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Validates non-strict fallback behavior for explicit elapsed-method mismatch
'
' WHY THIS EXISTS
'   In non-strict mode, an explicit elapsed-method mismatch should not raise.
'   Instead, the class should fall back to the active session method where
'   allowed
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim ElapsedS    As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "NonStrictMode: explicit elapsed-method mismatch fallback"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Force non-strict mode
        cPM.StrictMode = False

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

    'Start a session with method 1
        cPM.StartTimer 1, False

    'Perform a short pause
        cPM.Pause 0.05, 1

    'Request elapsed time with an explicit mismatched method
        ElapsedS = cPM.ElapsedSeconds(2)

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Assert that the active method remains the original session method
        AssertEqualLong 1, cPM.ActiveMethodID, _
                        "ActiveMethodID remains unchanged after non-strict mismatch fallback"

    'Assert that the fallback elapsed value is nonnegative
        AssertNonNegativeDouble ElapsedS, _
                                "Non-strict mismatch fallback returns nonnegative elapsed time"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_NonStrictMode_MethodMismatchFallback"

    Resume CleanExit
End Sub

Private Sub Test_OverheadMeasurement_Seconds()
'
'==============================================================================
'                  TEST OVERHEADMEASUREMENT IN SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates numeric overhead-measurement helpers across all methods
'
' WHY THIS EXISTS
'   Benchmark-support helpers are part of the public API and should remain
'   callable and sane even for coarse timing methods
'
' NOTES
'   Coarse timing methods can legitimately report very small or zero overhead
'   values, so this test asserts nonnegativity rather than strict positivity
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim iMethod     As Integer                'Timing backend index
    Dim OverheadS   As Double                 'Measured overhead in seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "OverheadMeasurement_Seconds across all methods"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6

            'Measure average near-empty timing overhead with a modest iteration count
                OverheadS = cPM.OverheadMeasurement_Seconds(iMethod, 25)

            'Assert that the reported overhead is nonnegative
                AssertNonNegativeDouble OverheadS, _
                                        "OverheadMeasurement_Seconds nonnegative for method " & CStr(iMethod)
        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_OverheadMeasurement_Seconds"

    Resume CleanExit
End Sub

Private Sub Test_OverheadMeasurement_Text()
'
'==============================================================================
'                    TEST OVERHEADMEASUREMENT TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Validates formatted overhead-measurement reporting across all methods
'
' WHY THIS EXISTS
'   The text-reporting companion to the numeric overhead helper should remain
'   readable, non-empty, and semantically informative
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim iMethod     As Integer                'Timing backend index
    Dim TextOut     As String                 'Formatted overhead text

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "OverheadMeasurement_Text across all methods"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6

            'Read formatted overhead text for the current backend using a small explicit iteration count
                TextOut = cPM.OverheadMeasurement_Text(iMethod, 25)

            'Assert that the formatted string is non-empty
                AssertTrue (Len(TextOut) > 0), _
                           "OverheadMeasurement_Text non-empty for method " & CStr(iMethod)

            'Assert that the backend label appears in the formatted string
                AssertContainsString TextOut, cPM.MethodName(iMethod), _
                                     "OverheadMeasurement_Text contains backend label for method " & CStr(iMethod)

            'Assert that the formatted string mentions seconds
                AssertContainsString TextOut, "seconds", _
                                     "OverheadMeasurement_Text contains seconds for method " & CStr(iMethod)
        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_OverheadMeasurement_Text"

    Resume CleanExit
End Sub

Private Sub Test_Diagnostics_Properties()
'
'==============================================================================
'                     TEST DIAGNOSTICS PROPERTIES
'------------------------------------------------------------------------------
' PURPOSE
'   Validates diagnostic and informational properties
'
' WHY THIS EXISTS
'   The class exposes several human-readable and numeric diagnostics that are
'   useful for environment inspection, troubleshooting, and documentation
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim TextOut     As String                 'Diagnostic text
    Dim QpcHz       As Double                 'Numeric QPC frequency

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Diagnostic and informational properties"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Read and validate the nominal system tick-interval text
        TextOut = cPM.Get_SystemTickInterval
        AssertTrue (Len(TextOut) > 0), "Get_SystemTickInterval is non-empty"
        AssertContainsString TextOut, "Tick Interval", "Get_SystemTickInterval contains label text"

    'Read and validate the QPC tick-interval text
        TextOut = cPM.QPC_Get_SystemTickInterval
        AssertTrue (Len(TextOut) > 0), "QPC_Get_SystemTickInterval is non-empty"
        AssertContainsString TextOut, "QPC Tick interval", "QPC_Get_SystemTickInterval contains label text"

    'Read and validate the QPC frequency text
        TextOut = cPM.QPC_FrequencyPerSecond
        AssertTrue (Len(TextOut) > 0), "QPC_FrequencyPerSecond is non-empty"
        AssertContainsString TextOut, "QPC Tick frequency", "QPC_FrequencyPerSecond contains label text"

    'Read and validate the numeric QPC frequency
        QpcHz = cPM.QPC_FrequencyPerSecond_Value
        AssertNonNegativeDouble QpcHz, "QPC_FrequencyPerSecond_Value is nonnegative"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Diagnostics_Properties"

    Resume CleanExit
End Sub

Private Sub Test_Pause_Method1()
'
'==============================================================================
'                           TEST PAUSE METHOD 1
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Pause method 1 (Sleep API) using QPC timing
'
' WHY THIS EXISTS
'   Pause method 1 is the lowest-overhead pause path and should produce a delay
'   that is reasonably close to the requested duration
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim ElapsedS    As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Pause method 1 (Sleep)"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Start QPC timing
        cPM.StartTimer 5, False

    'Pause for 1 second using method 1
        cPM.Pause 1, 1

    'Measure elapsed time using QPC
        ElapsedS = cPM.ElapsedSeconds(5)

    'Assert that the measured pause lies within a practical tolerance band
        AssertInRangeDouble 0.8, 1.25, ElapsedS, "Pause method 1 elapsed range"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Pause_Method1"

    Resume CleanExit
End Sub

Private Sub Test_Pause_Method2()
'
'==============================================================================
'                           TEST PAUSE METHOD 2
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Pause method 2 (Timer + DoEvents loop) using QPC timing
'
' WHY THIS EXISTS
'   Pause method 2 is a yielding pause path and should still respect the
'   requested duration within a practical tolerance
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim ElapsedS    As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Pause method 2 (Timer + DoEvents)"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Start QPC timing
        cPM.StartTimer 5, False

    'Pause for 1 second using method 2
        cPM.Pause 1, 2

    'Measure elapsed time using QPC
        ElapsedS = cPM.ElapsedSeconds(5)

    'Assert that the measured pause lies within a practical tolerance band
        AssertInRangeDouble 0.8, 1.25, ElapsedS, "Pause method 2 elapsed range"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Pause_Method2"

    Resume CleanExit
End Sub

Private Sub Test_Pause_Method3()
'
'==============================================================================
'                           TEST PAUSE METHOD 3
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Pause method 3 (Application.Wait) using QPC timing
'
' WHY THIS EXISTS
'   Application.Wait is coarse and should not be expected to behave like a
'   fine-grained pause, but it should still produce a reasonable whole-second
'   delay
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim ElapsedS    As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Pause method 3 (Application.Wait)"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Start QPC timing
        cPM.StartTimer 5, True

    'Pause for 1 second using method 3
        cPM.Pause 1, 3

    'Measure elapsed time using QPC
        ElapsedS = cPM.ElapsedSeconds(5)

    'Assert that the measured pause lies within a broad practical range
        AssertInRangeDouble 0.8, 1.5, ElapsedS, "Pause method 3 elapsed range"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Pause_Method3"

    Resume CleanExit
End Sub

Private Sub Test_Pause_Method4()
'
'==============================================================================
'                           TEST PAUSE METHOD 4
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Pause method 4 (Now + DoEvents loop) using QPC timing
'
' WHY THIS EXISTS
'   The Date/Now loop path is coarser and higher-overhead than Sleep or QPC, but
'   it should still approximate the requested delay within a practical range
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim ElapsedS    As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Pause method 4 (Now + DoEvents)"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Start QPC timing
        cPM.StartTimer 5, True

    'Pause for 1 second using method 4
        cPM.Pause 1, 4

    'Measure elapsed time using QPC
        ElapsedS = cPM.ElapsedSeconds(5)

    'Assert that the measured pause lies within a broad practical range
        AssertInRangeDouble 0.8, 1.5, ElapsedS, "Pause method 4 elapsed range"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Pause_Method4"

    Resume CleanExit
End Sub

Private Sub Test_TW_BlankKeyValidation()
'
'==============================================================================
'                      TEST TW BLANK KEY VALIDATION
'------------------------------------------------------------------------------
' PURPOSE
'   Validates blank-key behavior in the shared TW manager
'
' WHY THIS EXISTS
'   The strengthened TW manager now rejects blank instance keys. This should
'   remain enforced because blank keys would otherwise create collisions or
'   ambiguous shared-session bookkeeping
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "TW manager blank-key validation"

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT BEGINSESSION
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next

    'Attempt an invalid blank-key begin
        PM_TW_BeginSession ""

    'Assert that blank-key begin raises
        AssertTrue (Err.Number <> 0), "PM_TW_BeginSession raises on blank key"

    'Clear the expected error state
        Err.Clear

'------------------------------------------------------------------------------
' ASSERT ENDSESSION
'------------------------------------------------------------------------------
    'Attempt an invalid blank-key end
        PM_TW_EndSession ""

    'Assert that blank-key end raises
        AssertTrue (Err.Number <> 0), "PM_TW_EndSession raises on blank key"

    'Clear the expected error state
        Err.Clear

'------------------------------------------------------------------------------
' ASSERT ISINSTANCEACTIVE
'------------------------------------------------------------------------------
    'Attempt an invalid blank-key activity query
        Call PM_TW_IsInstanceActive("")

    'Assert that blank-key activity query raises
        AssertTrue (Err.Number <> 0), "PM_TW_IsInstanceActive raises on blank key"

    'Clear the expected error state
        Err.Clear

    'Restore normal error handling
        On Error GoTo CleanFail

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Force the shared TW manager to a clean baseline
        PM_TW_EndAllSessions

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_TW_BlankKeyValidation"

    Resume CleanExit
End Sub

Private Sub Test_TW_SingleInstance()
'
'==============================================================================
'                        TEST TW SINGLE INSTANCE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates single-instance TW lifecycle behavior
'
' WHY THIS EXISTS
'   The class publicly exposes TW lifecycle control, and the simplest shared
'   manager behavior must work correctly before overlapping cases are tested
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim Baseline    As T_AppState             'Captured Application baseline

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "TW single-instance lifecycle"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Capture the current Application baseline
        CaptureAppState Baseline

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT (PRECONDITIONS)
'------------------------------------------------------------------------------
    'Assert that the instance starts inactive with zero shared TW sessions
        AssertEqualBoolean False, cPM.TW_IsActive, "TW_IsActive before TW_Turn_OFF"
        AssertEqualLong 0, cPM.TW_ActiveSessionCount, "TW_ActiveSessionCount before TW_Turn_OFF"

'------------------------------------------------------------------------------
' ACTIVATE TW SUPPRESSION
'------------------------------------------------------------------------------
    'Begin TW suppression for the instance with no exemptions
        cPM.TW_Turn_OFF TW_Enum.None

'------------------------------------------------------------------------------
' ASSERT (ACTIVE STATE)
'------------------------------------------------------------------------------
    'Assert that the instance is now active
        AssertEqualBoolean True, cPM.TW_IsActive, "TW_IsActive after TW_Turn_OFF"

    'Assert that exactly one shared TW session is active
        AssertEqualLong 1, cPM.TW_ActiveSessionCount, "TW_ActiveSessionCount after TW_Turn_OFF"

    'Assert forced benchmark/performance-state values
        AssertEqualBoolean False, Application.ScreenUpdating, "ScreenUpdating forced OFF"
        AssertEqualBoolean False, Application.EnableEvents, "EnableEvents forced OFF"
        AssertEqualBoolean False, Application.DisplayAlerts, "DisplayAlerts forced OFF"
        AssertEqualLong xlCalculationManual, Application.Calculation, "Calculation forced MANUAL"
        AssertEqualLong xlWait, Application.Cursor, "Cursor forced WAIT"

'------------------------------------------------------------------------------
' DEACTIVATE TW SUPPRESSION
'------------------------------------------------------------------------------
    'End TW suppression for the instance
        cPM.TW_Turn_ON

'------------------------------------------------------------------------------
' ASSERT (RESTORED STATE)
'------------------------------------------------------------------------------
    'Assert that the instance is now inactive
        AssertEqualBoolean False, cPM.TW_IsActive, "TW_IsActive after TW_Turn_ON"

    'Assert that the shared TW manager is idle again
        AssertEqualLong 0, cPM.TW_ActiveSessionCount, "TW_ActiveSessionCount after TW_Turn_ON"

    'Assert baseline restoration
        AssertEqualBoolean Baseline.ScreenUpdating, Application.ScreenUpdating, "ScreenUpdating restored"
        AssertEqualBoolean Baseline.EnableEvents, Application.EnableEvents, "EnableEvents restored"
        AssertEqualBoolean Baseline.DisplayAlerts, Application.DisplayAlerts, "DisplayAlerts restored"
        AssertEqualLong Baseline.Calculation, Application.Calculation, "Calculation restored"
        AssertEqualLong Baseline.Cursor, Application.Cursor, "Cursor restored"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Force the shared TW manager to a clean baseline
        PM_TW_EndAllSessions

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_TW_SingleInstance"

    Resume CleanExit
End Sub

Private Sub Test_TW_OverlappingInstances()
'
'==============================================================================
'                     TEST TW OVERLAPPING INSTANCES
'------------------------------------------------------------------------------
' PURPOSE
'   Validates overlapping multi-instance TW lifecycle behavior
'
' WHY THIS EXISTS
'   The shared TW manager exists specifically because overlapping class
'   instances must be coordinated safely. This is one of the most important
'   architectural regression surfaces in the project
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM1        As cPerformanceManager    'First class instance
    Dim cPM2        As cPerformanceManager    'Second class instance
    Dim Baseline    As T_AppState             'Captured Application baseline

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "TW overlapping multi-instance lifecycle"

    'Create both class instances
        Set cPM1 = New cPerformanceManager
        Set cPM2 = New cPerformanceManager

    'Capture the current Application baseline
        CaptureAppState Baseline

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ACTIVATE INSTANCE 1
'------------------------------------------------------------------------------
    'Begin TW suppression on the first instance with no exemptions
        cPM1.TW_Turn_OFF TW_Enum.None

    'Assert the shared active-session count after instance 1 begins
        AssertEqualLong 1, cPM1.TW_ActiveSessionCount, "Shared TW count after instance 1 begins"

'------------------------------------------------------------------------------
' ACTIVATE INSTANCE 2
'------------------------------------------------------------------------------
    'Begin TW suppression on the second instance while exempting ScreenUpdating
        cPM2.TW_Turn_OFF TW_Enum.ScreenUpdating

    'Assert the shared active-session count after instance 2 begins
        AssertEqualLong 2, cPM2.TW_ActiveSessionCount, "Shared TW count after instance 2 begins"

    'Assert that ScreenUpdating is still forced OFF because instance 1 still requires it
        AssertEqualBoolean False, Application.ScreenUpdating, "ScreenUpdating remains forced OFF with overlapping sessions"

    'Assert that the remaining shared flags are still forced OFF / MANUAL / WAIT
        AssertEqualBoolean False, Application.EnableEvents, "EnableEvents remains forced OFF with overlapping sessions"
        AssertEqualBoolean False, Application.DisplayAlerts, "DisplayAlerts remains forced OFF with overlapping sessions"
        AssertEqualLong xlCalculationManual, Application.Calculation, "Calculation remains MANUAL with overlapping sessions"
        AssertEqualLong xlWait, Application.Cursor, "Cursor remains WAIT with overlapping sessions"

'------------------------------------------------------------------------------
' END INSTANCE 1
'------------------------------------------------------------------------------
    'End the first instance's TW participation
        cPM1.TW_Turn_ON

    'Assert the shared active-session count after instance 1 ends
        AssertEqualLong 1, cPM2.TW_ActiveSessionCount, "Shared TW count after instance 1 ends"

    'Assert instance-local activity state after instance 1 ends
        AssertEqualBoolean False, cPM1.TW_IsActive, "Instance 1 inactive after TW_Turn_ON"
        AssertEqualBoolean True, cPM2.TW_IsActive, "Instance 2 still active after instance 1 ends"

    'Assert that ScreenUpdating now returns to baseline because the remaining
    'instance exempts that flag
        AssertEqualBoolean Baseline.ScreenUpdating, Application.ScreenUpdating, _
                           "ScreenUpdating restored to baseline when only instance 2 remains"

    'Assert that the remaining flags are still forced by the second instance
        AssertEqualBoolean False, Application.EnableEvents, "EnableEvents still forced OFF by instance 2"
        AssertEqualBoolean False, Application.DisplayAlerts, "DisplayAlerts still forced OFF by instance 2"
        AssertEqualLong xlCalculationManual, Application.Calculation, "Calculation still MANUAL by instance 2"
        AssertEqualLong xlWait, Application.Cursor, "Cursor still WAIT by instance 2"

'------------------------------------------------------------------------------
' END INSTANCE 2
'------------------------------------------------------------------------------
    'End the second instance's TW participation
        cPM2.TW_Turn_ON

    'Assert the shared manager is now idle
        AssertEqualLong 0, cPM2.TW_ActiveSessionCount, "Shared TW count after instance 2 ends"

    'Assert full baseline restoration
        AssertEqualBoolean Baseline.ScreenUpdating, Application.ScreenUpdating, "ScreenUpdating restored after final TW session ends"
        AssertEqualBoolean Baseline.EnableEvents, Application.EnableEvents, "EnableEvents restored after final TW session ends"
        AssertEqualBoolean Baseline.DisplayAlerts, Application.DisplayAlerts, "DisplayAlerts restored after final TW session ends"
        AssertEqualLong Baseline.Calculation, Application.Calculation, "Calculation restored after final TW session ends"
        AssertEqualLong Baseline.Cursor, Application.Cursor, "Cursor restored after final TW session ends"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the first instance
        If Not cPM1 Is Nothing Then
            cPM1.ResetEnvironment
            Set cPM1 = Nothing
        End If

    'Release any environment changes held by the second instance
        If Not cPM2 Is Nothing Then
            cPM2.ResetEnvironment
            Set cPM2 = Nothing
        End If

    'Force the shared TW manager to a clean baseline
        PM_TW_EndAllSessions

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_TW_OverlappingInstances"

    Resume CleanExit
End Sub

Private Sub Test_ResetEnvironment_Idempotent()
'
'==============================================================================
'                   TEST RESETENVIRONMENT IDEMPOTENT
'------------------------------------------------------------------------------
' PURPOSE
'   Validates that ResetEnvironment is safe to call more than once and correctly
'   cleans up active environment changes
'
' WHY THIS EXISTS
'   ResetEnvironment is the explicit cleanup contract for the class. Its
'   idempotence is important for defensive calling patterns and error-handling
'   cleanup paths
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Class under test
    Dim Baseline    As T_AppState             'Captured Application baseline
    Dim Dummy       As Double                 'Throwaway elapsed-time holder

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ResetEnvironment idempotence"

    'Create a fresh class instance
        Set cPM = New cPerformanceManager

    'Capture the current Application baseline
        CaptureAppState Baseline

    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' EXERCISE CLEANUP SURFACES
'------------------------------------------------------------------------------
    'Activate TW suppression for the instance
        cPM.TW_Turn_OFF TW_Enum.None

    'Start method 3 so that timer-resolution activation may occur
        cPM.StartTimer 3, False

    'Perform a short pause and read elapsed time to exercise the method 3 path
        cPM.Pause 0.03, 1
        Dummy = cPM.ElapsedSeconds(3)

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Call the explicit cleanup routine for the first time
        cPM.ResetEnvironment

    'Call the explicit cleanup routine a second time to validate idempotence
        cPM.ResetEnvironment

    'Assert that the instance is no longer active in TW
        AssertEqualBoolean False, cPM.TW_IsActive, "TW_IsActive is False after repeated ResetEnvironment"

    'Assert that the shared TW manager is idle
        AssertEqualLong 0, cPM.TW_ActiveSessionCount, "TW_ActiveSessionCount is 0 after repeated ResetEnvironment"

    'Assert Application baseline restoration
        AssertEqualBoolean Baseline.ScreenUpdating, Application.ScreenUpdating, "ScreenUpdating restored after repeated ResetEnvironment"
        AssertEqualBoolean Baseline.EnableEvents, Application.EnableEvents, "EnableEvents restored after repeated ResetEnvironment"
        AssertEqualBoolean Baseline.DisplayAlerts, Application.DisplayAlerts, "DisplayAlerts restored after repeated ResetEnvironment"
        AssertEqualLong Baseline.Calculation, Application.Calculation, "Calculation restored after repeated ResetEnvironment"
        AssertEqualLong Baseline.Cursor, Application.Cursor, "Cursor restored after repeated ResetEnvironment"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any remaining environment changes held by the instance
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Force the shared TW manager to a clean baseline
        PM_TW_EndAllSessions

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ResetEnvironment_Idempotent"

    Resume CleanExit
End Sub


