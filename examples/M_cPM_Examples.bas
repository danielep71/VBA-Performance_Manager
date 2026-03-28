Attribute VB_Name = "M_cPM_Examples"
'==============================================================================
' MODULE: mPerformanceManager_Examples
'------------------------------------------------------------------------------
' PURPOSE
'   Example, demonstration, and launcher module for cPerformanceManager.
'
' WHY THIS EXISTS
'   This module serves as the user-facing companion to cPerformanceManager by:
'     - showing common usage patterns
'     - demonstrating diagnostic and benchmark helpers
'     - illustrating strict/non-strict behavior
'     - illustrating shared TW control patterns
'     - providing one-click launchers for grouped demonstrations
'
' INPUTS
'   None at module level.
'
' RETURNS
'   None at module level.
'
' BEHAVIOR
'   - Individual public procedures demonstrate a specific feature or usage
'     pattern of cPerformanceManager.
'   - Launcher procedures execute multiple examples in sequence and print clear
'     section markers to the Immediate Window.
'
' ERROR POLICY
'   Individual examples manage their own local error behavior.
'   Group launchers raise normally unless a called example handles errors
'   internally.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Excel Object Model
'   - Shared TW support module(s), for TW-related examples
'
' NOTES
'   - Run examples from a safe workbook because several examples write to cells,
'     formulas, or calculation state.
'   - Results are printed primarily to the Immediate Window.
'   - Press Ctrl+G in the VBA editor to open the Immediate Window.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit     'Force explicit declaration of all variables


'
'------------------------------------------------------------------------------
'
'                           PUBLIC: MASTER LAUNCHERS
'
'------------------------------------------------------------------------------
'

Public Sub Run_All_Examples()
'
'==============================================================================
'                             RUN ALL EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes all demonstration examples in a logical sequence.
'
' WHY THIS EXISTS
'   Provides a single entry point for reviewing the full cPerformanceManager
'   example suite without manually running each procedure one by one.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Prints a module-level start banner.
'   - Runs all example groups in a sensible teaching order.
'   - Prints a module-level completion banner.
'
' ERROR POLICY
'   Raises errors normally unless a called example handles them internally.
'
' DEPENDENCIES
'   - Run_CoreTiming_Examples
'   - Run_StateAndValidation_Examples
'   - Run_DiagnosticsAndBenchmark_Examples
'   - Run_ExecutionControl_Examples
'   - Run_TimeWaster_Examples
'
' NOTES
'   This is the best launcher when you want a complete walkthrough.
'
' UPDATED
'   2026-03-28
'==============================================================================

'Print the overall suite start banner
    PrintModuleBanner "RUN ALL cPerformanceManager EXAMPLES - START"

'Run groups
    Run_CoreTiming_Examples
    Run_StateAndValidation_Examples
    Run_DiagnosticsAndBenchmark_Examples
    Run_ExecutionControl_Examples
    Run_TimeWaster_Examples

'Print the overall suite completion banner.
    PrintModuleBanner "RUN ALL cPerformanceManager EXAMPLES - COMPLETE"

End Sub


Public Sub Run_CoreTiming_Examples()
'
'==============================================================================
'                        RUN CORE TIMING EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the core timing demonstrations.
'
' WHY THIS EXISTS
'   Groups together the examples that most directly show how to time work with
'   cPerformanceManager.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Runs:
'     - Example_BasicTiming_DefaultQPC
'     - Example_ElapsedTime_Text
'     - Example_SpecificMethod
'     - Example_AlignedStart
'
' ERROR POLICY
'   Raises errors normally unless a called example handles them internally.
'
' DEPENDENCIES
'   - Example_BasicTiming_DefaultQPC
'   - Example_ElapsedTime_Text
'   - Example_SpecificMethod
'   - Example_AlignedStart
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "CORE TIMING EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the basic default-QPC timing example
        PrintExampleBanner "Example_BasicTiming_DefaultQPC"
        Example_BasicTiming_DefaultQPC

    'Run the formatted elapsed-time example
        PrintExampleBanner "Example_ElapsedTime_Text"
        Example_ElapsedTime_Text

    'Run the explicit-method example
        PrintExampleBanner "Example_SpecificMethod"
        Example_SpecificMethod

    'Run the aligned-start example
        PrintExampleBanner "Example_AlignedStart"
        Example_AlignedStart
End Sub


Public Sub Run_StateAndValidation_Examples()
'
'==============================================================================
'                  RUN STATE AND VALIDATION EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes examples related to session inspection and strict/non-strict
'   validation behavior.
'
' WHY THIS EXISTS
'   These examples explain how the class behaves internally and how it enforces
'   correct timing-session usage.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Runs:
'     - Example_StateInspection
'     - Example_StrictMode
'     - Example_NonStrictMode
'
' ERROR POLICY
'   Raises errors normally unless a called example handles them internally.
'
' DEPENDENCIES
'   - Example_StateInspection
'   - Example_StrictMode
'   - Example_NonStrictMode
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "STATE / VALIDATION EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the state-inspection example
        PrintExampleBanner "Example_StateInspection"
        Example_StateInspection

    'Run the strict-mode validation example.
        PrintExampleBanner "Example_StrictMode"
        Example_StrictMode

    'Run the non-strict-mode validation example
        PrintExampleBanner "Example_NonStrictMode"
        Example_NonStrictMode
End Sub


Public Sub Run_DiagnosticsAndBenchmark_Examples()
'
'==============================================================================
'               RUN DIAGNOSTICS AND BENCHMARK EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the benchmark-support and diagnostics demonstrations.
'
' WHY THIS EXISTS
'   These examples show how to inspect the timing environment and estimate the
'   cost of the timing framework itself.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Runs:
'     - Example_OverheadMeasurement
'     - Example_Diagnostics
'
' ERROR POLICY
'   Raises errors normally unless a called example handles them internally.
'
' DEPENDENCIES
'   - Example_OverheadMeasurement
'   - Example_Diagnostics
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "DIAGNOSTICS / BENCHMARK EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the overhead-measurement example
        PrintExampleBanner "Example_OverheadMeasurement"
        Example_OverheadMeasurement

    'Run the diagnostics example
        PrintExampleBanner "Example_Diagnostics"
        Example_Diagnostics
End Sub


Public Sub Run_ExecutionControl_Examples()
'
'==============================================================================
'                 RUN EXECUTION CONTROL EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes examples related to pause/wait control and safe cleanup structure.
'
' WHY THIS EXISTS
'   These routines show how cPerformanceManager can support controlled waiting
'   and robust usage patterns in real procedures.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Runs:
'     - Example_Pause
'     - Example_SafePattern
'
' ERROR POLICY
'   Raises errors normally unless a called example handles them internally.
'
' DEPENDENCIES
'   - Example_Pause
'   - Example_SafePattern
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "EXECUTION CONTROL EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the pause-method example
        PrintExampleBanner "Example_Pause"
        Example_Pause

    'Run the structured cleanup pattern example
        PrintExampleBanner "Example_SafePattern"
        Example_SafePattern
End Sub


Public Sub Run_TimeWaster_Examples()
'
'==============================================================================
'                   RUN TIME-WASTER CONTROL EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the examples related to shared Excel "time-waster" suppression.
'
' WHY THIS EXISTS
'   These examples show how to benchmark worksheet/application work with Excel
'   noise reduced through the shared TW manager model.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Runs:
'     - Example_TimeWasters_Basic
'     - Example_TimeWasters_WithExceptions
'
' ERROR POLICY
'   Raises errors normally unless a called example handles them internally.
'
' DEPENDENCIES
'   - Example_TimeWasters_Basic
'   - Example_TimeWasters_WithExceptions
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner.
        PrintSectionBanner "TIME-WASTER CONTROL EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the basic TW-suppression example
        PrintExampleBanner "Example_TimeWasters_Basic"
        Example_TimeWasters_Basic

    'Run the TW-suppression-with-exceptions example
        PrintExampleBanner "Example_TimeWasters_WithExceptions"
        Example_TimeWasters_WithExceptions
End Sub


'
'------------------------------------------------------------------------------
'
'                       PRIVATE: OUTPUT / BANNER HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub PrintModuleBanner(ByVal Title As String)
'
'==============================================================================
'                           PRINT MODULE BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints a visually distinct top-level banner to the Immediate Window.
'
' WHY THIS EXISTS
'   Improves readability when running large groups of examples.
'
' INPUTS
'   Title
'     Banner title text.
'
' RETURNS
'   None.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' PRINT BANNER
'------------------------------------------------------------------------------
    Debug.Print vbNullString
    Debug.Print String$(78, "=")
    Debug.Print Title
    Debug.Print String$(78, "=")
End Sub


Private Sub PrintSectionBanner(ByVal Title As String)
'
'==============================================================================
'                           PRINT SECTION BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints a section-level banner to the Immediate Window.
'
' WHY THIS EXISTS
'   Makes grouped example runs easier to read and review.
'
' INPUTS
'   Title
'     Section title text.
'
' RETURNS
'   None.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' PRINT BANNER
'------------------------------------------------------------------------------
    Debug.Print vbNullString
    Debug.Print String$(78, "-")
    Debug.Print Title
    Debug.Print String$(78, "-")
End Sub


Private Sub PrintExampleBanner(ByVal ProcName As String)
'
'==============================================================================
'                           PRINT EXAMPLE BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints a small marker identifying the example about to run.
'
' WHY THIS EXISTS
'   Helps the user associate Immediate Window output with the procedure that
'   produced it.
'
' INPUTS
'   ProcName
'     Name of the example procedure being executed.
'
' RETURNS
'   None.
'
' UPDATED
'   2026-03-28
'==============================================================================

'Print the example marker
    Debug.Print ">>> " & ProcName
End Sub


'
'------------------------------------------------------------------------------
'
'                     PUBLIC: CORE TIMING USAGE EXAMPLES
'
'------------------------------------------------------------------------------
'

Public Sub Example_BasicTiming_DefaultQPC()
'
'==============================================================================
'                       EXAMPLE: BASIC TIMING (DEFAULT QPC)
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates the simplest recommended usage pattern for cPerformanceManager
'   by timing a small piece of worksheet work with the default timing backend.
'
' WHY THIS EXISTS
'   This is the canonical "getting started" example:
'     - instantiate the class
'     - start a timer
'     - perform work
'     - read elapsed seconds
'     - clean up
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a new cPerformanceManager instance.
'   - Starts timing with the default method, which is method 5 (QPC)
'   - Writes a constant into a worksheet range.
'   - Reads numeric elapsed time in seconds.
'   - Prints the result to the Immediate Window.
'   - Restores environment state and releases the instance.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Range
'   - Debug.Print
'
' NOTES
'   This is the preferred example to show the normal benchmark path.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Performance manager instance
    Dim ElapsedS    As Double                 'Elapsed time in seconds
    Dim ElapsedT    As String                 'Elapsed time as string

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing with the default method (5 = QPC)
        cPM.StartTimer

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Write a constant value into a worksheet range
        Range("A1:A10000").Value = 1

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Read elapsed time (choose one of the two)
        ElapsedS = cPM.ElapsedSeconds
        ElapsedT = cPM.ElapsedTime

    'Print the result
        Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")
        Debug.Print "Elapsed time: " & ElapsedT

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        cPM.ResetEnvironment
    'Release the class instance
        Set cPM = Nothing
End Sub


Public Sub Example_ElapsedTime_Text()
'
'==============================================================================
'                       EXAMPLE: FORMATTED ELAPSED TIME
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how to read the human-readable elapsed-time string returned by
'   ElapsedTime()
'
' WHY THIS EXISTS
'   Numeric elapsed seconds are ideal for code and benchmarks, but logs,
'   diagnostics, and demonstrations often benefit from a display-oriented text
'   result.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a timing manager instance.
'   - Starts timing explicitly with method 5 (QPC)
'   - Performs a calculation workload.
'   - Reads the formatted elapsed-time string.
'   - Prints the formatted result to the Immediate Window.
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Application.Calculate
'   - Debug.Print
'
' NOTES
'   ElapsedTime() is intentionally heavier than ElapsedSeconds() because it
'   formats a presentation-oriented string.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM  As cPerformanceManager    'Performance manager instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing explicitly with QPC
        cPM.StartTimer 5

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Force a workbook/application calculation pass
        Application.Calculate

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Read the formatted elapsed-time string
        Debug.Print cPM.ElapsedTime

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        cPM.ResetEnvironment
    'Release the class instance
        Set cPM = Nothing
End Sub


Public Sub Example_SpecificMethod()
'
'==============================================================================
'                          EXAMPLE: SPECIFIC METHOD
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how to start a timing session with an explicitly selected
'   timing backend.
'
' WHY THIS EXISTS
'   cPerformanceManager supports multiple timing methods. This example shows
'   that the class can be pointed at a chosen backend and then queried for the
'   backend name and elapsed result.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Randomly selects one of the six supported timing methods.
'   - Starts a session using that method.
'   - Performs a small worksheet operation.
'   - Prints the method name used.
'   - Prints the measured elapsed seconds.
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Randomize
'   - Rnd
'   - Worksheets
'   - Debug.Print
'
' NOTES
'   - This is a demonstration example, not a deterministic benchmark.
'   - For reproducible tests, use a fixed Idx instead of a random one.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM     As cPerformanceManager    'Performance manager instance
    Dim Idx     As Integer                'Selected timing method identifier

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' SELECT METHOD
'------------------------------------------------------------------------------
    'Initialize the VBA random-number generator
        Randomize

    'Select a random method in the inclusive range 1..6
        Idx = Int(6 * Rnd) + 1

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start a timing session using the selected method
        cPM.StartTimer Idx

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Write a volatile worksheet formula
        Worksheets(1).Range("A1").Formula = "=RAND()"

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Print the readable method label
        Debug.Print "Method used: " & cPM.MethodName(cPM.ActiveMethodID)
    'Print the measured elapsed seconds
        Debug.Print "Elapsed seconds: " & Format$(cPM.ElapsedSeconds, "0.000000000")

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        cPM.ResetEnvironment
    'Release the class instance
        Set cPM = Nothing
End Sub


Public Sub Example_AlignedStart()
'
'==============================================================================
'                          EXAMPLE: ALIGNED START
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates the use of AlignToNextTick when starting a timing session.
'
' WHY THIS EXISTS
'   For certain very small benchmark scenarios, aligning the start capture to
'   the next observable tick can reduce start-capture quantization jitter.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a timing manager instance.
'   - Starts timing with QPC and alignment enabled.
'   - Executes a tiny workload.
'   - Prints the aligned elapsed seconds.
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - DoEvents
'   - Debug.Print
'
' NOTES
'   AlignToNextTick is useful mainly for micro-benchmark experiments.
'   It should not be the default path for ordinary application timing.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM  As cPerformanceManager    'Performance manager instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start QPC timing aligned to the next observable tick
        cPM.StartTimer 5, True

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Run a very small operation
        DoEvents

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Print the aligned elapsed-time measurement
        Debug.Print "Aligned elapsed seconds: " & _
                    Format$(cPM.ElapsedSeconds, "0.000000000")

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        cPM.ResetEnvironment
    'Release the class instance
        Set cPM = Nothing
End Sub


Public Sub Example_StateInspection()
'
'==============================================================================
'                         EXAMPLE: STATE INSPECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how to inspect the session state and raw timing values exposed
'   by the class.
'
' WHY THIS EXISTS
'   The class intentionally exposes several internal state surfaces for
'   diagnostics and teaching:
'     - HasActiveSession
'     - ActiveMethodID
'     - MethodName
'     - T1
'     - T2
'     - ET
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Starts a QPC timing session.
'   - Prints session metadata and the raw start timestamp.
'   - Pauses briefly to create measurable elapsed time.
'   - Prints elapsed seconds, raw end timestamp, and cached elapsed time.
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   This is primarily a teaching/debugging example, not a production pattern.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM  As cPerformanceManager    'Performance manager instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing with QPC
        cPM.StartTimer 5

'------------------------------------------------------------------------------
' INSPECT INITIAL STATE
'------------------------------------------------------------------------------
    'Print whether the session is active
        Debug.Print "HasActiveSession = "; cPM.HasActiveSession

    'Print the active method identifier
        Debug.Print "ActiveMethodID   = "; cPM.ActiveMethodID

    'Print the human-readable method name
        Debug.Print "MethodName       = "; cPM.MethodName(cPM.ActiveMethodID)

    'Print the raw start timestamp
        Debug.Print "T1               = "; cPM.T1

'------------------------------------------------------------------------------
' APPLY SMALL DELAY
'------------------------------------------------------------------------------
    'Pause briefly to produce a non-trivial elapsed reading
        cPM.Pause 0.05, 1

'------------------------------------------------------------------------------
' INSPECT FINAL STATE
'------------------------------------------------------------------------------
    'Print numeric elapsed seconds
        Debug.Print "ElapsedSeconds   = "; cPM.ElapsedSeconds
    'Print the raw end timestamp
        Debug.Print "T2               = "; cPM.T2
    'Print the cached elapsed time
        Debug.Print "ET               = "; cPM.ET

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        cPM.ResetEnvironment
    'Release the class instance
        Set cPM = Nothing
End Sub


Public Sub Example_StrictMode()
'
'==============================================================================
'                           EXAMPLE: STRICT MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how StrictMode enforces correct method/session usage.
'
' WHY THIS EXISTS
'   A key design feature of cPerformanceManager is that elapsed reads are
'   session-bound. In strict mode, an invalid elapsed-method request raises an
'   error rather than being silently coerced.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a timing manager instance.
'   - Enables strict mode.
'   - Starts a session using method 5.
'   - Intentionally requests elapsed time with method 2.
'   - Captures and prints the raised error.
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Uses local On Error Resume Next to demonstrate the raised error safely.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'   - Err
'
' NOTES
'   This example is intentionally invalid by design.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM  As cPerformanceManager    'Performance manager instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' CONFIGURE
'------------------------------------------------------------------------------
    'Enable strict validation behavior
        cPM.StrictMode = True

    'Start timing with method 5
        cPM.StartTimer 5

'------------------------------------------------------------------------------
' APPLY DEMONSTRATION ERROR HANDLING
'------------------------------------------------------------------------------
    'Use local error suppression so the example can inspect the raised error
        On Error Resume Next

'------------------------------------------------------------------------------
' TRIGGER INVALID USAGE
'------------------------------------------------------------------------------
    'This is intentionally invalid because the active session uses method 5
        Debug.Print cPM.ElapsedSeconds(2)

'------------------------------------------------------------------------------
' INSPECT ERROR
'------------------------------------------------------------------------------
    'Print the error information if strict mode raised as expected
        If Err.Number <> 0 Then
            Debug.Print "Error number: " & Err.Number
            Debug.Print "Error text  : " & Err.Description
            Err.Clear
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Return to normal error handling semantics
        On Error GoTo 0

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        cPM.ResetEnvironment
    'Release the class instance
        Set cPM = Nothing
End Sub


Public Sub Example_NonStrictMode()
'
'==============================================================================
'                         EXAMPLE: NON-STRICT MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how the class behaves in non-strict mode when the caller asks
'   for an elapsed method that does not match the active session.
'
' WHY THIS EXISTS
'   Non-strict mode is the forgiving mode of the class. Instead of raising, the
'   class may coerce the request and continue using the active session method.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a timing manager instance.
'   - Disables strict mode.
'   - Starts a session with method 5.
'   - Waits briefly.
'   - Requests elapsed time using the wrong method identifier.
'   - Prints the returned value, which should come from the active session.
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Raises errors normally, though this example is expected not to raise.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   This example contrasts directly with Example_StrictMode.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM  As cPerformanceManager    'Performance manager instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' CONFIGURE
'------------------------------------------------------------------------------
    'Disable strict validation behavior
        cPM.StrictMode = False

    'Start timing with method 5.
        cPM.StartTimer 5

'------------------------------------------------------------------------------
' APPLY SMALL DELAY
'------------------------------------------------------------------------------
    'Pause briefly to make the elapsed reading visible
        cPM.Pause 0.03, 1

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'In non-strict mode this falls back to the active session method
        Debug.Print cPM.ElapsedSeconds(2)

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        cPM.ResetEnvironment
    'Release the class instance
        Set cPM = Nothing
End Sub


'
'------------------------------------------------------------------------------
'
'                   PUBLIC: DIAGNOSTICS / BENCHMARK EXAMPLES
'
'------------------------------------------------------------------------------
'

Public Sub Example_OverheadMeasurement()
'
'==============================================================================
'                      EXAMPLE: OVERHEAD MEASUREMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates both the numeric and text-based overhead measurement helpers.
'
' WHY THIS EXISTS
'   Measuring the timing framework's own near-empty cost helps the user
'   understand the approximate cost of a minimal timing cycle through the class.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a timing manager instance.
'   - Measures average near-empty timing overhead in seconds.
'   - Prints the numeric result.
'   - Prints the formatted text result.
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   This measures the class timing path, not only the raw operating-system API.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM         As cPerformanceManager    'Performance manager instance
    Dim OverheadS   As Double                 'Measured average overhead in seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' MEASURE OVERHEAD
'------------------------------------------------------------------------------
    'Measure average near-empty timing overhead using QPC over 1000 iterations
        OverheadS = cPM.OverheadMeasurement_Seconds(5, 1000)

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Print the numeric overhead value
        Debug.Print "Numeric overhead: " & Format$(OverheadS, "0.000000000")

    'Print the formatted overhead report
        Debug.Print cPM.OverheadMeasurement_Text(5)

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        cPM.ResetEnvironment
    'Release the class instance.
        Set cPM = Nothing
End Sub


Public Sub Example_Diagnostics()
'
'==============================================================================
'                           EXAMPLE: DIAGNOSTICS
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates the diagnostic properties exposed by cPerformanceManager.
'
' WHY THIS EXISTS
'   The class exposes both system-level and QPC-level diagnostic information to
'   help the caller understand the timing environment in which measurements are
'   being taken.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a timing manager instance.
'   - Prints the nominal system tick interval.
'   - Prints the QPC tick interval.
'   - Prints the QPC frequency as text.
'   - Prints the QPC frequency as a numeric value.
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   These values are diagnostic/contextual, not direct elapsed measurements.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM  As cPerformanceManager    'Performance manager instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance.
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' READ DIAGNOSTICS
'------------------------------------------------------------------------------
    'Print the nominal system clock tick interval.
        Debug.Print cPM.Get_SystemTickInterval

    'Print the QPC tick interval.
        Debug.Print cPM.QPC_Get_SystemTickInterval

    'Print the QPC frequency as formatted text.
        Debug.Print cPM.QPC_FrequencyPerSecond

    'Print the QPC frequency as a numeric value.
        Debug.Print cPM.QPC_FrequencyPerSecond_Value

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance.
        cPM.ResetEnvironment

    'Release the class instance.
        Set cPM = Nothing
End Sub


'
'------------------------------------------------------------------------------
'
'                 PUBLIC: EXECUTION CONTROL / WAITING EXAMPLES
'
'------------------------------------------------------------------------------
'

Public Sub Example_Pause()
'
'==============================================================================
'                              EXAMPLE: PAUSE
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates the different pause strategies supported by the class.
'
' WHY THIS EXISTS
'   cPerformanceManager includes a small utility pause method that supports
'   multiple waiting behaviors with different trade-offs.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a timing manager instance.
'   - Executes pause method 1 (Sleep)
'   - Executes pause method 2 (Timer + DoEvents)
'   - Executes pause method 3 (Application.Wait)
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - cPerformanceManager
'
' NOTES
'   This example demonstrates waiting behavior, not elapsed-time measurement.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM  As cPerformanceManager    'Performance manager instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance.
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' APPLY PAUSE METHODS
'------------------------------------------------------------------------------
    'Pause using method 1 = Sleep.
        cPM.Pause 0.25, 1

    'Pause using method 2 = Timer + DoEvents.
        cPM.Pause 0.25, 2

    'Pause using method 3 = Application.Wait.
        cPM.Pause 1, 3

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance.
        cPM.ResetEnvironment

    'Release the class instance.
        Set cPM = Nothing
End Sub


'
'------------------------------------------------------------------------------
'
'              PUBLIC: SHARED TIME-WASTER SUPPRESSION EXAMPLES
'
'------------------------------------------------------------------------------
'

Public Sub Example_TimeWasters_Basic()
'
'==============================================================================
'                     EXAMPLE: TIMEWASTERS (BASIC)
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates basic shared Excel "time-waster" suppression for a benchmark
'   run.
'
' WHY THIS EXISTS
'   Excel application behaviors such as ScreenUpdating, events, alerts,
'   calculation mode, and cursor changes can add noise to benchmarks. This
'   example shows the normal suppression pattern.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a timing manager instance.
'   - Turns off all supported TW settings for this shared scope.
'   - Starts QPC timing.
'   - Performs a worksheet workload.
'   - Prints elapsed seconds.
'   - Turns TW back on for this instance.
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - TW shared manager module(s)
'   - Range
'   - Debug.Print
'
' NOTES
'   TW control is shared/global in effect, so this relies on the shared
'   manager model rather than direct instance-local restoration.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM  As cPerformanceManager    'Performance manager instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance.
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' SUPPRESS TIME-WASTERS
'------------------------------------------------------------------------------
    'Disable all supported TW settings for this instance's shared session.
        cPM.TW_Turn_OFF

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing with QPC.
        cPM.StartTimer 5

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Execute a worksheet workload for benchmarking.
        Range("A1:A50000").Formula = "=ROW()"

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Print elapsed seconds with TW suppression in effect.
        Debug.Print "Elapsed seconds with TW off: " & cPM.ElapsedSeconds

'------------------------------------------------------------------------------
' RESTORE TW PARTICIPATION
'------------------------------------------------------------------------------
    'End this instance's shared TW suppression session.
        cPM.TW_Turn_ON

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any remaining environment changes made by this instance.
        cPM.ResetEnvironment

    'Release the class instance.
        Set cPM = Nothing
End Sub


Public Sub Example_TimeWasters_WithExceptions()
'
'==============================================================================
'                EXAMPLE: TIMEWASTERS (WITH EXCEPTIONS)
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates TW suppression while preserving selected Excel application
'   behaviors.
'
' WHY THIS EXISTS
'   In some scenarios the caller wants most benchmark noise suppressed while
'   deliberately leaving a subset of settings active.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a timing manager instance.
'   - Starts TW suppression but keeps ScreenUpdating and EnableEvents active.
'   - Starts QPC timing.
'   - Performs a calculation workload.
'   - Prints the formatted elapsed time.
'   - Ends the TW session.
'   - Cleans up and releases the instance.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - TW_Enum
'   - Application.CalculateFull
'   - Debug.Print
'
' NOTES
'   This example is useful for showing the TW_Enum bitmask behavior.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM  As cPerformanceManager    'Performance manager instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance.
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' SUPPRESS TIME-WASTERS WITH EXCEPTIONS
'------------------------------------------------------------------------------
    'Keep ScreenUpdating and EnableEvents ON while suppressing the remaining TWs.
        cPM.TW_Turn_OFF TW_Enum.ScreenUpdating Or TW_Enum.EnableEvents

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing with QPC.
        cPM.StartTimer 5

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Execute a calculation workload.
        Application.CalculateFull

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Print the formatted elapsed result.
        Debug.Print cPM.ElapsedTime

'------------------------------------------------------------------------------
' RESTORE TW PARTICIPATION
'------------------------------------------------------------------------------
    'End this instance's shared TW suppression session.
        cPM.TW_Turn_ON

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any remaining environment changes made by this instance.
        cPM.ResetEnvironment

    'Release the class instance.
        Set cPM = Nothing
End Sub


'
'------------------------------------------------------------------------------
'
'                     PUBLIC: RECOMMENDED SAFETY PATTERN
'
'------------------------------------------------------------------------------
'

Public Sub Example_SafePattern()
'
'==============================================================================
'                         EXAMPLE: SAFE CLEANUP PATTERN
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates the recommended structured pattern for using
'   cPerformanceManager safely in real procedures.
'
' WHY THIS EXISTS
'   Benchmarks and TW suppression can modify environment state. A structured
'   cleanup block ensures that environment restoration still happens when the
'   workload raises an error.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a timing manager instance.
'   - Starts shared TW suppression.
'   - Starts QPC timing.
'   - Performs a workload.
'   - Prints the formatted elapsed time.
'   - Uses a cleanup label to ensure ResetEnvironment and object release happen
'     even if an error occurs.
'
' ERROR POLICY
'   Uses structured local error handling with CleanFail / CleanExit labels.
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Worksheets
'   - Debug.Print
'
' NOTES
'   This is the best example to follow when integrating the class into real
'   project code.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM  As cPerformanceManager    'Performance manager instance

'------------------------------------------------------------------------------
' INITIALIZE ERROR HANDLING
'------------------------------------------------------------------------------
    'Route runtime failures to the cleanup-aware failure block.
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create a new timing manager instance.
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' CONFIGURE BENCHMARK ENVIRONMENT
'------------------------------------------------------------------------------
    'Start shared TW suppression for this instance.
        cPM.TW_Turn_OFF

    'Start timing with QPC.
        cPM.StartTimer 5

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Execute a workbook workload.
        Worksheets(1).UsedRange.Calculate

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Print the formatted elapsed-time report.
        Debug.Print "Elapsed: " & cPM.ElapsedTime

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release environment changes and the instance if the object exists.
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Exit normally after cleanup.
        Exit Sub

'------------------------------------------------------------------------------
' FAILURE EXIT
'------------------------------------------------------------------------------
CleanFail:
    'Print the error information for diagnostics.
        Debug.Print "Error " & Err.Number & " - " & Err.Description

    'Always route through the normal cleanup block.
        Resume CleanExit
End Sub

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

'Here is what each example is doing and why it matters.
'
'Example_BasicTiming_DefaultQPC is the normal entry point. It shows the _
'    recommended way to use the class for ordinary measurement: create the object, _
'    start timing with the default method, run some work, read ElapsedSeconds, then _
'    call ResetEnvironment. The default method is QPC, so this is effectively the _
'    “best practice” example.
'
'Example_ElapsedTime_Text shows the difference between machine-readable and _
'    presentation-readable output. ElapsedSeconds is what you use in code, _
'    comparisons, assertions, benchmark logs, and averages. ElapsedTime is for people _
'    reading the result. It gives a formatted duration string, which is useful in _
'    demos, diagnostics, and ad hoc inspection.
'
'Example_SpecificMethod teaches that the class is not tied to one backend. It can _
'    run against any of the six timing methods. That matters because sometimes you _
'    want to compare behavior across methods, or you may want to test how resilient _
'    your timing logic is when not using QPC. In this example, the method is chosen _
'    randomly just to demonstrate flexibility. In serious testing, you would replace _
'    the random choice with a fixed method.
'
'Example_AlignedStart demonstrates a specialized benchmarking concept: _
'    start-alignment to the next observable tick. This is not important for ordinary _
'    timing of meaningful workloads. It becomes relevant when the measured operation _
'    is so small that start-capture quantization error can distort the result. In _
'    other words, this example is about benchmarking technique, not standard _
'    application timing.
'
'Example_StateInspection is a teaching and debugging example. It shows that the _
'    class exposes not just the final elapsed value, but also session metadata and _
'    raw internal captures. HasActiveSession tells you whether a session exists. _
'    ActiveMethodID and MethodName tell you what backend is active. T1 is the raw _
'    start capture, T2 is the raw end capture, and ET is the cached elapsed time. _
'    This is especially useful when you want to verify that the class is behaving _
'    exactly as expected.
'
'Example_StrictMode is about correctness enforcement. The class is designed so _
'    that a session started with one method should be read back using that same _
'    method. In strict mode, violating that rule raises an error. This protects the _
'    caller from subtle misuse. The example intentionally triggers invalid usage to _
'    show how the class reacts.
'
'Example_NonStrictMode is the contrast to strict mode. In non-strict mode, the _
'    class is more forgiving. Instead of raising an error when the caller asks for _
'    elapsed time with the wrong method identifier, it coerces the request back to _
'    the active session method where possible. This mode is useful when you want _
'    resilience instead of fail-fast validation.
'
'Example_OverheadMeasurement explains a very important benchmarking concept: the _
'    timer framework itself has a cost. Measuring that cost gives you a rough sense _
'    of how much noise is introduced by the act of measuring. This matters especially _
'    for micro-benchmarks, where the measured operation may be on the same order of _
'    magnitude as the measurement overhead itself.
'
'Example_Diagnostics is not measuring work. It is measuring the measurement _
'    environment. Get_SystemTickInterval gives a coarse system-level clock increment. _
'    QPC_Get_SystemTickInterval and QPC_FrequencyPerSecond tell you about the QPC _
'    timebase itself. These diagnostics help you understand why one method is coarser _
'    than another and why QPC is usually preferred.
'
'Example_Pause shows that the class is also offering controlled waiting behavior. _
'    The three pause modes reflect different trade-offs. Sleep is low-overhead and _
'    direct. Timer + DoEvents yields control and keeps Excel more responsive, but at _
'    the cost of more noise and re-entrancy exposure. Application.Wait is simple and _
'    coarse. This example is useful because it shows that the class is not only about _
'    measurement but also about controlled timing behavior.
'
'Example_TimeWasters_Basic introduces one of the most practical performance _
'    features in Excel benchmarking: suppression of expensive application behaviors. _
'    Screen updating, events, alerts, calculation changes, and cursor updates can _
'    pollute benchmarks badly. This example shows the simple pattern: turn them off, _
'    run the benchmark, turn them back on, and clean up.
'
'Example_TimeWasters_WithExceptions is the more refined version. It shows that TW _
'    suppression is not all-or-nothing. You can keep some behaviors on while _
'    suppressing others. That is useful when, for example, you still want events or _
'    visible screen behavior during a controlled benchmark.
'
'Example_SafePattern is the most important integration example. It demonstrates _
'    the production-safe structure: initialize the object, do the work, and always _
'    funnel through a cleanup block. This is critical because timing helpers and TW _
'    suppression can change environment state. Even if an error occurs, the cleanup _
'    block ensures that the environment is restored and the object is released _
'    properly.


