Attribute VB_Name = "M_cPM_Test"
Option Explicit

'
'//////////////////////////////////////////////////////////////////////////////
'                                 TEST
'//////////////////////////////////////////////////////////////////////////////
'
'Copy this routine in a standard module and remove comment block
'______________________________________________________________________________
'
Public Sub Test_PerformanceMonitor()
'------------------------------------------------------------------------------
'DECLARE
'------------------------------------------------------------------------------
    Dim cPM     As cPerformanceManager
    Dim i       As Integer              'Loop counter
    Dim j       As Long                 'Loop counter
'------------------------------------------------------------------------------
'INITIALIZE
'------------------------------------------------------------------------------
    Set cPM = New cPerformanceManager
'------------------------------------------------------------------------------
'TEST CLASS MODULE
'------------------------------------------------------------------------------
    For i = 1 To 20
        cPM.StartTimer (5)
        cPM.Pause 1              '1 second
        Debug.Print "Method " & "5" & " - "; cPM.ElapsedTime(5)
    Next i
'------------------------------------------------------------------------------
'TEST YOUR CODE
'------------------------------------------------------------------------------
    For i = 1 To 6
        cPM.StartTimer (i)
        'Your code here (example)
        For j = 1 To 10000000: Next j   'Empty loop
        Debug.Print cPM.ElapsedTime(i)
    Next i
'------------------------------------------------------------------------------
'TEST OverheadMeasurement_Text
'------------------------------------------------------------------------------
    For i = 1 To 6
        Debug.Print cPM.OverheadMeasurement_Text(i)
    Next i
'------------------------------------------------------------------------------
'TEST TICK INTERVAL
'------------------------------------------------------------------------------
    Debug.Print cPM.Get_SystemTickInterval
'------------------------------------------------------------------------------
'TEST QPC TICK FREQUENCY
'------------------------------------------------------------------------------
    Debug.Print cPM.QPC_FrequencyPerSecond
'------------------------------------------------------------------------------
'TEST QPC TICK FREQUENCY
'------------------------------------------------------------------------------
    Debug.Print cPM.QPC_Get_SystemTickInterval
'------------------------------------------------------------------------------
'TEST METHODS NAME
'------------------------------------------------------------------------------
    For i = 1 To 6
        Debug.Print i & " - " & cPM.MethodName(i)
    Next i
'------------------------------------------------------------------------------
'EXIT
'------------------------------------------------------------------------------
    Set cPM = Nothing
End Sub

Public Sub Test_cPerformanceManager()
'
'==============================================================================
'                         TEST: cPerformanceManager
'------------------------------------------------------------------------------
' PURPOSE
'   Exercises the main functionalities of cPerformanceManager:
'     - StartTimer / ElapsedSeconds / ElapsedTime for all timing backends
'     - AlignToNextTick behavior (sanity check)
'     - OverheadMeasurement_Text() diagnostic
'     - Get_SystemTickInterval / QPC_* informational properties
'     - TW_Turn_OFF / TW_Turn_ON (state save/restore)
'     - Pause() (basic behavior)
'
' NOTES
'   - This is a functional + sanity test, not a rigorous benchmark.
'   - Place this in a STANDARD MODULE (not in the class module).
'   - Ensure the class name is exactly: cPerformanceManager
'
' UPDATED: 2026-01-24
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PM              As cPerformanceManager   'Class under test
    Dim i               As Integer              'Method index 1..6
    Dim S               As String               'Formatted elapsed-time string
    Dim Sec             As Double               'Numeric elapsed seconds
    Dim BeforeSU        As Boolean              'Application state snapshot (ScreenUpdating)
    Dim BeforeEE        As Boolean              'Application state snapshot (EnableEvents)
    Dim BeforeDA        As Boolean              'Application state snapshot (DisplayAlerts)
    Dim BeforeCalc      As Long                 'Application state snapshot (Calculation)
    Dim BeforeCursor    As Long                 'Application state snapshot (Cursor)
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    Debug.Print String$(90, "=")
    Debug.Print "TEST START: cPerformanceManager"
    Debug.Print "Timestamp: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(90, "=")
    Set PM = New cPerformanceManager
'
'------------------------------------------------------------------------------
' INFO PROPERTIES (SMOKE)
'------------------------------------------------------------------------------
    Debug.Print "Get_SystemTickInterval: " & PM.Get_SystemTickInterval
    Debug.Print "QPC_FrequencyPerSecond: " & PM.QPC_FrequencyPerSecond
    Debug.Print "QPC_Get_SystemTickInterval: " & PM.QPC_Get_SystemTickInterval
    Debug.Print String$(90, "-")
'
'------------------------------------------------------------------------------
' MethodName NAMES (INDEXING SANITY)
'------------------------------------------------------------------------------
    Debug.Print "MethodName labels:"
    For i = 1 To 6
        Debug.Print "  " & i & " -> " & PM.MethodName(i)
    Next i
    Debug.Print String$(90, "-")
'
'------------------------------------------------------------------------------
' OverheadMeasurement_Text ESTIMATES
'------------------------------------------------------------------------------
    Debug.Print "OverheadMeasurement_Text estimates:"
    For i = 1 To 6
        Debug.Print "  " & PM.OverheadMeasurement_Text(i)
    Next i
    Debug.Print String$(90, "-")
'
'------------------------------------------------------------------------------
' BASIC TIMING: ElapsedSeconds + ElapsedTime (NO ALIGN)
'------------------------------------------------------------------------------
    Debug.Print "Basic timing (AlignToNextTick = FALSE):"
    For i = 1 To 6
        PM.StartTimer i, False
        PM.Pause 0.05, 1                          '50 ms pause using Sleep path (stable + simple)
        Sec = PM.ElapsedSeconds(i)
        S = CStr(PM.ElapsedTime(i))
        Debug.Print "  Method " & i & " (" & PM.MethodName(i) & ")"
        Debug.Print "    ElapsedSeconds = " & Format$(Sec, "0.000000")
        Debug.Print "    ElapsedTime    = " & S
    Next i
    Debug.Print String$(90, "-")
'
'------------------------------------------------------------------------------
' BASIC TIMING: ALIGN TO NEXT TICK (SANITY)
'------------------------------------------------------------------------------
    Debug.Print "Aligned timing (AlignToNextTick = TRUE):"
    For i = 1 To 6
        PM.StartTimer i, True
        PM.Pause 0.05, 1                          '50 ms pause using Sleep path
        Sec = PM.ElapsedSeconds(i)
        S = CStr(PM.ElapsedTime(i))
        Debug.Print "  Method " & i & " (" & PM.MethodName(i) & ")"
        Debug.Print "    ElapsedSeconds = " & Format$(Sec, "0.000000")
        Debug.Print "    ElapsedTime    = " & S
    Next i
    Debug.Print String$(90, "-")
'
'------------------------------------------------------------------------------
' T1/T2/ET ACCESSORS (SMOKE)
'------------------------------------------------------------------------------
    PM.StartTimer 5, False                         'QPC
    PM.Pause 0.02, 1
    Sec = PM.ElapsedSeconds(5)
    Debug.Print "Accessors (after QPC measurement):"
    Debug.Print "  t1 = " & Format$(PM.T1, "0.000000")
    Debug.Print "  t2 = " & Format$(PM.T2, "0.000000")
    Debug.Print "  ET = " & Format$(PM.ET, "0.000000")
    Debug.Print "  ElapsedSeconds(5) = " & Format$(Sec, "0.000000")
    Debug.Print String$(90, "-")
'
'------------------------------------------------------------------------------
' TW_Turn_OFF / TW_Turn_ON (STATE SAVE/RESTORE)
'------------------------------------------------------------------------------
    Debug.Print "TW_Turn_OFF / TW_Turn_ON state restore test:"
    With Application
        BeforeSU = .ScreenUpdating
        BeforeEE = .EnableEvents
        BeforeDA = .DisplayAlerts
        BeforeCalc = .Calculation
        BeforeCursor = .Cursor
    End With
    PM.TW_Turn_OFF TW_Enum.None
    Debug.Print "  After TW_Turn_OFF:"
    Debug.Print "    ScreenUpdating=" & Application.ScreenUpdating & _
                " EnableEvents=" & Application.EnableEvents & _
                " DisplayAlerts=" & Application.DisplayAlerts & _
                " Calculation=" & Application.Calculation & _
                " Cursor=" & Application.Cursor
    PM.TW_Turn_ON
    Debug.Print "  After TW_Turn_ON (restored):"
    Debug.Print "    ScreenUpdating=" & Application.ScreenUpdating & _
                " EnableEvents=" & Application.EnableEvents & _
                " DisplayAlerts=" & Application.DisplayAlerts & _
                " Calculation=" & Application.Calculation & _
                " Cursor=" & Application.Cursor
    Debug.Print "  Restore OK? " & _
                CStr((Application.ScreenUpdating = BeforeSU) And _
                     (Application.EnableEvents = BeforeEE) And _
                     (Application.DisplayAlerts = BeforeDA) And _
                     (Application.Calculation = BeforeCalc) And _
                     (Application.Cursor = BeforeCursor))
    Debug.Print String$(90, "-")
'
'------------------------------------------------------------------------------
' PAUSE() SANITY (ALL MODES)
'------------------------------------------------------------------------------
    Debug.Print "Pause() sanity (short waits):"
    PM.StartTimer 5, False
    PM.Pause 0.05, 1                               'Sleep API
    Debug.Print "  Pause method 1 (Sleep) : " & PM.ElapsedTime(5)
    PM.StartTimer 5, False
    PM.Pause 0.05, 2                               'Timer+DoEvents loop (coarser)
    Debug.Print "  Pause method 2 (Timer) : " & PM.ElapsedTime(5)
    PM.StartTimer 5, False
    PM.Pause 0.05, 3                               'Application.Wait (whole seconds granularity)
    Debug.Print "  Pause method 3 (Wait)  : " & PM.ElapsedTime(5)
    PM.StartTimer 5, False
    PM.Pause 0.05, 4                               'Now()+DoEvents loop (coarser)
    Debug.Print "  Pause method 4 (Now)   : " & PM.ElapsedTime(5)
    Debug.Print String$(90, "-")
'
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    Set PM = Nothing
    Debug.Print "TEST END: cPerformanceManager"
    Debug.Print String$(90, "=")
End Sub

