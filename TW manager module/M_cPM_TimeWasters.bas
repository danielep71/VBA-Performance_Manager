Attribute VB_Name = "M_cPM_TimeWasters"
Option Explicit

'
'==============================================================================
'
'                               M_PM_TIMEWASTERS
'
'==============================================================================
' PURPOSE
'   Shared, process-wide manager for Excel "time waster" suppression.
'
' WHY THIS EXISTS
'   Application.ScreenUpdating / EnableEvents / DisplayAlerts / Calculation /
'   Cursor are global Excel state, not instance-local state. Therefore:
'
'     - multiple cPerformanceManager instances can overlap,
'     - each instance may request a different exemption mask,
'     - restore logic must be coordinated globally.
'
' DESIGN
'   - The first active session captures the original Application state.
'   - Each active instance registers its own disable-mask.
'   - The effective disable-mask is the OR of all active instance masks.
'   - Whenever a session begins/updates/ends, we recompute the effective state.
'   - When the final session ends, we restore the original baseline exactly once.
'
' UPDATED
'   2026-03-28
'==============================================================================


'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Known TW mask bits.
        Private Const PM_TW_MASK_NONE   As Long = 0
        Private Const PM_TW_MASK_SU     As Long = 1
        Private Const PM_TW_MASK_EE     As Long = 2
        Private Const PM_TW_MASK_DA     As Long = 4
        Private Const PM_TW_MASK_CALC   As Long = 8
        Private Const PM_TW_MASK_CURSOR As Long = 16
        Private Const PM_TW_MASK_ALL    As Long = 31


'------------------------------------------------------------------------------
' PRIVATE SHARED STATE
'------------------------------------------------------------------------------
    'Dictionary:
    '   key   = instance key (string)
    '   item  = disable-mask (Long)
        Private g_TW_Sessions          As Object

    'True once the baseline Application state has been captured.
        Private g_TW_BaselineSaved     As Boolean

    'Saved baseline Application state
        Private g_TW_SU                As Boolean
        Private g_TW_EE                As Boolean
        Private g_TW_DA                As Boolean
        Private g_TW_Calc              As Long
        Private g_TW_Cursor            As Long


Public Sub PM_TW_BeginSession( _
    ByVal InstanceKey As String, _
    Optional ByVal ExceptMask As Long = 0)
'
'==============================================================================
'                             PM_TW_BEGINSESSION
'------------------------------------------------------------------------------
' PURPOSE
'   Starts or updates a TW suppression session for one class instance.
'
' INPUTS
'   InstanceKey
'     Unique key identifying the calling class instance.
'
'   ExceptMask
'     Bitmask of TW flags to EXEMPT (i.e., do not disable these flags).
'
' BEHAVIOR
'   - Captures the baseline Application state on the first active session.
'   - Stores/updates this instance's disable-mask.
'   - Recomputes and applies the effective shared TW state.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Ensure the shared dictionary exists.
        PM_TW_EnsureStore

'------------------------------------------------------------------------------
' CAPTURE BASELINE (FIRST ACTIVE SESSION ONLY)
'------------------------------------------------------------------------------
    'Capture the original Application state only when the first shared session begins.
        If g_TW_Sessions.Count = 0 Then
            PM_TW_SaveBaseline
        End If

'------------------------------------------------------------------------------
' REGISTER / UPDATE INSTANCE MASK
'------------------------------------------------------------------------------
    'Store this instance's requested disable-mask.
        g_TW_Sessions(InstanceKey) = PM_TW_DisableMaskFromExcept(ExceptMask)

'------------------------------------------------------------------------------
' APPLY EFFECTIVE SHARED STATE
'------------------------------------------------------------------------------
    'Recompute the aggregate disable-mask and apply it.
        PM_TW_ApplyEffectiveState PM_TW_AggregateDisableMask()
End Sub

Public Sub PM_TW_EndSession( _
    ByVal InstanceKey As String)
'
'==============================================================================
'                              PM_TW_ENDSESSION
'------------------------------------------------------------------------------
' PURPOSE
'   Ends a TW suppression session for one class instance.
'
' BEHAVIOR
'   - Removes the instance from the shared active-session dictionary.
'   - If sessions remain, recomputes the effective shared state.
'   - If no sessions remain, restores the original baseline exactly once.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Ensure the shared dictionary exists.
        PM_TW_EnsureStore

'------------------------------------------------------------------------------
' REMOVE INSTANCE (IF PRESENT)
'------------------------------------------------------------------------------
    'Remove the calling instance from the active session set.
        If g_TW_Sessions.Exists(InstanceKey) Then
            g_TW_Sessions.Remove InstanceKey
        End If

'------------------------------------------------------------------------------
' RESTORE OR REAPPLY
'------------------------------------------------------------------------------
    'If no sessions remain, restore the original baseline.
        If g_TW_Sessions.Count = 0 Then

            'Restore the original Application state only if we actually captured it.
                If g_TW_BaselineSaved Then
                    PM_TW_ApplyEffectiveState PM_TW_MASK_NONE
                    g_TW_BaselineSaved = False
                End If

            'Release the dictionary to return to a clean idle state.
                Set g_TW_Sessions = Nothing

            Exit Sub
        End If

    'Otherwise, recompute and apply the remaining aggregate disable-mask.
        PM_TW_ApplyEffectiveState PM_TW_AggregateDisableMask()
End Sub

Public Function PM_TW_ActiveCount() As Long
'------------------------------------------------------------------------------
' PM_TW_ACTIVECOUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of currently active shared TW sessions.
'------------------------------------------------------------------------------
    'Ensure the shared dictionary exists.
        PM_TW_EnsureStore

    'Return the active session count.
        PM_TW_ActiveCount = g_TW_Sessions.Count
End Function

Public Function PM_TW_IsInstanceActive( _
    ByVal InstanceKey As String) _
    As Boolean
'------------------------------------------------------------------------------
' PM_TW_ISINSTANCEACTIVE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE if the specified class instance currently has an active shared
'   TW session registered in the global manager.
'------------------------------------------------------------------------------
    'Ensure the shared dictionary exists.
        PM_TW_EnsureStore

    'Return instance activity state.
        PM_TW_IsInstanceActive = g_TW_Sessions.Exists(InstanceKey)
End Function

Public Sub PM_TW_EndAllSessions()
'------------------------------------------------------------------------------
' PM_TW_ENDALLSESSIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Emergency/global reset for development or recovery scenarios.
'
' NOTES
'   This is not the normal lifecycle path. Normal callers should use
'   PM_TW_EndSession for the specific active instance.
'------------------------------------------------------------------------------
    'Restore the original Application baseline if available.
        If g_TW_BaselineSaved Then
            PM_TW_ApplyEffectiveState PM_TW_MASK_NONE
            g_TW_BaselineSaved = False
        End If

    'Release all active-session bookkeeping.
        Set g_TW_Sessions = Nothing
End Sub


Private Sub PM_TW_EnsureStore()
'------------------------------------------------------------------------------
' PM_TW_ENSURESTORE
'------------------------------------------------------------------------------
' PURPOSE
'   Lazily creates the shared session dictionary.
'------------------------------------------------------------------------------
    'Create the shared dictionary only once.
        If g_TW_Sessions Is Nothing Then
            Set g_TW_Sessions = CreateObject("Scripting.Dictionary")
        End If
End Sub

Private Sub PM_TW_SaveBaseline()
'------------------------------------------------------------------------------
' PM_TW_SAVEBASELINE
'------------------------------------------------------------------------------
' PURPOSE
'   Captures the original Application state before any TW suppression is applied.
'------------------------------------------------------------------------------
    'Read the current Application state into shared baseline variables.
        With Application
            g_TW_SU = .ScreenUpdating
            g_TW_EE = .EnableEvents
            g_TW_DA = .DisplayAlerts
            g_TW_Calc = .Calculation
            g_TW_Cursor = .Cursor
        End With

    'Mark the baseline as captured.
        g_TW_BaselineSaved = True
End Sub

Private Function PM_TW_DisableMaskFromExcept( _
    ByVal ExceptMask As Long) _
    As Long
'------------------------------------------------------------------------------
' PM_TW_DISABLEMASKFROMEXCEPT
'------------------------------------------------------------------------------
' PURPOSE
'   Converts an exemption mask into a disable-mask.
'
' EXAMPLE
'   ExceptMask = ScreenUpdating Or EnableEvents
'   => disable all other TW flags except those two
'------------------------------------------------------------------------------
    'Clamp to known bits and invert within the supported TW mask set.
        PM_TW_DisableMaskFromExcept = (PM_TW_MASK_ALL And Not ExceptMask)
End Function

Private Function PM_TW_AggregateDisableMask() As Long
'------------------------------------------------------------------------------
' PM_TW_AGGREGATEDISABLEMASK
'------------------------------------------------------------------------------
' PURPOSE
'   ORs together the disable-masks of all active shared TW sessions.
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K           As Variant    'Dictionary key
    Dim Mask        As Long       'Accumulated disable-mask

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Ensure the shared dictionary exists.
        PM_TW_EnsureStore

'------------------------------------------------------------------------------
' AGGREGATE
'------------------------------------------------------------------------------
    'OR together every active instance's disable-mask.
        For Each K In g_TW_Sessions.Keys
            Mask = (Mask Or CLng(g_TW_Sessions(K)))
        Next K

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the aggregate disable-mask.
        PM_TW_AggregateDisableMask = Mask
End Function

Private Sub PM_TW_ApplyEffectiveState( _
    ByVal DisableMask As Long)
'------------------------------------------------------------------------------
' PM_TW_APPLYEFFECTIVESTATE
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the effective shared TW state using:
'     - the original saved baseline, and
'     - the aggregate disable-mask of all active sessions.
'
' RULE
'   For each flag:
'     - if disabled by any active session => force the "performance" setting
'     - otherwise => restore the original baseline setting
'------------------------------------------------------------------------------
    'Apply the effective state flag-by-flag.
        With Application

            'ScreenUpdating
                If (DisableMask And PM_TW_MASK_SU) <> 0 Then
                    .ScreenUpdating = False
                Else
                    .ScreenUpdating = g_TW_SU
                End If

            'EnableEvents
                If (DisableMask And PM_TW_MASK_EE) <> 0 Then
                    .EnableEvents = False
                Else
                    .EnableEvents = g_TW_EE
                End If

            'DisplayAlerts
                If (DisableMask And PM_TW_MASK_DA) <> 0 Then
                    .DisplayAlerts = False
                Else
                    .DisplayAlerts = g_TW_DA
                End If

            'Calculation
                If (DisableMask And PM_TW_MASK_CALC) <> 0 Then
                    .Calculation = xlCalculationManual
                Else
                    .Calculation = g_TW_Calc
                End If

            'Cursor
                If (DisableMask And PM_TW_MASK_CURSOR) <> 0 Then
                    .Cursor = xlWait
                Else
                    .Cursor = g_TW_Cursor
                End If

        End With
End Sub

