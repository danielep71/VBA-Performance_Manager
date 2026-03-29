# cPerformanceManager

High-precision timing and benchmark-support utility for VBA on Windows.

`cPerformanceManager` provides a single, session-bound interface for multiple timing backends, numeric elapsed-time measurement, human-readable elapsed-time diagnostics, benchmark overhead measurement, pause helpers, and shared Excel “time-waster” suppression for cleaner benchmark runs.

---

## Overview

VBA’s built-in timing options are often not ideal for instrumentation and benchmarking:

- `Timer` has limited resolution and rolls over at midnight
- `Now()` is wall-clock based and is not a preferred monotonic benchmark source
- Windows exposes better monotonic and high-resolution counters such as `QueryPerformanceCounter`

`cPerformanceManager` wraps those timing sources behind a consistent API and adds benchmark-oriented helpers for Excel/VBA projects.

---

## Features

- Multiple timing methods under one interface
- Session-bound timing model
- Low-overhead numeric elapsed-time measurement
- Human-readable elapsed-time formatting
- Benchmark overhead measurement helpers
- Timing/source diagnostics
- Pause/wait helpers
- Shared Excel “time-waster” suppression for benchmark runs
- Strict-mode validation for safer usage

---

## Timer methods

The class supports the following timing backends:

| ID | Method | Notes |
|---:|---|---|
| 1 | `Timer` | Seconds since midnight; rolls over every 24 hours |
| 2 | `GetTickCount / GetTickCount64` | Milliseconds since boot; 32-bit path wraps at about 49.7 days |
| 3 | `timeGetTime` | Millisecond counter with 32-bit rollover semantics; can use 1ms timer resolution |
| 4 | `timeGetSystemTime (MMTIME / TIME_MS)` | Millisecond source treated with 32-bit rollover semantics |
| 5 | `QueryPerformanceCounter (QPC)` | Default and recommended high-resolution benchmark source |
| 6 | `Now() * 86400` | Wall-clock seconds; useful mainly for diagnostics |

---

## Requirements

- Microsoft Excel with VBA enabled
- Windows host environment for API-backed timing methods (`2..5`)
- `VBA7` / `Win64` conditional-compilation support as required by the host
- The following source files:
  - `cPerformanceManager.cls`
  - `M_cPM_TimeWasters.bas`

### Compatibility note

On non-Windows hosts, only `Timer()` / `Now()`-based methods are conceptually portable, if retained.

---

## Required companion module

This project includes a required companion standard module:

- `M_cPM_TimeWasters.bas`

That module manages shared Excel Application state for:

- `ScreenUpdating`
- `EnableEvents`
- `DisplayAlerts`
- `Calculation`
- `Cursor`

The class directly depends on the companion module’s shared manager procedures, so both files must be imported into the same VBA project.

### Required shared procedures

The companion module exposes these procedures/functions used by the class:

- `PM_TW_BeginSession`
- `PM_TW_EndSession`
- `PM_TW_ActiveCount`

Additional diagnostic/recovery helpers are also exposed there.

---

## Installation

1. Open the target workbook, add-in, or VBA project.
2. Open the VBA Editor with `ALT + F11`.
3. Import the class module:
   - `File` -> `Import File...`
   - select `cPerformanceManager.cls`
4. Import the companion standard module:
   - `File` -> `Import File...`
   - select `M_cPM_TimeWasters.bas`
5. Save the project as a macro-enabled file type such as:
   - `.xlsm`
   - `.xlam`
6. Compile the project:
   - `Debug` -> `Compile VBAProject`
7. Run a small smoke test.

---

## Quick start

### Basic timing with default QPC backend

```vb
Sub Example_BasicTiming()

    Dim cPM As cPerformanceManager
    Dim ElapsedS As Double

    Set cPM = New cPerformanceManager

    cPM.StartTimer

    Range("A1:A10000").Value = 1

    ElapsedS = cPM.ElapsedSeconds

    Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

### Human-readable elapsed time

```vb
Sub Example_ElapsedTimeText()

    Dim cPM As cPerformanceManager

    Set cPM = New cPerformanceManager

    cPM.StartTimer 5

    Application.Calculate

    Debug.Print cPM.ElapsedTime

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

---

## Public API

### Core timing

- `StartTimer(Optional ByVal iMethod As Integer = 5, Optional ByVal AlignToNextTick As Boolean = False)`
- `ElapsedSeconds(Optional ByVal iMethod As Integer = 0) As Double`
- `ElapsedTime(Optional ByVal iMethod As Integer = 0) As String`

### Session / state inspection

- `T1 As Double`
- `T2 As Double`
- `ET As Double`
- `StrictMode As Boolean`
- `ActiveMethodID As Integer`
- `HasActiveSession As Boolean`
- `MethodName(ByVal Idx As Integer) As String`

### Diagnostics / benchmark helpers

- `OverheadMeasurement_Seconds(Optional ByVal iMethod As Integer = 5, Optional ByVal Iterations As Long = 1000) As Double`
- `OverheadMeasurement_Text(Optional ByVal iMethod As Integer = 5) As String`
- `Get_SystemTickInterval As String`
- `QPC_Get_SystemTickInterval As String`
- `QPC_FrequencyPerSecond As String`
- `QPC_FrequencyPerSecond_Value As Double`

### Execution control / environment

- `Pause(ByVal dSeconds As Double, Optional ByVal iMethod As Integer = 1)`
- `ResetEnvironment()`

### Shared time-waster control

- `TW_Turn_OFF(Optional ByVal Except As TW_Enum = TW_Enum.None)`
- `TW_Turn_ON()`
- `TW_IsActive As Boolean`
- `TW_ActiveSessionCount As Long`

---

## Strict mode

The class defaults to:

```vb
cPM.StrictMode = True
```

In strict mode, invalid usage raises errors. Examples include:

- invalid timer method values
- calling `ElapsedSeconds` before `StartTimer`
- trying to read elapsed time with a method different from the active session method
- requesting `QPC` when unavailable

In non-strict mode, the class falls back where possible.

---

## Time-waster suppression

The class can suppress selected Excel-side overhead during benchmark runs by coordinating shared Application state through the companion module.

### Disable all supported time-wasters

```vb
Sub Example_TW_AllOff()

    Dim cPM As cPerformanceManager

    Set cPM = New cPerformanceManager

    cPM.TW_Turn_OFF
    cPM.StartTimer 5

    Range("A1:A50000").Formula = "=ROW()"

    Debug.Print cPM.ElapsedSeconds

    cPM.TW_Turn_ON
    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

### Disable with exemptions

```vb
Sub Example_TW_WithExceptions()

    Dim cPM As cPerformanceManager

    Set cPM = New cPerformanceManager

    cPM.TW_Turn_OFF TW_Enum.ScreenUpdating Or TW_Enum.EnableEvents
    cPM.StartTimer 5

    Application.CalculateFull

    Debug.Print cPM.ElapsedTime

    cPM.TW_Turn_ON
    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

---

## Safe cleanup pattern

```vb
Sub Example_SafePattern()

    Dim cPM As cPerformanceManager

    On Error GoTo CleanFail

    Set cPM = New cPerformanceManager

    cPM.TW_Turn_OFF
    cPM.StartTimer 5

    Worksheets(1).UsedRange.Calculate

    Debug.Print "Elapsed: " & cPM.ElapsedTime

CleanExit:
    If Not cPM Is Nothing Then
        cPM.ResetEnvironment
        Set cPM = Nothing
    End If
    Exit Sub

CleanFail:
    Debug.Print "Error " & Err.Number & " - " & Err.Description
    Resume CleanExit

End Sub
```

---

## Timing design notes

### Session-bound model

Timing is session-bound:

- `StartTimer` establishes the active timing backend
- `ElapsedSeconds` and `ElapsedTime` are validated against that same backend

This helps prevent accidental cross-method timing mistakes.

### QPC storage

QPC ticks and frequency are stored as `Currency` to preserve stable tick precision while remaining efficient in VBA.

### Rollover handling

The class includes rollover-aware logic for:

- `Timer`
- `GetTickCount` (32-bit path)
- `timeGetTime`
- `timeGetSystemTime`

### Multimedia timer resolution

Method `3` may enable `timeBeginPeriod(1)` to request 1ms timer resolution. This affects system timer resolution globally and can increase power usage. The class balances that through `ResetEnvironment`, with `Class_Terminate` as a fallback safety net.

---

## Diagnostics

The class exposes diagnostics to inspect the timing environment:

- nominal system tick interval
- QPC tick interval
- QPC frequency
- measurement overhead estimate

```vb
Sub Example_Diagnostics()

    Dim cPM As cPerformanceManager

    Set cPM = New cPerformanceManager

    Debug.Print cPM.Get_SystemTickInterval
    Debug.Print cPM.QPC_Get_SystemTickInterval
    Debug.Print cPM.QPC_FrequencyPerSecond
    Debug.Print cPM.QPC_FrequencyPerSecond_Value

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

---

## Benchmark guidance

For most benchmark scenarios:

- prefer `ElapsedSeconds()` for numeric measurement
- prefer method `5` (`QPC`) for the primary benchmark path
- use `ElapsedTime()` for presentation and logging
- use `AlignToNextTick := True` only when the extra polling cost is justified
- use TW suppression only when you explicitly want to reduce Excel-side noise

---

## Limitations

- Primarily designed for Windows/VBA
- API-backed methods are not intended for non-Windows hosts
- Shared TW control requires the companion module
- `Now()` is not a preferred monotonic benchmark source
- `Application.Wait` is coarse and not suitable for fine-grained timing
- “Nanoseconds” in formatted output are display precision, not guaranteed measurement resolution

---

## Author

Daniele Penza

---

## Version

1.0
