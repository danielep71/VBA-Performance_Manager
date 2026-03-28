# cPerformanceManager

High-precision timing and benchmark-support utility for VBA on Windows.

`cPerformanceManager` provides a single, session-bound interface for multiple timing backends, human-readable elapsed-time diagnostics, benchmark overhead measurement, pause helpers, and optional shared Excel “time-waster” suppression for cleaner benchmarking.

---

## Why this exists

VBA's built-in timing options are often not ideal for benchmarking:

- `Timer` has limited resolution and rolls over at midnight.
- `Now()` is wall-clock based and is not ideal as a primary benchmark source.
- Windows exposes better monotonic and high-resolution counters such as `QueryPerformanceCounter`.

This class wraps those sources behind a consistent API and adds benchmark-oriented utilities for Excel/VBA work.

---

## Features

- Multiple timing methods under one interface
- Session-bound timing model
- Low-overhead numeric elapsed-time measurement
- Human-readable elapsed-time formatting
- Benchmark overhead measurement helpers
- Timer/source diagnostics
- Pause/wait helpers
- Optional shared Excel “time-waster” suppression for benchmark runs

---

## Timer methods

The class supports the following timing backends:

| ID | Method | Notes |
|---:|---|---|
| 1 | `Timer` | Seconds since midnight; rolls over every 24 hours |
| 2 | `GetTickCount / GetTickCount64` | Milliseconds since boot; the 32-bit version wraps at about 49.7 days |
| 3 | `timeGetTime` | Millisecond counter with 32-bit rollover semantics; can use 1ms timer resolution |
| 4 | `timeGetSystemTime (MMTIME / TIME_MS)` | Millisecond source treated with 32-bit rollover semantics |
| 5 | `QueryPerformanceCounter (QPC)` | Default and recommended high-resolution benchmark source |
| 6 | `Now() * 86400` | Wall-clock seconds; useful mainly for diagnostics |

---

## Requirements

- Microsoft Excel / VBA
- Windows for API-backed methods (`2..5`)
- Appropriate `VBA7` / `Win64` conditional compilation support
- The class file: `cPerformanceManager.cls`

### Compatibility note

On non-Windows hosts, only `Timer()` / `Now()`-based methods are conceptually portable, if retained.

---

## External dependency for TW control

The class includes shared “time-waster” suppression, but that feature depends on a companion shared manager module exposing these procedures/functions:

- `PM_TW_BeginSession`
- `PM_TW_EndSession`
- `PM_TW_ActiveCount`

If you do **not** include that companion module, the core timing features still work, but the TW-related methods will not.

---

## Installation

1. Export or copy the class file as `cPerformanceManager.cls`.
2. Import it into your VBA project:
   - Open the VBA Editor.
   - Choose **File -> Import File...**
3. If you want TW suppression support, also import the companion TW manager module.
4. Save, compile, and test.

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

## Strict mode behavior

The class defaults to:

```vb
PM.StrictMode = True
```

In strict mode, invalid usage raises errors. Examples include:

- invalid timer method values
- calling `ElapsedSeconds` before `StartTimer`
- trying to read elapsed time with a method different from the active session method
- requesting QPC when unavailable

In non-strict mode, the class falls back where possible.

---

## Basic usage

### 1. Default timing with QPC

```vb
Sub Example_BasicTiming()

    Dim PM          As cPerformanceManager
    Dim ElapsedS    As Double

    Set PM = New cPerformanceManager

    PM.StartTimer

    Range("A1:A10000").Value = 1

    ElapsedS = PM.ElapsedSeconds

    Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")

    PM.ResetEnvironment
    Set PM = Nothing

End Sub
```

### 2. Human-readable elapsed time

```vb
Sub Example_ElapsedTimeText()

    Dim PM As cPerformanceManager

    Set PM = New cPerformanceManager

    PM.StartTimer 5

    Application.Calculate

    Debug.Print PM.ElapsedTime

    PM.ResetEnvironment
    Set PM = Nothing

End Sub
```

### 3. Specific timing backend

```vb
Sub Example_SpecificMethod()

    Dim PM As cPerformanceManager

    Set PM = New cPerformanceManager

    PM.StartTimer 2

    Worksheets(1).Range("A1").Formula = "=RAND()"

    Debug.Print "Method used: " & PM.MethodName(PM.ActiveMethodID)
    Debug.Print "Elapsed seconds: " & Format$(PM.ElapsedSeconds, "0.000000000")

    PM.ResetEnvironment
    Set PM = Nothing

End Sub
```

### 4. Aligned start

```vb
Sub Example_AlignedStart()

    Dim PM As cPerformanceManager

    Set PM = New cPerformanceManager

    PM.StartTimer 5, True

    DoEvents

    Debug.Print "Aligned elapsed seconds: " & _
                Format$(PM.ElapsedSeconds, "0.000000000")

    PM.ResetEnvironment
    Set PM = Nothing

End Sub
```

### 5. TW suppression during a benchmark

```vb
Sub Example_TimeWasters()

    Dim PM As cPerformanceManager

    Set PM = New cPerformanceManager

    PM.TW_Turn_OFF
    PM.StartTimer 5

    Range("A1:A50000").Formula = "=ROW()"

    Debug.Print PM.ElapsedSeconds

    PM.TW_Turn_ON
    PM.ResetEnvironment
    Set PM = Nothing

End Sub
```

### 6. Safe cleanup pattern

```vb
Sub Example_SafePattern()

    Dim PM As cPerformanceManager

    On Error GoTo CleanFail

    Set PM = New cPerformanceManager

    PM.TW_Turn_OFF
    PM.StartTimer 5

    Worksheets(1).UsedRange.Calculate

    Debug.Print "Elapsed: " & PM.ElapsedTime

CleanExit:
    If Not PM Is Nothing Then
        PM.ResetEnvironment
        Set PM = Nothing
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

- `StartTimer` establishes the active timing backend.
- `ElapsedSeconds` and `ElapsedTime` are validated against that same backend.

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

    Dim PM As cPerformanceManager

    Set PM = New cPerformanceManager

    Debug.Print PM.Get_SystemTickInterval
    Debug.Print PM.QPC_Get_SystemTickInterval
    Debug.Print PM.QPC_FrequencyPerSecond
    Debug.Print PM.QPC_FrequencyPerSecond_Value

    PM.ResetEnvironment
    Set PM = Nothing

End Sub
```

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
- TW support requires a companion shared manager module
- `Now()` is not a preferred monotonic benchmark source
- `Application.Wait` is coarse and not suitable for fine-grained timing
- “Nanoseconds” in formatted output are display precision, not guaranteed measurement resolution

---

## Author

Daniele Penza

## Version

2.0
