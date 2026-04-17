# cPerformanceManager

High-precision timing and performance-support utility for VBA on Windows.

`cPerformanceManager` provides a single, session-bound interface for multiple timing backends, numeric elapsed-time measurement, human-readable elapsed-time diagnostics, benchmark overhead measurement, pause helpers, and shared Excel “time-waster” suppression. Those suppression features can be used both for cleaner benchmark runs and as a general-purpose Excel/VBA performance aid in procedures that do not measure elapsed time.

---

## Overview

VBA’s built-in timing options are often not ideal for instrumentation and benchmarking:

- `Timer` has limited resolution and rolls over at midnight
- `Now()` is wall-clock-based and is not a preferred monotonic benchmark source
- Windows exposes better monotonic and high-resolution counters, such as `QueryPerformanceCounter`

`cPerformanceManager` wraps those timing sources behind a consistent API and adds performance-oriented helpers for Excel/VBA projects.

The shared Excel “time-waster” suppression features are not limited to timed benchmarks. They can also be used in ordinary workbook automation to reduce overhead from screen refresh, events, alerts, calculation-mode churn, and cursor updates during heavy procedures.

---

## Repository Contents

```text
/README.md
/src/cPerformanceManager.cls
/src/M_cPM_TimeWasters.bas
/examples/M_cPM_USAGE_EXAMPLES.bas
/test/M_cPM_RegressionTests.bas
```

### Required files

These two files are required for normal use:

- `src/cPerformanceManager.cls`
- `src/M_cPM_TimeWasters.bas`

### Optional companion files

These files are optional but useful:

- `examples/M_cPM_USAGE_EXAMPLES.bas`
  - compact usage examples and recommended integration patterns
- `test/M_cPM_RegressionTests.bas`
  - regression test harness

---

## Features

- Multiple timing methods under one interface
- Session-bound timing model
- Low-overhead numeric elapsed-time measurement
- Human-readable elapsed-time formatting (`ElapsedTime`)
- Formatting of an already measured elapsed-seconds value without taking a second timing sample
- Benchmark overhead measurement helpers
- Timing/source diagnostics
- Pause/wait helper (`Pause`)
- Shared Excel “time-waster” suppression (`TW_Turn_OFF` / `TW_Turn_ON`) for both benchmarking and general Excel/VBA performance improvement
- Strict-mode validation for safer usage

---

## Timer Methods

| ID | Method | Notes |
|---:|---|---|
| 1 | `Timer` | Seconds since midnight; rolls over every 24 hours |
| 2 | `GetTickCount / GetTickCount64` | Milliseconds since boot; 32-bit path wraps at ~49.7 days |
| 3 | `timeGetTime` | Millisecond counter with 32-bit rollover semantics |
| 4 | `timeGetSystemTime (MMTIME / TIME_MS)` | Millisecond source treated with 32-bit rollover semantics |
| 5 | `QueryPerformanceCounter (QPC)` | Default and recommended high-resolution benchmark source |
| 6 | `Now() * 86400` | Wall-clock seconds; mainly useful for diagnostics |

---

## Requirements

- Microsoft Excel with VBA enabled
- Windows host environment for API-backed timing methods (`2..5`)
- `VBA7` / `Win64` conditional-compilation support as required by the host
- Required source files:
  - `cPerformanceManager.cls`
  - `M_cPM_TimeWasters.bas`

---

## Installation

1. Open the target workbook/add-in/VBA project.
2. Open the VBA Editor (`ALT + F11`).
3. Import:
   - `src/cPerformanceManager.cls`
   - `src/M_cPM_TimeWasters.bas`
4. Save as macro-enabled (`.xlsm` / `.xlam`).
5. Compile (`Debug` → `Compile VBAProject`).
6. Run a smoke test.

Optional: also import `examples/M_cPM_USAGE_EXAMPLES.bas` and `test/M_cPM_RegressionTests.bas`.

---

## Quick Start

### 1) Basic timing (default QPC backend)

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

### 2) Format an already measured elapsed value

```vb
Sub Example_FormatExistingElapsed()

    Dim cPM As cPerformanceManager
    Dim ElapsedS As Double
    Dim ElapsedT As String

    Set cPM = New cPerformanceManager

    cPM.StartTimer 5, False
    Application.Calculate

    ElapsedS = cPM.ElapsedSeconds
    ElapsedT = cPM.ElapsedTime(, ElapsedS)

    Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")
    Debug.Print "Elapsed time   : " & ElapsedT

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

### 3) Improve performance without measuring elapsed time

```vb
Sub Example_PerformanceOnly()

    Dim cPM As cPerformanceManager

    Set cPM = New cPerformanceManager

    cPM.TW_Turn_OFF

    Range("A1:A50000").Formula = "=ROW()"
    Application.Calculate

    cPM.TW_Turn_ON
    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

### 4) Benchmark with Excel TW suppression

```vb
Sub Example_BenchmarkWithSuppression()

    Dim cPM As cPerformanceManager
    Dim ElapsedS As Double

    Set cPM = New cPerformanceManager

    cPM.TW_Turn_OFF
    cPM.StartTimer 5

    Range("A1:A50000").Formula = "=ROW()"

    ElapsedS = cPM.ElapsedSeconds
    Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")

    cPM.TW_Turn_ON
    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

---

## Core Public API (class)

### Timing

- `StartTimer(Optional iMethod As Integer = 5, Optional AlignToNextTick As Boolean = False)`
- `ElapsedSeconds(Optional iMethod As Integer = 0) As Double`
- `ElapsedTime(Optional iMethod As Integer = 0, Optional ElapsedSecondsIn As Variant) As String`

### Session/state inspection

- `T1 As Double`
- `T2 As Double`
- `ET As Double`
- `ActiveMethodID As Integer`
- `HasActiveSession As Boolean`
- `MethodName(ByVal Idx As Integer) As String`
- `StrictMode As Boolean` (Get/Let)

### Diagnostics / benchmarking

- `OverheadMeasurement_Seconds(Optional iMethod As Integer = 5, Optional Iterations As Long = 1000) As Double`
- `OverheadMeasurement_Text(Optional iMethod As Integer = 5, Optional Iterations As Long = 1000) As String`
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

## TW_Enum flags

Use with `TW_Turn_OFF Except:=...` as bitmask flags:

- `TW_Enum.None`
- `TW_Enum.ScreenUpdating`
- `TW_Enum.EnableEvents`
- `TW_Enum.DisplayAlerts`
- `TW_Enum.Calculation`
- `TW_Enum.Cursor`

Example:

```vb
cPM.TW_Turn_OFF TW_Enum.ScreenUpdating Or TW_Enum.EnableEvents
```

---

## Strict Mode Behavior

- Default: `StrictMode = True`
- Strict mode raises on:
  - invalid timer method
  - `ElapsedSeconds` before `StartTimer`
  - explicit elapsed method mismatch vs active session
- Non-strict mode attempts fallback/coercion where supported

---

## Notes

- Prefer method `5` (QPC) for benchmark-grade timing.
- Use `ElapsedSeconds` for numeric logic; use `ElapsedTime` for display.
- When you already have a numeric elapsed value, prefer `ElapsedTime(, ElapsedSecondsIn)` to avoid a second timing sample.
- Always call `ResetEnvironment` in normal flows (and in error cleanup paths).
- TW suppression is useful both for cleaner benchmark runs and for improving performance in ordinary Excel/VBA procedures, even when no elapsed-time measurement is being taken.

---

## Running Examples / Tests

- Import `examples/M_cPM_USAGE_EXAMPLES.bas` for compact usage examples.
- Import `test/M_cPM_RegressionTests.bas` and run:
  - `Run_cPerformanceManager_RegressionSuite`

---

## License

This project is licensed under the terms in `LICENSE`.
