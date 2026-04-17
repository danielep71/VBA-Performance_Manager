# cPerformanceManager

High-precision timing and benchmark-support utility for VBA on Windows.

`cPerformanceManager` provides a single, session-bound interface for multiple timing backends, numeric elapsed-time measurement, human-readable elapsed-time diagnostics, benchmark overhead measurement, pause helpers, and shared Excel “time-waster” suppression. That suppression can be used both to reduce benchmark noise and, more generally, to improve performance in ordinary Excel/VBA procedures even when no elapsed-time measurement is being taken.

---

## Overview

VBA’s built-in timing options are often not ideal for instrumentation and benchmarking:

- `Timer` has limited resolution and rolls over at midnight
- `Now()` is wall-clock based and is not a preferred monotonic benchmark source
- Windows exposes better monotonic and high-resolution counters such as `QueryPerformanceCounter`

`cPerformanceManager` wraps those timing sources behind a consistent interface and adds benchmark-oriented helpers for Excel/VBA projects.

---

## Highlights

- Six timing backends under one interface
- Session-bound timing model with strict/non-strict validation
- Low-overhead numeric elapsed-time measurement via `ElapsedSeconds`
- Human-readable elapsed-time formatting via `ElapsedTime`
- Ability to format an already measured elapsed value without taking a second timing sample
- Overhead-measurement helpers for benchmark interpretation
- System and QPC diagnostics
- Pause helper with multiple wait strategies
- Shared Excel “time-waster” suppression (`TW_Turn_OFF` / `TW_Turn_ON`) for both benchmarking and general Excel/VBA performance improvement. The shared Excel “time-waster” suppression features are not limited to timed benchmarks. They can also be used in ordinary workbook automation to reduce overhead from screen refresh, events, alerts, calculation-mode churn, and cursor updates during heavy procedures.
- Optional demo workbook builder and regression-test suite

---

## Timer Methods

| ID | Method | Notes |
|---:|---|---|
| 1 | `Timer` | Seconds since midnight; rolls over every 24 hours |
| 2 | `GetTickCount / GetTickCount64` | Milliseconds since boot; the 32-bit path wraps at about 49.7 days |
| 3 | `timeGetTime` | Millisecond counter with 32-bit rollover semantics |
| 4 | `timeGetSystemTime (MMTIME / TIME_MS)` | Millisecond source treated with 32-bit rollover semantics |
| 5 | `QueryPerformanceCounter (QPC)` | Default and recommended high-resolution benchmark source |
| 6 | `Now() * 86400` | Wall-clock seconds; mainly useful for diagnostics |

---

## Repository Contents

A practical repository layout is:

```text
/README.md
/src/cPerformanceManager.cls
/src/M_cPM_TimeWasters.bas
/src/M_cPM_USAGE_EXAMPLES.bas          (optional)
/src/M_cPM_DEMO.bas                    (optional)
/test/M_cPM_RegressionTests.bas        (optional)
```

### Required files

These two files are required for normal use:

- `src/cPerformanceManager.cls`
- `src/M_cPM_TimeWasters.bas`

### Optional companion files

These files are optional but useful:

- `src/M_cPM_USAGE_EXAMPLES.bas`
  - compact usage examples and recommended patterns
- `src/M_cPM_DEMO.bas`
  - builds demo sheets and demo actions inside Excel
- `test/M_cPM_RegressionTests.bas`
  - regression-test harness with worksheet-based logging

---

## Requirements

- Microsoft Excel with VBA enabled
- Windows host environment for API-backed timing methods (`2..5`)
- Appropriate `VBA7` / `Win64` conditional compilation support as required by the host

---

## Installation

1. Open the target workbook, add-in, or VBA project.
2. Open the VBA Editor with `ALT + F11`.
3. Import the required source files:
   - `cPerformanceManager.cls`
   - `M_cPM_TimeWasters.bas`
4. Save the workbook as macro-enabled (`.xlsm` or `.xlam`).
5. Compile the VBA project from `Debug` → `Compile VBAProject`.
6. Run a small smoke test.

Optional: also import the usage examples, demo module, and regression tests.

---

## Quick Start

### 1) Basic timing with the default backend

```vb
Sub Example_BasicTiming()

    Dim cPM As cPerformanceManager
    Dim ElapsedS As Double

    Set cPM = New cPerformanceManager

    cPM.StartTimer

    ActiveSheet.Range("A1:A10000").Value = 1

    ElapsedS = cPM.ElapsedSeconds

    Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

### 2) Format an already measured elapsed value

```vb
Sub Example_ElapsedTime_FromMeasuredSeconds()

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

### 3) Benchmark with shared Excel TW suppression

```vb
Sub Example_BenchmarkWithSuppression()

    Dim cPM As cPerformanceManager
    Dim ElapsedS As Double

    Set cPM = New cPerformanceManager

    cPM.TW_Turn_OFF
    cPM.StartTimer 5, False

    ActiveSheet.Range("A1:A50000").Formula = "=ROW()"

    ElapsedS = cPM.ElapsedSeconds

    Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

---

## Core Public API

### Timing

- `StartTimer(Optional ByVal iMethod As Integer = 5, Optional ByVal AlignToNextTick As Boolean = False)`
- `ElapsedSeconds(Optional ByVal iMethod As Integer = 0) As Double`
- `ElapsedTime(Optional ByVal iMethod As Integer = 0, Optional ByVal ElapsedSecondsIn As Variant) As String`

### Session / state inspection

- `T1 As Double`
- `T2 As Double`
- `ET As Double`
- `ActiveMethodID As Integer`
- `HasActiveSession As Boolean`
- `MethodName(ByVal Idx As Integer) As String`
- `StrictMode As Boolean` (`Get` / `Let`)

### Diagnostics / benchmarking

- `OverheadMeasurement_Seconds(Optional ByVal iMethod As Integer = 5, Optional ByVal Iterations As Long = 1000) As Double`
- `OverheadMeasurement_Text(Optional ByVal iMethod As Integer = 5, Optional ByVal Iterations As Long = 1000) As String`
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

## TW_Enum Flags

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

## Strict Mode

- Default: `StrictMode = True`
- Strict mode raises on:
  - invalid timer method
  - `ElapsedSeconds` before `StartTimer`
  - explicit elapsed-method mismatch versus the active session
- Non-strict mode attempts fallback / coercion where supported

---

## Demo Workbook

If you import the optional demo module, you can build a ready-to-use in-workbook surface for the class.

Typical demo sheets include:

- `DEMO_cPM`
  - control panel, run buttons, results log
- `DATA_cPM`
  - value-fill / formula-fill workload anchors
- `HELP_cPM`
  - embedded usage guidance

This is useful when you want a workbook-native demo instead of Immediate Window examples.

---

## Regression Tests

If you import the optional regression-test module, you can run a structured regression suite covering:

- constructor/default state
- session-bound timing behavior
- strict / non-strict validation
- elapsed-time formatting
- pause methods
- TW shared-state lifecycle
- cleanup idempotence

The modern regression harness is best used with worksheet-based logging rather than relying only on `Debug.Print`.

---

## Recommended Usage Notes

- Prefer method `5` (QPC) for benchmark-grade timing.
- Use `ElapsedSeconds` for numeric logic; use `ElapsedTime` for display.
- Always call `ResetEnvironment` in normal flows (and in error cleanup paths).
- TW suppression is useful both for cleaner benchmark runs and for improving performance in ordinary Excel/VBA procedures even when no elapsed-time measurement is being taken.

---

## License

This project is licensed under the terms in `LICENSE`.
