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

## Repository Contents

```text
/README.md
/src/cPerformanceManager.cls
/src/M_cPM_TimeWasters.bas
/examples/M_cPM_Examples.bas
/test/M_cPM_Test.bas
```

### Required files

These two files are required for normal use:

- `src/cPerformanceManager.cls`
- `src/M_cPM_TimeWasters.bas`

### Optional companion files

These files are optional but strongly recommended:

- `examples/M_cPM_Examples.bas`
  - example and launcher routines
  - practical usage patterns
  - small smoke-test style demonstrations

- `test/M_cPM_Test.bas`
  - regression-test harness
  - public validation entry point for repository testing

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

## Timer Methods

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
- The following required source files:
  - `cPerformanceManager.cls`
  - `M_cPM_TimeWasters.bas`

### Compatibility Note

On non-Windows hosts, only `Timer()` / `Now()`-based methods are conceptually portable, if retained.

---

## Required Companion Module

This project includes a required companion standard module:

- `M_cPM_TimeWasters.bas`

That module manages shared Excel `Application` state for:

- `ScreenUpdating`
- `EnableEvents`
- `DisplayAlerts`
- `Calculation`
- `Cursor`

The class directly depends on the companion module’s shared manager procedures, so both required files must be imported into the same VBA project.

### Required shared procedures

The companion module exposes these procedures/functions used by the class:

- `PM_TW_BeginSession`
- `PM_TW_EndSession`
- `PM_TW_ActiveCount`

Additional diagnostic and recovery helpers are also exposed there.

---

## Installation

### Core installation

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

### Optional installation

If you want repository examples and the test harness as well, also import:

- `M_cPM_Examples.bas`
- `M_cPM_Test.bas`

---

## Quick Start

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

    cPM.PauseSeconds 1.25

    Debug.Print cPM.ElapsedTimeText

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

### Benchmark run with time-waster suppression

```vb
Sub Example_BenchmarkWithSuppression()

    Dim cPM As cPerformanceManager
    Dim ElapsedS As Double

    Set cPM = New cPerformanceManager

    cPM.StartTimer 5
    cPM.SuspendTimeWasters

    Range("A1:A50000").Formula = "=ROW()"

    ElapsedS = cPM.ElapsedSeconds

    Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

---

## Public API Summary

Representative public members include:

- timer lifecycle:
  - `StartTimer`
  - `RestartTimer`
  - `ResetTimer`

- elapsed-time access:
  - `ElapsedSeconds`
  - `ElapsedMilliseconds`
  - `ElapsedMicroseconds`
  - `ElapsedTicks`
  - `ElapsedTimeText`

- diagnostics / metadata:
  - `MethodID`
  - `MethodName`
  - `Frequency`
  - `IsRunning`

- benchmark helpers:
  - overhead measurement helpers
  - pause / wait helpers

- environment control:
  - time-waster suppression helpers
  - environment reset helpers

Refer to the class source and wiki for the full contract and procedure-level behavior.

---

## Strict Mode

The class supports stricter validation behavior to reduce misuse in benchmark scenarios.

Typical uses include:

- enforcing correct timer lifecycle
- detecting invalid method requests earlier
- making caller mistakes more obvious during development

Use strict mode when correctness is more important than permissive convenience.

---

## Examples Module

`examples/M_cPM_Examples.bas` provides runnable examples and utility launchers.

Use it when you want:

- a practical starting point
- smoke-test style demonstrations
- example calling patterns for common timing scenarios

This module is not required for production use, but it is useful for onboarding and repository exploration.

---

## Regression Tests

`test/M_cPM_Test.bas` provides the regression-test harness for the repository.

Use it when you want to validate the behavior of the class after refactoring or packaging changes.

Recommended workflow:

1. import the required core files
2. import `M_cPM_Test.bas`
3. run the public regression entry point exposed by the test module
4. review the Immediate Window output and any assertion failures

---

## Design Notes

### Session-bound model

The class is designed around a session-bound timing model rather than a purely stateless collection of helper functions. This supports:

- consistent elapsed-time queries
- benchmark-oriented lifecycle handling
- environment suppression sessions

### Shared time-waster suppression

Excel benchmark runs often benefit from temporarily suppressing costly `Application` behaviors. The companion module centralizes that shared state so nested or repeated benchmark use remains safer and easier to unwind.

### Windows-backed timing

The most accurate and useful benchmark methods depend on Windows timing APIs. In practice, `QueryPerformanceCounter` is the preferred default for serious timing work.

---

## Limitations

- Primarily intended for Windows-hosted Excel/VBA environments
- API-backed methods depend on the host and its available Windows timing services
- Some timing methods have rollover semantics by design
- `Now()` is provided mainly for diagnostics, not as the preferred benchmark source
- Shared Excel environment suppression should always be reset cleanly after use

---

## Suggested Use Cases

- benchmarking worksheet writes / reads
- comparing VBA implementations
- timing UDF or macro pipelines
- measuring Excel automation overhead
- controlled benchmark runs with reduced UI / event noise
- diagnostic timing during development

---

## Documentation

In addition to this README, the repository includes a GitHub Wiki with deeper documentation on:

- installation
- quick start
- timer methods
- core API
- diagnostics
- execution control
- strict mode
- testing
- architecture
- limitations
- version history

---

## License

MIT License.

See `LICENSE` for details.

---

## Status

This repository is intended as a reusable VBA utility component for high-precision timing and benchmark-support scenarios in Excel on Windows.
