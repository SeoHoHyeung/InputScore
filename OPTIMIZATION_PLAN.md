# Optimization and Refactoring Plan

This document outlines a high-level plan to refactor and optimize the **InputScore** project for better maintainability and faster execution.

## 1. Profiling and Benchmarking
- Establish baseline performance using Python's `cProfile` or other profiling tools.
- Measure startup time, file loading time and UI responsiveness.
- Identify the slowest functions or modules.

## 2. Modularization
- Reorganize code into clear modules:
  - `core` for business logic (`app_logic.py` and helpers).
  - `ui` for UI components (`main_window.py`, widgets, resources).
  - `services` for TTS and external integrations.
- Adopt consistent naming conventions and docstrings.

## 3. Dependency Management
- Create a `requirements.txt` or `pyproject.toml` to manage dependencies.
- Pin library versions for reproducible builds.
- Remove unused imports and packages.

## 4. Excel Handling
- Continue using `openpyxl` in read-only mode but consider `pandas` for vectorized operations if datasets become large.
- Cache heavy operations (e.g., header parsing) to reduce repeated work.
- When saving files, batch cell writes as already done but verify that formulas or formatting are preserved correctly.

## 5. UI Performance
- Delay expensive UI updates using timers (`QTimer`) to avoid blocking the main thread (already partially implemented).
- Ensure signals are connected/disconnected properly to prevent redundant work.
- Consider lazy loading for large tables: only populate visible rows and load more as the user scrolls.

## 6. TTS Manager
- Keep the worker thread approach but profile queue handling.
- Reuse a single `win32com.client.Dispatch` instance rather than creating multiple instances.
- Handle thread shutdown gracefully on exit to avoid leaks.

## 7. Memory Optimization
- Release references to large objects (such as loaded workbooks) as soon as they are no longer needed.
- Use `gc.collect()` selectively rather than globally.
- Monitor memory with tools like `tracemalloc` in debug builds.

## 8. Testing and Continuous Integration
- Add unit tests for core logic (e.g., data loading, score updates, saving).
- Integrate with GitHub Actions or similar CI to run tests and linting on each commit.

## 9. Packaging
- Ensure the `PyInstaller` spec is up to date and includes all resources.
- Provide clear build instructions in a `README`.

## 10. Future Improvements
- Evaluate switching to a faster Excel library (e.g., `xlsxwriter` for writing) if profiling shows bottlenecks.
- Consider asynchronous I/O or multiprocessing for heavy tasks, keeping thread-safety in mind with Qt.

---
This plan can serve as a starting point for refactoring the project with speed and maintainability as top priorities.
