# 🧪 openstaadpy Test Suite

This folder contains automated tests for the **openstaadpy** library, which provides a Python interface to Bentley STAAD.Pro via the OpenSTAAD COM API.

---

## 📖 Overview

The tests validate the integration and functionality of openstaadpy, ensuring correct interaction with STAAD.Pro for geometry, loads, properties, and output operations.

---

## 💻 System Requirements

- 🪟 **Windows OS** (required for STAAD.Pro and COM automation)
- 🏗️ **Bentley STAAD.Pro 2026** installed at  
  `C:\Program Files\Bentley\Engineering\STAAD.Pro 2026\STAAD\Bentley.Staad.exe`
- 🐍 **Python 3.8+**
- 🧪 **pytest** (for running the tests)
- 🧩 **psutil** (Python package, used for process management)
- 📦 **openstaadpy** library (must be installed or available in the project)

---

## 🚀 Getting Started

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
   Or, if you only need the test dependencies:
   ```bash
   pip install pytest psutil
   ```

2. **Ensure STAAD.Pro is installed** and licensed on your machine.

3. **Verify the STAAD.Pro path** in `conftest.py` matches your installation.  
   Update the path if STAAD.Pro is installed elsewhere.

---

## 🏃 Running the Tests

From the `Tests/` directory, run:
```bash
pytest --junitxml=TestResults/allTestResults.xml
```
Or, to run a specific test file:
```bash
pytest test_geometry.py --junitxml=TestResults/geometryTestResults.xml
```
Or, to run a specific test:
```bash
pytest test_geometry.py::TestGeometry::test_SelectMultipleNodes --junitxml=TestResults/test_SelectMultipleNodes.xml
```

### Alternate Entry Points

- **`run_openstaadpy_tests.bat`**: Windows batch script that installs dependencies and runs the full suite. Accepts an optional output filename argument:
  ```cmd
  run_openstaadpy_tests.bat MyResults.xml
  ```
- **`run_tests.py`**: Python script that launches STAAD.Pro, runs pytest, and quits cleanly. Supports all pytest arguments:
  ```bash
  python run_tests.py test_geometry.py -v
  ```

---

## 🗂️ Test Structure

### Configuration & Helpers
- **conftest.py**: Session-wide setup/teardown — launches STAAD.Pro, initializes OpenSTAAD COM objects.
- **test_fixtures_base.py**: Base test class with common setup/teardown for model lifecycle.
- **openstaad_test_helpers.py**: Shared helper functions for test setup, assertions, and model management.
- **openstaad_error_codes.py**: Centralized error code constants for OpenSTAAD operations.
- **requirements.txt**: Python dependencies for the test suite.
- **pytest.ini**: Custom pytest markers for tracking known bugs and skipped tests.

### Test Files
- **test_osroot.py**: Root API tests — file operations, analysis, units, job info.
- **test_geometry.py**: Node, beam, plate, solid, and group operations.
- **test_load.py**: Load case creation, load application, and combinations.
- **test_output.py**: Post-analysis result queries — displacements, forces, reactions.
- **test_property.py**: Material assignment, section properties, and member offsets.
- **test_support.py**: Boundary conditions — fixed, pinned, spring, inclined, and foundation supports.
- **test_command.py**: Analysis and design command execution.
- **test_design.py**: Design brief and parameter management.
- **test_tables.py**: Report and table generation.

### Entry Points
- **run_openstaadpy_tests.bat**: Batch script — installs deps and runs full suite.
- **run_tests.py**: Python script — launches STAAD.Pro, runs pytest, quits cleanly.

---

## 🏅 Best Practices

- 🔹 **Isolate test cases**: Each test should focus on a single feature or method.
- 🔹 **Use fixtures**: Shared setup/teardown logic is managed via `pytest` fixtures in `conftest.py`.
- 🔹 **Mock where possible**: For pure Python logic, use mocking to avoid launching STAAD.Pro unnecessarily.
- 🔹 **Clean up resources**: The test suite ensures STAAD.Pro is closed after tests.
- 🔹 **Document tests**: Use docstrings and comments to clarify test intent.
- 🔹 **Organize tests**: Group related tests into separate files for maintainability.

---

## 📝 Notes

- ⚠️ The tests will attempt to launch STAAD.Pro automatically if it is not already running.
- ⏳ Some tests may take longer to execute due to STAAD.Pro startup and COM initialization.
- 💾 Make sure no unsaved work is open in STAAD.Pro before running tests, as the test suite may close the application.

---

**For more information, see the main project [README.md](../README.md) or the documentation in the `docs/` folder.**

