# openstaadpy

**openstaadpy** is a Python library that provides seamless access to the power of STAAD.Pro through the OpenSTAAD API. It exposes STAAD.Pro’s internal functions and routines, allowing engineers and developers to automate, analyze, and interact with STAAD models directly with Python.

---

## Overview

OpenSTAAD provides APIs for STAAD.Pro that enables external programs and scripts to access the data and functionality of an open instance of STAAD.Pro. While OpenSTAAD itself is a natively COM-based service and designed to be used within a macro based module of applications like STAAD.Pro, Microsoft Excel or AutoCAD, **openstaadpy** bridges the gap for Python users by offering a standard Python interface to access the same powerful automation features.

> **Note:**
> 1. OpenSTAAD requires a suitable version of STAAD.Pro to be installed. **openstaadpy** acts as a wrapper and cannot function without a valid STAAD.Pro installation.
> 2. As these functions access a current instance of STAAD.Pro, the STAAD.Pro application must be running for the OpenSTAAD functions to operate.
> 3. STAAD.Pro can be run with a standard STAAD.Pro license or with a STAAD.Pro Advanced license.

---

## Features

- Python wrappers for commonly used OpenSTAAD functions
- Easy access to geometry, loads, analysis results, and design data
- Simplified connection to STAAD.Pro instances
- Robust error handling and logging for automation workflows
- Suitable for both beginners and advanced users of STAAD automation

---

## Requirements

- Windows OS 11 or newer
- STAAD.Pro 2025 or newer installed and running
- Python 3.11 or higher

---

## Installation

Install directly from GitHub:
```bash
pip install git+https://github.com/BentleySystems/openstaadpy.git
```

---

## Usage
Use the following steps to start a Python project that can access the OpenSTAADpy functions

- Create a new Python project folder
- Install the openstaadpy package if not installed (see above).
- Launch STAAD.Pro and either create a new analytical model or open an existing file.
- Open your preferred Python IDE, open the Python project folder and insert the following code

```python
from openstaadpy import os_analytical
```
- Add the following code to create a connection to the STAAD.Pro instance.

```python
staad = os_analytical.connect()
```
If the system has multiple instances of STAAD.Pro running, then the specific file name must be referenced in the connection call, e.g. if the model to be accessed and open in an instance of STAAD.Pro is “C:\STAAD Model\Example.std”, then use

```python
staad = os_analytical.connect(filepath: "C:\\STAAD Model\\Example.std")
```

---

## Documentation

The  documentation can be built locally by navigating to your local openstaadpy repository and opening a terminal and entering the following:
```bash
pip install .[docs]
cd docs
.\make.bat html
```
Open `docs/build/html/index.html` in your browser.

---
