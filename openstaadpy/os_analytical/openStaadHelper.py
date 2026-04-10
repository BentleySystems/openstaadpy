# ---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
# ---------------------------------------------------------------------------------------------
from comtypes import automation
from comtypes.automation import VARIANT
import ctypes

__all__ = [
    "make_safe_array_double",
    "make_safe_array_int",
    "make_safe_array_long",
    "make_safe_str",
    "make_safe_array_string",
    "make_variant_vt_ref",
    "make_variant_vt",
    "create_variant_int",
    "create_variant_float",
    "create_variant_bool",
    "create_variant_string",
    "make_variant_byref",
    "make_byref",
    "create_c_int",
    "create_c_float",
    "create_c_double",
    "create_c_bool",
    "make_safe_array_long_input",
    "make_safe_array_double_input",
    "make_safe_array_string_input",
    "create_bstr",
]


def make_safe_array_double(size):
    return automation._midlSAFEARRAY(ctypes.c_double).create([0] * size)


def make_safe_array_int(size):
    return automation._midlSAFEARRAY(ctypes.c_int).create([0] * size)


def make_safe_array_long(size):
    return automation._midlSAFEARRAY(ctypes.c_long).create([0] * size)


def make_safe_str():
    return automation.c_char_p()


def make_safe_array_string(size):
    return automation._midlSAFEARRAY(automation.BSTR).create([""] * size)


def make_variant_vt_ref(obj, var_type):
    var = automation.VARIANT()
    var._.c_void_p = ctypes.addressof(obj)
    var.vt = var_type | automation.VT_BYREF
    return var


def make_variant_vt(obj=None):
    var = automation.VARIANT()
    if obj:
        var.value = obj
    return var


def create_variant_int(value: int) -> VARIANT:
    """Create a VARIANT for an integer."""
    return VARIANT(value, automation.VT_BYREF | automation.VT_I4)


def create_variant_float(value: float) -> VARIANT:
    """Create a VARIANT for a float."""
    return VARIANT(value, automation.VT_BYREF | automation.VT_R8)


def create_variant_bool(value: bool) -> VARIANT:
    """Create a VARIANT for a boolean."""
    return VARIANT(value, automation.VT_BYREF | automation.VT_BOOL)


def create_variant_string(value: str) -> VARIANT:
    """Create a VARIANT for a string."""
    return VARIANT(value, automation.VT_BYREF | automation.VT_BSTR)


def make_variant_byref(vt: VARIANT) -> VARIANT:
    """Take an existing VARIANT and return it as BYREF."""
    vt_byref = VARIANT(vt.value, automation.VT_BYREF | vt.vt)
    return vt_byref


def make_byref(obj):
    return ctypes.byref(obj)


def create_c_int(obj: int):
    return ctypes.c_int(obj)


def create_c_float(obj: float):
    return ctypes.c_float(obj)


def create_c_double(obj: float):
    return ctypes.c_double(obj)


def create_c_bool(obj: bool):
    return ctypes.c_bool(obj)


def make_safe_array_long_input(lista):
    if not lista or len(lista) == 0:
        var = automation.VARIANT()
        var.vt = automation.VT_EMPTY
        return var
    return automation._midlSAFEARRAY(ctypes.c_long).create(lista)


def make_safe_array_double_input(lista):
    if not lista or len(lista) == 0:
        var = automation.VARIANT()
        var.vt = automation.VT_EMPTY
        return var
    return automation._midlSAFEARRAY(ctypes.c_double).create(lista)


def make_safe_array_string_input(lista):
    if not lista or len(lista) == 0:
        var = automation.VARIANT()
        var.vt = automation.VT_EMPTY
        return var
    return automation._midlSAFEARRAY(automation.BSTR).create(lista)


def create_bstr(initial_value=""):
    return automation.BSTR(initial_value)


APICALL = {"file": "", "geometry": "Geometry", "view": "View", "support": "Support"}
