# ---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
# ---------------------------------------------------------------------------------------------
__all__ = ["connect", "oserrors"]

from .openstaadroot import getActiveObject as _getActiveObject


def connect(filePath: str = None):
    """
    Connect to STAAD.Pro via OpenSTAAD Python API.
    Returns an OSRoot object for all STAAD operations.
    """
    return _getActiveObject(filePath)
