# ---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
# ---------------------------------------------------------------------------------------------
from __future__ import annotations

__all__ = ["connect", "oserrors"]

from typing import Optional

from .openstaadroot import getActiveObject as _getActiveObject


def connect(filePath: Optional[str] = None):
    """
    Connect to STAAD.Pro via OpenSTAAD Python API.
    Returns an OSRoot object for all STAAD operations.
    """
    return _getActiveObject(filePath)
