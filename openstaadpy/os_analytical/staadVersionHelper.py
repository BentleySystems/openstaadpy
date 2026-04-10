# ---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
# ---------------------------------------------------------------------------------------------


# --- Simple version comparison utility ---
def compare_versions(version1: str, version2: str) -> int:
    """
    Compare two version strings (e.g., '25.1.0.2' vs '25.1.0.10').
    Returns:
        1 if version1 > version2
        0 if version1 == version2
       -1 if version1 < version2
    """
    v1 = [int(x) for x in version1.split(".")]
    v2 = [int(x) for x in version2.split(".")]
    # Pad to 4 parts
    while len(v1) < 4:
        v1.append(0)
    while len(v2) < 4:
        v2.append(0)
    for a, b in zip(v1, v2):
        if a > b:
            return 1
        elif a < b:
            return -1
    return 0


# Example usage:
# result = compare_versions('25.1.0.2', '25.1.0.10')
# if result == 1:
#     print('First version is newer')
# elif result == -1:
#     print('Second version is newer')
# else:
#     print('Versions are equal')
