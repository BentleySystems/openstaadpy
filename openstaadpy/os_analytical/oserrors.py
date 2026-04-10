# ---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
# ---------------------------------------------------------------------------------------------
class OsErrorBase(Exception):
    """
    Base class for all custom OpenSTAAD-related errors.

    All specific error types should inherit from this class.

    :param message: Description of the error.
    :type message: str
    :param code: Error code associated with the error.
    :type code: int
    """

    def __init__(self, message, code):
        super().__init__(f"[{code}] {message}")
        self.code = code


# General Errors


class OsError(OsErrorBase):
    """
    Raised for a general, unspecified OpenSTAAD error.
    """

    def __init__(self):
        super().__init__("General error.", -1)


class OsUnableToCreateProp(OsErrorBase):
    """
    Raised when a property cannot be created in OpenSTAAD.
    """

    def __init__(self):
        super().__init__("Unable to create property.", 0)


class OsInvalidArgument(OsErrorBase):
    """
    Raised when an invalid argument is provided to an OpenSTAAD function.
    """

    def __init__(self):
        super().__init__("Invalid argument.", -100)


class OsMultidimArrayExpected(OsErrorBase):
    """
    Raised when a multidimensional array is expected but not provided.
    """

    def __init__(self):
        super().__init__("Multidimensional array expected.", -106)


class OsArrayExpected(OsErrorBase):
    """
    Raised when an array is expected but not provided.
    """

    def __init__(self):
        super().__init__("Array expected.", -107)


class OsArraySizeSmall(OsErrorBase):
    """
    Raised when an array is too small for the operation.
    """

    def __init__(self):
        super().__init__("Array size too small.", -108)


class OsArraySizeZero(OsErrorBase):
    """
    Raised when an array of size zero is encountered.
    """

    def __init__(self):
        super().__init__("Array size is zero.", -109)


class OsNoBeamPlateSolidSelected(OsErrorBase):
    """
    Raised when no beam, plate, or solid is selected in the model.
    """

    def __init__(self):
        super().__init__("No beam, plate, or solid selected.", -110)


class OsDoubleExpected(OsErrorBase):
    """
    Raised when a double (float) value is expected but not provided.
    """

    def __init__(self):
        super().__init__("Double value expected.", -112)


class OsIntegerExpected(OsErrorBase):
    """
    Raised when an integer value is expected but not provided.
    """

    def __init__(self):
        super().__init__("Integer value expected.", -113)


class OsArraySizeLessThanReqd(OsErrorBase):
    """
    Raised when an array is smaller than the required size.
    """

    def __init__(self):
        super().__init__("Array size less than required.", -125)


# File Errors


class OsFileAlreadyExists(OsErrorBase):
    """
    Raised when a file already exists and the operation cannot proceed.
    """

    def __init__(self):
        super().__init__("File already exists.", -1003)


# Node Errors


class OsNodeNotFound(OsErrorBase):
    """
    Raised when a node is not found in the model.
    """

    def __init__(self):
        super().__init__("Node not found.", -2001)


class OsErrorAddNode(OsErrorBase):
    """
    Raised when an error occurs while adding a node.
    """

    def __init__(self):
        super().__init__("Error adding node.", -2004)


class OsNoNodeSelected(OsErrorBase):
    """
    Raised when no node is selected in the model.
    """

    def __init__(self):
        super().__init__("No node selected.", -2005)


class OsInvalidNodeNo(OsErrorBase):
    """
    Raised when an invalid node number is provided.
    """

    def __init__(self):
        super().__init__("Invalid node number.", -2006)


# Beam Errors


class OsCannotFindMember(OsErrorBase):
    """
    Raised when a member (beam) cannot be found in the model.
    """

    def __init__(self):
        super().__init__("Cannot find member.", 0)


class OsBeamNotFound(OsErrorBase):
    """
    Raised when a beam is not found in the model.
    """

    def __init__(self):
        super().__init__("Beam not found.", -3001)


class OsBeamAlreadyExists(OsErrorBase):
    """
    Raised when a beam already exists and the operation cannot proceed.
    """

    def __init__(self):
        super().__init__("Beam already exists.", -3002)


class OsIdenticalBeamAlreadyExists(OsErrorBase):
    """
    Raised when an identical beam already exists in the model.
    """

    def __init__(self):
        super().__init__("Identical beam already exists.", -3003)


class OsErrorAddBeam(OsErrorBase):
    """
    Raised when an error occurs while adding a beam.
    """

    def __init__(self):
        super().__init__("Error adding beam.", -3004)


class OsNoBeamSelected(OsErrorBase):
    """
    Raised when no beam is selected in the model.
    """

    def __init__(self):
        super().__init__("No beam selected.", -3005)


# Plate Errors


class OsPlateNotFound(OsErrorBase):
    """
    Raised when a plate is not found in the model.
    """

    def __init__(self):
        super().__init__("Plate not found.", -4001)


class OsErrorAddPlate(OsErrorBase):
    """
    Raised when an error occurs while adding a plate.
    """

    def __init__(self):
        super().__init__("Error adding plate.", -4004)


class OsNoPlateSelected(OsErrorBase):
    """
    Raised when no plate is selected in the model.
    """

    def __init__(self):
        super().__init__("No plate selected.", -4005)


class OsInvalidPlateNo(OsErrorBase):
    """
    Raised when an invalid plate number is provided.
    """

    def __init__(self):
        super().__init__("Invalid plate number.", -4006)


class OsInvalidPlateNoFound(OsErrorBase):
    """
    Raised when an invalid plate number is found in the model.
    """

    def __init__(self):
        super().__init__("Invalid plate number found.", -4008)


class OsNoValidPlateNoFound(OsErrorBase):
    """
    Raised when no valid plate number is found in the model.
    """

    def __init__(self):
        super().__init__("No valid plate number found.", -4009)


# Solid Errors


class OsSolidNotFound(OsErrorBase):
    """
    Raised when a solid is not found in the model.
    """

    def __init__(self):
        super().__init__("Solid not found.", -5001)


class OsErrorAddSolid(OsErrorBase):
    """
    Raised when an error occurs while adding a solid.
    """

    def __init__(self):
        super().__init__("Error adding solid.", -5004)


class OsNoSolidSelected(OsErrorBase):
    """
    Raised when no solid is selected in the model.
    """

    def __init__(self):
        super().__init__("No solid selected.", -5005)


class OsMeshNotFound(OsErrorBase):
    """
    Raised when a mesh is not found in the model.
    """

    def __init__(self):
        super().__init__("Mesh not found.", -5603)


class OsInvalidPropRef(OsErrorBase):
    """
    Raised when an invalid property reference is provided.
    """

    def __init__(self):
        super().__init__("Invalid property reference.", -6001)


class OsLibErrorCreateProp(OsErrorBase):
    """
    Raised when a library error occurs while creating a property.
    """

    def __init__(self):
        super().__init__("Library error creating property.", -6003)


class OsProfileNotFound(OsErrorBase):
    """
    Raised when a profile is not found in the property library.
    """

    def __init__(self):
        super().__init__("Profile not found.", -6004)


class OsProfileDataNotFound(OsErrorBase):
    """
    Raised when profile data is not found in the property library.
    """

    def __init__(self):
        super().__init__("Profile data not found.", -6005)


class OsInvalidSectionType(OsErrorBase):
    """
    Raised when an invalid section type is provided.
    """

    def __init__(self):
        super().__init__("Invalid section type.", -6006)


class OsInvalidAssignType(OsErrorBase):
    """
    Raised when an invalid assign type is provided.
    """

    def __init__(self):
        super().__init__("Invalid assign type.", -6008)


class OsLibErrorBetaAssign(OsErrorBase):
    """
    Raised when a library error occurs while assigning beta.
    """

    def __init__(self):
        super().__init__("Library error assigning beta.", -6009)


class OsNoPropAttached(OsErrorBase):
    """
    Raised when no property is attached to an element.
    """

    def __init__(self):
        super().__init__("No property attached.", -6022)


class OsMaterialNotFound(OsErrorBase):
    """
    Raised when a material is not found in the property library.
    """

    def __init__(self):
        super().__init__("Material not found.", -6023)


class OsNoPropDefined(OsErrorBase):
    """
    Raised when no property is defined in the model.
    """

    def __init__(self):
        super().__init__("No property defined.", -6025)


class OsUptCreateFailed(OsErrorBase):
    """
    Raised when UPT (User Provided Table) creation fails.
    """

    def __init__(self):
        super().__init__("UPT creation failed.", -6031)


class OsAddUptSectionFailed(OsErrorBase):
    """
    Raised when adding a UPT section fails.
    """

    def __init__(self):
        super().__init__("Add UPT section failed.", -6032)


class OsUptSectionExists(OsErrorBase):
    """
    Raised when a UPT section already exists.
    """

    def __init__(self):
        super().__init__("UPT section already exists.", -6045)


# Group Errors
class OsGroupAlreadyExists(OsErrorBase):
    """
    Raised when a group already exists in the model.
    """

    def __init__(self):
        super().__init__("Group already exists.", -7001)


# Load Errors
class OsInvalidLoadDirection(OsErrorBase):
    """
    Raised when an invalid load direction is provided.
    """

    def __init__(self):
        super().__init__("Invalid load direction.", -8001)


class OsLoadCaseNotFound(OsErrorBase):
    """
    Raised when a load case is not found in the model.
    """

    def __init__(self):
        super().__init__("Load case not found.", -8002)


class OsCreateLoadFailed(OsErrorBase):
    """
    Raised when creating a load fails.
    """

    def __init__(self):
        super().__init__("Create load failed.", -8004)


class OsLoadExists(OsErrorBase):
    """
    Raised when a load already exists in the model.
    """

    def __init__(self):
        super().__init__("Load already exists.", -8029)


class OsSeismicCodeNotFound(OsErrorBase):
    """
    Raised when a seismic code is not found in the model.
    """

    def __init__(self):
        super().__init__("Seismic code not found.", -8034)


class OsInvalidLoadDefId(OsErrorBase):
    """
    Raised when an invalid load definition ID is provided.
    """

    def __init__(self):
        super().__init__("Invalid load definition ID.", -8039)


class OsInvalidLoadCombName(OsErrorBase):
    """
    Raised when an invalid load combination name is provided.
    """

    def __init__(self):
        super().__init__("Invalid load combination name.", -8040)


class OsInvalidLoadCombCategory(OsErrorBase):
    """
    Raised when an invalid load combination category is provided.
    """

    def __init__(self):
        super().__init__("Invalid load combination category.", -8041)


# Results Errors
class OsBeamForcesNotLoaded(OsErrorBase):
    """
    Raised when beam forces are not loaded in the results.
    """

    def __init__(self):
        super().__init__("Beam forces not loaded.", -9004)


class OsNoGnlResultSet(OsErrorBase):
    """
    Raised when no GNL (Geometric Nonlinear) result set is found.
    """

    def __init__(self):
        super().__init__("No GNL result set.", -9911)


# Miscellaneous
class OsMemberUpdated(OsErrorBase):
    """
    Raised when a member is updated in the model.
    """

    def __init__(self):
        super().__init__("Member updated.", 3007)


class OsPmemberNotFound(OsErrorBase):
    """
    Raised when a physical member is not found in the model.
    """

    def __init__(self):
        super().__init__("Physical member not found.", -5701)


def raise_os_error_if_error_code(code):
    """
    Given an error code, return and raise the corresponding OS error exception.
    If the code is not found, raises a generic OsErrorBase.
    """
    code_map = {
        -1: OsError,
        -100: OsInvalidArgument,
        -106: OsMultidimArrayExpected,
        -107: OsArrayExpected,
        -108: OsArraySizeSmall,
        -109: OsArraySizeZero,
        -110: OsNoBeamPlateSolidSelected,
        -112: OsDoubleExpected,
        -113: OsIntegerExpected,
        -125: OsArraySizeLessThanReqd,
        -1003: OsFileAlreadyExists,
        -2001: OsNodeNotFound,
        -2004: OsErrorAddNode,
        -2005: OsNoNodeSelected,
        -2006: OsInvalidNodeNo,
        -3001: OsBeamNotFound,
        -3002: OsBeamAlreadyExists,
        -3003: OsIdenticalBeamAlreadyExists,
        -3004: OsErrorAddBeam,
        -3005: OsNoBeamSelected,
        -4001: OsPlateNotFound,
        -4004: OsErrorAddPlate,
        -4005: OsNoPlateSelected,
        -4006: OsInvalidPlateNo,
        -4008: OsInvalidPlateNoFound,
        -4009: OsNoValidPlateNoFound,
        -5001: OsSolidNotFound,
        -5004: OsErrorAddSolid,
        -5005: OsNoSolidSelected,
        -5603: OsMeshNotFound,
        -6001: OsInvalidPropRef,
        -6003: OsLibErrorCreateProp,
        -6004: OsProfileNotFound,
        -6005: OsProfileDataNotFound,
        -6006: OsInvalidSectionType,
        -6008: OsInvalidAssignType,
        -6009: OsLibErrorBetaAssign,
        -6022: OsNoPropAttached,
        -6023: OsMaterialNotFound,
        -6025: OsNoPropDefined,
        -6031: OsUptCreateFailed,
        -6032: OsAddUptSectionFailed,
        -6045: OsUptSectionExists,
        -7001: OsGroupAlreadyExists,
        -8001: OsInvalidLoadDirection,
        -8002: OsLoadCaseNotFound,
        -8004: OsCreateLoadFailed,
        -8029: OsLoadExists,
        -8034: OsSeismicCodeNotFound,
        -8039: OsInvalidLoadDefId,
        -8040: OsInvalidLoadCombName,
        -8041: OsInvalidLoadCombCategory,
        -9004: OsBeamForcesNotLoaded,
        -9911: OsNoGnlResultSet,
        -5701: OsPmemberNotFound,
    }
    error_cls = code_map.get(code, None)
    if error_cls is not None:
        raise error_cls()
