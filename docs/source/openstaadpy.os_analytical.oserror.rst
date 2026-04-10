Errors
======

.. list-table::
   :header-rows: 1

   * - Error Class
     - Description
     - Error Code
   * - OsErrorBase
     - Base class for all custom OpenSTAAD-related errors.
     - (base class, no code)
   * - OsError
     - Raised for a general, unspecified OpenSTAAD error.
     - -1
   * - OsUnableToCreateProp
     - Raised when a property cannot be created in OpenSTAAD.
     - 0
   * - OsInvalidArgument
     - Raised when an invalid argument is provided to an OpenSTAAD function.
     - -100
   * - OsMultidimArrayExpected
     - Raised when a multidimensional array is expected but not provided.
     - -106
   * - OsArrayExpected
     - Raised when an array is expected but not provided.
     - -107
   * - OsArraySizeSmall
     - Raised when an array is too small for the operation.
     - -108
   * - OsArraySizeZero
     - Raised when an array of size zero is encountered.
     - -109
   * - OsNoBeamPlateSolidSelected
     - Raised when no beam, plate, or solid is selected in the model.
     - -110
   * - OsDoubleExpected
     - Raised when a double (float) value is expected but not provided.
     - -112
   * - OsIntegerExpected
     - Raised when an integer value is expected but not provided.
     - -113
   * - OsArraySizeLessThanReqd
     - Raised when an array is smaller than the required size.
     - -125
   * - OsFileAlreadyExists
     - Raised when a file already exists and the operation cannot proceed.
     - -1003
   * - OsNodeNotFound
     - Raised when a node is not found in the model.
     - -2001
   * - OsErrorAddNode
     - Raised when an error occurs while adding a node.
     - -2004
   * - OsNoNodeSelected
     - Raised when no node is selected in the model.
     - -2005
   * - OsInvalidNodeNo
     - Raised when an invalid node number is provided.
     - -2006
   * - OsCannotFindMember
     - Raised when a member (beam) cannot be found in the model.
     - 0
   * - OsBeamNotFound
     - Raised when a beam is not found in the model.
     - -3001
   * - OsBeamAlreadyExists
     - Raised when a beam already exists and the operation cannot proceed.
     - -3002
   * - OsIdenticalBeamAlreadyExists
     - Raised when an identical beam already exists in the model.
     - -3003
   * - OsErrorAddBeam
     - Raised when an error occurs while adding a beam.
     - -3004
   * - OsNoBeamSelected
     - Raised when no beam is selected in the model.
     - -3005
   * - OsPlateNotFound
     - Raised when a plate is not found in the model.
     - -4001
   * - OsErrorAddPlate
     - Raised when an error occurs while adding a plate.
     - -4004
   * - OsNoPlateSelected
     - Raised when no plate is selected in the model.
     - -4005
   * - OsInvalidPlateNo
     - Raised when an invalid plate number is provided.
     - -4006
   * - OsInvalidPlateNoFound
     - Raised when an invalid plate number is found in the model.
     - -4008
   * - OsNoValidPlateNoFound
     - Raised when no valid plate number is found in the model.
     - -4009
   * - OsSolidNotFound
     - Raised when a solid is not found in the model.
     - -5001
   * - OsErrorAddSolid
     - Raised when an error occurs while adding a solid.
     - -5004
   * - OsNoSolidSelected
     - Raised when no solid is selected in the model.
     - -5005
   * - OsMeshNotFound
     - Raised when a mesh is not found in the model.
     - -5603
   * - OsInvalidPropRef
     - Raised when an invalid property reference is provided.
     - -6001
   * - OsLibErrorCreateProp
     - Raised when a library error occurs while creating a property.
     - -6003
   * - OsProfileNotFound
     - Raised when a profile is not found in the property library.
     - -6004
   * - OsProfileDataNotFound
     - Raised when profile data is not found in the property library.
     - -6005
   * - OsInvalidSectionType
     - Raised when an invalid section type is provided.
     - -6006
   * - OsInvalidAssignType
     - Raised when an invalid assign type is provided.
     - -6008
   * - OsLibErrorBetaAssign
     - Raised when a library error occurs while assigning beta.
     - -6009
   * - OsNoPropAttached
     - Raised when no property is attached to an element.
     - -6022
   * - OsMaterialNotFound
     - Raised when a material is not found in the property library.
     - -6023
   * - OsNoPropDefined
     - Raised when no property is defined in the model.
     - -6025
   * - OsUptCreateFailed
     - Raised when UPT (User Provided Table) creation fails.
     - -6031
   * - OsAddUptSectionFailed
     - Raised when adding a UPT section fails.
     - -6032
   * - OsUptSectionExists
     - Raised when a UPT section already exists.
     - -6045
   * - OsGroupAlreadyExists
     - Raised when a group already exists in the model.
     - -7001
   * - OsInvalidLoadDirection
     - Raised when an invalid load direction is provided.
     - -8001
   * - OsLoadCaseNotFound
     - Raised when a load case is not found in the model.
     - -8002
   * - OsCreateLoadFailed
     - Raised when creating a load fails.
     - -8004
   * - OsLoadExists
     - Raised when a load already exists in the model.
     - -8029
   * - OsSeismicCodeNotFound
     - Raised when a seismic code is not found in the model.
     - -8034
   * - OsInvalidLoadDefId
     - Raised when an invalid load definition ID is provided.
     - -8039
   * - OsInvalidLoadCombName
     - Raised when an invalid load combination name is provided.
     - -8040
   * - OsInvalidLoadCombCategory
     - Raised when an invalid load combination category is provided.
     - -8041
   * - OsBeamForcesNotLoaded
     - Raised when beam forces are not loaded in the results.
     - -9004
   * - OsNoGnlResultSet
     - Raised when no GNL (Geometric Nonlinear) result set is found.
     - -9911
   * - OsMemberUpdated
     - Raised when a member is updated in the model.
     - 3007
   * - OsPmemberNotFound
     - Raised when a physical member is not found in the model.
     - -5701
