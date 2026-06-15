# ---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
# ---------------------------------------------------------------------------------------------
from .openStaadHelper import (
    create_variant_int,
    make_safe_array_double,
    make_safe_array_double_input,
    make_safe_array_long,
    make_safe_array_long_input,
    make_safe_array_string_input,
    make_variant_vt_ref,
)
from .oserrors import raise_os_error_if_error_code
from comtypes import automation
from comtypes import CoInitialize


class OSLoad:
    CoInitialize()

    def __init__(self, staadObj):
        self._staad = staadObj
        self._load = self._staad.Load

        self._functions = [
            "CreateNewPrimaryLoad",
            "CreateNewLoadCombination",
            "CreateNewReferenceLoad",
            "CreateLoadEnvelop",
            "CreateLoadList",
            "CreateNewPrimaryLoadEx",
            "CreateNewPrimaryLoadEx2",
            "SetLoadActive",
            "SetReferenceLoadActive",
            "SetLoadType",
            "SetASDLoadAttribute",
            "SetLSDLoadAttribute",
            "AddSelfWeightInXYZ",
            "AddSelfWeightInXYZToGeometry",
            "AddNodalLoad",
            "AddSupportDisplacement",
            "AddMemberUniformForce",
            "AddMemberUniformMoment",
            "AddMemberConcForce",
            "AddMemberConcMoment",
            "AddMemberLinearVari",
            "AddMemberTrapezoidal",
            "AddMemberAreaLoad",
            "AddMemberFixedEnd",
            "AddElementPressure",
            "AddElementPressure",
            "AddElementHydrostaticPressure",
            "AddTemperatureLoad",
            "AddStrainLoad",
            "AddLoadAndFactorToCombination",
            "AddMemberFloorLoad",
            "AddMemberFloorLoadEx",
            "AddElementTrapPressureEx",
            "AddWindDefinition",
            "AddWindIntensity",
            "AddWindExposure",
            "AddWindLoad",
            "AddSeismicDefinition",
            "AddSeismicDefSelfWeight",
            "AddSeismicDefMemberWeight",
            "AddSeismicDefJointWeight",
            "AddSeismicDefElementWeight",
            "AddSeismicDefFloorWeight",
            "AddSeismicLoad",
            "AddAutoLoadCombinations",
            "AddRepeatLoad",
            "AddLoadCasesToEnvelop",
            "AddReferenceLoad",
            "AddSeismicDefWallArea",
            "AddWindDefinitionASCE7Parameters",
            "AddNotionalLoad",
            "AddDirectAnalysisDefinitionParameter",
            "AddResponseSpectrumLoadEx",
            "AddAutoCombinationRepeat",
            "RemoveLoadCasesFromEnvelop",
            "RemoveAttribute",
            "ClearPrimaryLoadCase",
            "ClearReferenceLoadCase",
            "IsDynamicLoadIncluded",
            "IsCombinationCase",
            "SplitLoadsOnBeam",
            "MergeLoadsOnBeam",
            "BeginLoadMerging",
            "EndLoadMerging",
            "ModifySeismicDefinitionParams",
            "ComputeWallWindPressureProfile",
            "ComputeWallWindPressureProfileASCE72016",
            "DeleteLoadEnvelop",
            "DeleteLoadList",
            "DeletePrimaryLoadCases",
            "DeleteReferenceLoadCases",
            "DeleteWindDefinition",
            "DeleteDirectAnalysisDefinitionParameter",
            "DeleteDirectAnalysisDefinition",
            "GetPrimaryLoadCaseCount",
            "GetPrimaryLoadCaseNumbers",
            "GetLoadCombinationCaseCount",
            "GetLoadCombinationCaseNumbers",
            "GetReferenceLoadCount",
            "GetReferenceLoadCaseCount",
            "GetReferenceLoadCaseNumbers",
            "GetNoOfSetsInReferenceLoad",
            "GetReferenceLoadByIndex",
            "GetReferenceLoadType",
            "GetReferenceLoadCaseTitle",
            "GetBeamCountAtFloor",
            "GetInfluenceArea",
            "GetActiveLoad",
            "GetNodalLoadCount",
            "GetNodalLoads",
            "GetUDLLoadCount",
            "GetUDLLoads",
            "GetUNIMomentCount",
            "GetUNIMoments",
            "GetTrapLoadCount",
            "GetTrapLoads",
            "GetConcForceCount",
            "GetConcForces",
            "GetConcMomentCount",
            "GetConcMoments",
            "GetNoOfLoadAndFactorPairsForCombination",
            "GetLoadAndFactorForCombination",
            "GetLoadCaseTitle",
            "GetElementPressureLoadCount",
            "GetElementPressureLoads",
            "GetElementConcLoadCount",
            "GetElementConcLoads",
            "GetLoadType",
            "GetLoadListCount",
            "GetLoadCountInLoadList",
            "GetLoadsInLoadList",
            "GetAttribute",
            "GetLoadType",
            "GetRepeatLoadCount",
            "GetNoLoadFactorInRepeatLoad",
            "GetRepeatLoadByIndex",
            "GetLinearVaryingLoadCount",
            "GetLinearVaryingLoads",
            "GetLoadTypeCount",
            "GetListSizeForLoadType",
            "GetAssignmentListForLoadType",
            "GetNodalLoadInfo",
            "GetMemberLoadInfo",
            "GetElementLoadInfo",
            "GetNotionalLoadCount",
            "GetNoLoadFactorDirectionInNotionalLoad",
            "GetNotionalLoadByIndex",
            "GetLoadItemsCount",
            "GetLoadItemType",
            "GetEnvelopeCount",
            "GetLoadEnvelopeDetails",
            "GetLoadListfromLoadEnvelope",
            "GetEnvelopeIDs",
        ]

        for function_name in self._functions:
            self._load._FlagAsMethod(function_name)

    # SUPPORT FUNCTIONS

    def CreateNewPrimaryLoad(self, primaryLoadTitle: str):
        """
        Creates a new PRIMARY load case.

        Parameters
        ----------
        primaryLoadTitle : str
            Title of the primary load case.

        Returns
        -------
        int
            Load number ID if the load case is created successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> load_id = staad_obj.load.CreateNewPrimaryLoad("Dead Load")
        >>> print(load_id)
        """

        retVal = self._load.CreateNewPrimaryLoad(primaryLoadTitle)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def CreateNewLoadCombination(self, loadCombTitle: str, loadCombNo: int):
        """
        Creates a new load combination case.

        Parameters
        ----------
        loadCombTitle : str
            Title of the load combination.
        loadCombNo : int
            Load combination number.

        Returns
        -------
        int
            Load number ID assigned to the load combination.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> comb_id = staad_obj.Load.CreateNewLoadCombination("DL+LL", 2)
        >>> print(comb_id)
        """

        retVal = self._load.CreateNewLoadCombination(loadCombTitle, loadCombNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def CreateNewReferenceLoad(
        self, nodeNo: int, referenceLoadCaseTitle: str, loadType: int
    ):
        """
        Creates a new reference load case.

        Parameters
        ----------
        nodeNo : int
            Reference ID to be assigned to the new reference load case.
        referenceLoadCaseTitle : str
            Title of the reference load case.
        loadType : int
            Type of load.

        Returns
        -------
        int
            Reference load case number ID.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> ref_id = staad_obj.Load.CreateNewReferenceLoad(1, "Ref Load", 0)
        >>> print(ref_id)
        """
        retVal = self._load.CreateNewReferenceLoad(
            nodeNo, referenceLoadCaseTitle, loadType
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def CreateLoadEnvelop(
        self, envelopNumber: int, envelopType: int, loadCaseList: list
    ):
        """
        Creates a Load Envelop with specified primary load case(s) and envelop type.

        Parameters
        ----------
        envelopNumber : int
            Load Envelop reference ID
        envelopType : int
            Type of the load envelop:
                +-------+------------------+
                | Value |Load Envelop Type |
                +=======+==================+
                | 0     |NONE              |
                +-------+------------------+
                | 1     |STRESS            |
                +-------+------------------+
                | 2     |SERVICEABILITY    |
                +-------+------------------+
                | 3     |COLUMN            |
                +-------+------------------+
                | 4     |CONNECTION        |
                +-------+------------------+
                | 5     |STRENGTH          |
                +-------+------------------+
                | 6     |TEMPORARY         |
                +-------+------------------+
        loadCaseList : list of int
            Load Case IDs for which to create a load envelop

        Returns
        -------
        bool
            True OK.
            False General error.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.CreateLoadEnvelop(1, 1, [1,2,3])
        >>> print(result)
        """
        safe_LoadCaseList = make_safe_array_long_input(loadCaseList)
        return self._load.CreateLoadEnvelop(
            envelopNumber, envelopType, safe_LoadCaseList
        )

    def CreateLoadList(self, listType: int, loadCaseList: list[int]):
        """
        Creates a load list.

        Parameters
        ----------
        listType : int
            Load list type: 0 and 1 for load list and load envelope list, respectively.
        loadCaseList : list of int
            Load Case reference IDs for which to create a load envelop

        Returns
        -------
        bool
            True if the load list was created successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.CreateLoadList(0, [1,2])
        >>> print(result)
        """
        safe_LoadCaseList = make_safe_array_long_input(loadCaseList)
        retval = self._load.CreateLoadList(listType, safe_LoadCaseList)
        return bool(retval)

    def CreateNewPrimaryLoadEx(self, primaryLoadTitle: str, loadType: int):
        """
        Creates new PRIMARY load case.

        Parameters
        ----------
        primaryLoadTitle : string
            The primary load case string title.
        loadType : int
            Type of the load:
                +-------+-----------+-------+-------------+
                | Value | Load Type | Value | Load Type   |
                +=======+===========+=======+=============+
                | 0     | Dead      | 12    | Traffic     |
                +-------+-----------+-------+-------------+
                | 1     | Live      | 13    | Temp        |
                +-------+-----------+-------+-------------+
                | 2     | Roof Live | 14    | Imperfection|
                +-------+-----------+-------+-------------+
                | 3     | Wind      | 15    | Accidental  |
                +-------+-----------+-------+-------------+
                | 4     | Seismic-H | 16    | Flood       |
                +-------+-----------+-------+-------------+
                | 5     | Seismic-V | 17    | Ice         |
                +-------+-----------+-------+-------------+
                | 6     | Snow      | 18    | Wind Ice    |
                +-------+-----------+-------+-------------+
                | 7     | Fluids    | 19    | Crane Hook  |
                +-------+-----------+-------+-------------+
                | 8     | Soil      | 20    | Mass        |
                +-------+-----------+-------+-------------+
                | 9     | Rain      | 21    | Gravity     |
                +-------+-----------+-------+-------------+
                | 10    | Ponding   | 22    | Push        |
                +-------+-----------+-------+-------------+
                | 11    | Dust      | 23    | None        |
                +-------+-----------+-------+-------------+

        Returns
        -------
        int
            Returns load Number of newly created Primary load Case.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> load_id = staad_obj.Load.CreateNewPrimaryLoadEx("Live Load", 1)
        >>> print(load_id)
        """
        retVal = self._load.CreateNewPrimaryLoadEx(primaryLoadTitle, loadType)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def CreateNewPrimaryLoadEx2(
        self, primaryLoadTitle: str, loadType: int, loadCaseNo: int
    ):
        """
        Creates new PRIMARY load case.

        Parameters
        ----------
        primaryLoadTitle : string
            The primary load case string title.
        loadType : int
            Type of the load:
                +-------+-----------+-------+-------------+
                | Value | Load Type | Value | Load Type   |
                +=======+===========+=======+=============+
                | 0     | Dead      | 12    | Traffic     |
                +-------+-----------+-------+-------------+
                | 1     | Live      | 13    | Temp        |
                +-------+-----------+-------+-------------+
                | 2     | Roof Live | 14    | Imperfection|
                +-------+-----------+-------+-------------+
                | 3     | Wind      | 15    | Accidental  |
                +-------+-----------+-------+-------------+
                | 4     | Seismic-H | 16    | Flood       |
                +-------+-----------+-------+-------------+
                | 5     | Seismic-V | 17    | Ice         |
                +-------+-----------+-------+-------------+
                | 6     | Snow      | 18    | Wind Ice    |
                +-------+-----------+-------+-------------+
                | 7     | Fluids    | 19    | Crane Hook  |
                +-------+-----------+-------+-------------+
                | 8     | Soil      | 20    | Mass        |
                +-------+-----------+-------+-------------+
                | 9     | Rain      | 21    | Gravity     |
                +-------+-----------+-------+-------------+
                | 10    | Ponding   | 22    | Push        |
                +-------+-----------+-------+-------------+
                | 11    | Dust      | 23    | None        |
                +-------+-----------+-------+-------------+

        loadCaseNo : int
            The load case number.

        Returns
        -------
        int
            Returns load Case number of newly created Primary load Case.
            Returns 0 if not successfully

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> load_id = staad_obj.Load.CreateNewPrimaryLoadEx2("Wind", 3, 5)
        >>> print(load_id)
        """
        return self._load.CreateNewPrimaryLoadEx2(
            primaryLoadTitle, loadType, loadCaseNo
        )

    def SetLoadActive(self, loadNumber: int):
        """
        Activates the specified load number to allow adding or removing load items.

        Parameters
        ----------
        loadNumber : int
            Load case reference number ID.

        Returns
        -------
        bool
            True if the load case was successfully activated.
            False if an error occurred.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.SetLoadActive(1)
        >>> print(result)
        """
        return self._load.SetLoadActive(loadNumber)

    def SetReferenceLoadActive(self, nLoadCaseNo: int):
        """
        Activates a reference load case to allow operations on its items.

        Parameters
        ----------
        nLoadCaseNo : int
            Reference load case ID in Load Case Details.

        Returns
        -------
        int
            Reference load case number ID.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.SetReferenceLoadActive(2)
        >>> print(result)
        """
        retVal = self._load.SetReferenceLoadActive(nLoadCaseNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def SetLoadType(self, loadCaseNumber: int, loadType: int):
        """
        Set load type to load case for considering load combination.

        Parameters
        ----------
        loadCaseNumber : int
            The load case reference number ID.
        loadType : int
            Type of the load.:
                +-------+-----------+-------+-------------+
                | Value | Load Type | Value | Load Type   |
                +=======+===========+=======+=============+
                | 0     | Dead      | 12    | Traffic     |
                +-------+-----------+-------+-------------+
                | 1     | Live      | 13    | Temp        |
                +-------+-----------+-------+-------------+
                | 2     | Roof Live | 14    | Imperfection|
                +-------+-----------+-------+-------------+
                | 3     | Wind      | 15    | Accidental  |
                +-------+-----------+-------+-------------+
                | 4     | Seismic-H | 16    | Flood       |
                +-------+-----------+-------+-------------+
                | 5     | Seismic-V | 17    | Ice         |
                +-------+-----------+-------+-------------+
                | 6     | Snow      | 18    | Wind Ice    |
                +-------+-----------+-------+-------------+
                | 7     | Fluids    | 19    | Crane Hook  |
                +-------+-----------+-------+-------------+
                | 8     | Soil      | 20    | Mass        |
                +-------+-----------+-------+-------------+
                | 9     | Rain      | 21    | Gravity     |
                +-------+-----------+-------+-------------+
                | 10    | Ponding   | 22    | Push        |
                +-------+-----------+-------+-------------+
                | 11    | Dust      | 23    | None        |
                +-------+-----------+-------+-------------+

        Returns
        -------
        bool
            True if the load type was set successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.SetLoadType(1, 0)
        >>> print(result)
        """
        retVal = self._load.SetLoadType(loadCaseNumber, loadType)
        return bool(retVal)

    def SetASDLoadAttribute(
        self, loadCaseRefID: int, strengthType: int, allowStressIncrease: bool
    ):
        """
        Sets Allowable Stress Design (ASD) load attribute.

        Parameters
        ----------
        loadCaseRefID : int
            Load case reference ID.
        strengthType : int
            Strength Type :
                +------------------------------------------------------+---------+
                | Value                                                | Integer |
                +======================================================+=========+
                | STRENGTH_TYPE_NONE                                   | 0       |
                +------------------------------------------------------+---------+
                | NORMAL_ASD_WORKING_STRESS_FORCES_WITHOUT_P_DELTA     | 1       |
                +------------------------------------------------------+---------+
                | NORMAL_ASD_WORKING_STRESS_FORCES_WITH_P_DELTA        | 2       |
                +------------------------------------------------------+---------+
                | STRENGTH_TYPE_OF_FORCES_WITHOUT_P_DELTA              | 3       |
                +------------------------------------------------------+---------+
                | STRENGTH_TYPE_OF_FORCES_WITH_P_DELTA                 | 4       |
                +------------------------------------------------------+---------+
                | COLUMN_ONLY_STRENGTH_TYPE_OF_FORCES_WITHOUT_P_DELTA  | 5       |
                +------------------------------------------------------+---------+
                | COLUMN_ONLY_STRENGTH_TYPE_OF_FORCES_WITH_P_DELTA     | 6       |
                +------------------------------------------------------+---------+

        allowStressIncrease : bool
            Allow 1/3 stress increase in ASD.

        Returns
        -------
        bool
            True if the ASD load attribute was set successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.SetASDLoadAttribute(1, 1, True)
        >>> print(result)
        """
        retVal = self._load.SetASDLoadAttribute(
            loadCaseRefID, strengthType, allowStressIncrease
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def SetLSDLoadAttribute(self, loadCaseRefID: int):
        """
        Sets Limit State Design (LSD) load attribute.

        Parameters
        ----------
        loadCaseRefID : int
            Load case reference ID.

        Returns
        -------
        bool
            True if the LSD load attribute was set successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.SetLSDLoadAttribute(1)
        >>> print(result)
        """
        retVal = self._load.SetLSDLoadAttribute(loadCaseRefID)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def AddSelfWeightInXYZ(self, varInDirection: int, varLoadFactor: float):
        """
        Adds self-weight to the active load case for all entities (beams, plates, solids).

        Parameters
        ----------
        varInDirection : int
            Direction index for self-weight (1 = X, 2 = Y, 3 = Z).
        varLoadFactor : float
            Multiplying factor for self-weight.

        Returns
        -------
        bool
            True if self-weight was added successfully.
            False if an error occurred.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.AddSelfWeightInXYZ(2, 1.0)
        >>> print(result)
        """
        return self._load.AddSelfWeightInXYZ(varInDirection, varLoadFactor)

    def AddSelfWeightInXYZToGeometry(
        self, varGeomNumberIDs: list, varInDirection: int, varLoadFactor: float
    ):
        """
        Adds self-weight to specified geometry entities in the active load case.

        Parameters
        ----------
        varGeomNumberIDs : list of int
            List of beam, plate, or solid number IDs.
        varInDirection : int
            Direction index for self-weight (1 = X, 2 = Y, 3 = Z).
        varLoadFactor : float
            Multiplying factor for self-weight.

        Returns
        -------
        bool
            True if self-weight was added to the specified geometries.
            False if an error occurred.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.AddSelfWeightInXYZToGeometry([1,2], 3, 0.9)
        >>> print(result)
        """
        safe_GeomMemNoList = make_safe_array_long_input(varGeomNumberIDs)
        return self._load.AddSelfWeightInXYZToGeometry(
            safe_GeomMemNoList, varInDirection, varLoadFactor
        )

    def AddNodalLoad(
        self,
        nodeIds: list,
        forceInXDir: float,
        forceInYDir: float,
        forceInZDir: float,
        momentInXDir: float,
        momentInYDir: float,
        momentInZDir: float,
    ):
        """
        Adds joint load to the specified node numbers.

        Parameters
        ----------
        nodeIds : list of int
            List of node IDs to apply the joint load.
        forceInXDir : float
            Force in the X direction.
        forceInYDir : float
            Force in the Y direction.
        forceInZDir : float
            Force in the Z direction.
        momentInXDir : float
            Moment in the X direction.
        momentInYDir : float
            Moment in the Y direction.
        momentInZDir : float
            Moment in the Z direction.

        Returns
        -------
        bool
            True if the joint load was added successfully.
            False if an error occurred.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.AddNodalLoad([1,2], 10, 0, 0, 0, 0, 0)
        >>> print(result)
        """
        safe_NodeIdList = make_safe_array_long_input(nodeIds)
        retVal = self._load.AddNodalLoad(
            safe_NodeIdList,
            forceInXDir,
            forceInYDir,
            forceInZDir,
            momentInXDir,
            momentInYDir,
            momentInZDir,
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def AddSupportDisplacement(
        self, nodeIds: list, varDirection: int, varDispValue: float
    ):
        """
        Adds support displacement to one or more nodes.

        Parameters
        ----------
        nodeIds : list of int
            List of node IDs.
        varDirection : int
            Direction index (1 = X, 2 = Y, 3 = Z).
        varDispValue : float
            Displacement value in the specified direction.

        Returns
        -------
        bool
            True if the support displacement was added successfully.
            False if an error occurred.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.AddSupportDisplacement([1], 1, 5.0)
        >>> print(result)
        """
        safe_NodeIdList = make_safe_array_long_input(nodeIds)
        retVal = self._load.AddSupportDisplacement(
            safe_NodeIdList, varDirection, varDispValue
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def AddMemberUniformForce(
        self,
        beamIds: list,
        varDirection: int,
        varForce: float,
        varD1: float,
        varD2: float,
        varD3: float,
    ):
        """
        Adds a uniform force to the specified beams.

        Parameters
        ----------
        beamIds : list of int
            List of beam IDs.
        varDirection : int
            Load direction (1 to 9 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ, ProjectedX, ProjectedY, ProjectedZ respectively).
        varForce : float
            Magnitude of the uniform force.
        varD1 : float
            Distance from the start of the member to the start of the load.
        varD2 : float
            Distance from the start of the member to the end of the load.
        varD3 : float
            Perpendicular distance from the member shear center to the local plane of loading.

        Returns
        -------
        bool
            True if the uniform force was added successfully.
            False if an error occurred.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.AddMemberUniformForce([1,2], 1, 5.0, 0, 5, 0)
        >>> print(result)
        """
        safe_BeamIdList = make_safe_array_long_input(beamIds)
        retVal = self._load.AddMemberUniformForce(
            safe_BeamIdList, varDirection, varForce, varD1, varD2, varD3
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def AddMemberUniformMoment(
        self,
        beamIds: list,
        varDirection: int,
        varMoment: float,
        varD1: float,
        varD2: float,
        varD3: float,
    ):
        """
        Adds a uniform moment to the specified beams.

        Parameters
        ----------
        beamIds : list of int
            List of beam IDs.
        varDirection : int
            Load direction (1 to 9 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ, ProjectedX, ProjectedY, ProjectedZ).
        varMoment : float
            Magnitude of the uniform moment.
        varD1 : float
            Distance from the start of the member to the start of the load.
        varD2 : float
            Distance from the start of the member to the end of the load.
        varD3 : float
            Perpendicular distance from the member shear center to the local plane of loading.

        Returns
        -------
        bool
            True if the uniform moment was added successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.AddMemberUniformMoment([1], 2, 10.0, 0, 5, 0)
        >>> print(result)
        """
        safe_BeamIdList = make_safe_array_long_input(beamIds)
        retval = self._load.AddMemberUniformMoment(
            safe_BeamIdList, varDirection, varMoment, varD1, varD2, varD3
        )
        if retval < 0:
            raise_os_error_if_error_code(retval)
        return bool(retval)

    def AddMemberConcForce(
        self,
        beamIds: list,
        varDirection: int,
        varForce: float,
        varD1: float,
        varD2: float,
    ):
        """
        Adds a concentrated force to the specified beams.

        Parameters
        ----------
        beamIds : list of int
            List of beam IDs.
        varDirection : int
            Load direction (1 to 6 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ respectively).
        varForce : float
            Magnitude of the concentrate force in current units.
        varD1 : float
            Distance from the start of the member to concentrated force.
        varD2 : float
            Perpendicular distance from the member shear center to the local plane of loading.

        Returns
        -------
        bool
            True if the concentrated force was added successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.AddMemberConcForce([1], 1, 20.0, 2.5, 0)
        >>> print(result)
        """
        safe_BeamIdList = make_safe_array_long_input(beamIds)
        retVal = self._load.AddMemberConcForce(
            safe_BeamIdList, varDirection, varForce, varD1, varD2
        )
        return bool(retVal)

    def AddMemberConcMoment(
        self,
        beamIds: list,
        varDirection: int,
        varMoment: float,
        varD1: float,
        varD2: float,
    ):
        """
        Adds a concentrated moment to the specified beams.

        Parameters
        ----------
        beamIds : list of int
            List of beam IDs.
        varDirection : int
            Load direction (1 to 6 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ respectively).
        varMoment : float
            Magnitude of the concentrate moment in current units.
        varD1 : float
            Distance from the start of the member to concentrated moment.
        varD2 : float
            Perpendicular distance from the member shear center to the local plane of loading.

        Returns
        -------
        int
            Returns True if successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.AddMemberConcMoment([1], 2, 15.0, 3.0, 0)
        >>> print(result)
        """
        safe_BeamIdList = make_safe_array_long_input(beamIds)
        retval = self._load.AddMemberConcMoment(
            safe_BeamIdList, varDirection, varMoment, varD1, varD2
        )
        if retval < 0:
            raise_os_error_if_error_code(retval)
        return bool(retval)

    def AddMemberLinearVari(
        self,
        memberIds: list[int],
        varDirection: int,
        varW1: float,
        varW2: float,
        varW3: float,
    ):
        """
        Adds LINEARLY VARYING load to beams.

        Parameters
        ----------
        memberIds : list of int
            List of member IDs.
        varDirection : int
            Load direction (1 to 3 for LocalX, LocalY, LocalZ respectively).
        varW1 : float
            Load at the start of the member.
        varW2 : float
            Load at the end of the member.
        varW3 : float
            Load in the middle of the member (for triangular load).

        Returns
        -------
        bool
            True if the linear varying load was added successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.AddMemberLinearVari([1], 2, 2.0, 0.0, 0.0)
        >>> print(result)
        """
        safe_MemberIdList = make_safe_array_long_input(memberIds)
        retVal = self._load.AddMemberLinearVari(
            safe_MemberIdList, varDirection, float(varW1), float(varW2), float(varW3)
        )
        return bool(retVal)

    def AddMemberTrapezoidal(
        self,
        memberIds: list,
        varDirection: int,
        varW1: float,
        varW2: float,
        varD1: float,
        varD2: float,
    ):
        """
        Adds trapezoidal linearly varying load to beams.

        Parameters
        ----------
        memberIds : list of int
            List of member IDs.
        varDirection : int
            Load direction (1 to 9 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ, ProjectedX, ProjectedY, ProjectedZ respectively).
        varW1 : float
            Load at the start of the member.
        varW2 : float
            Load at the end of the member.
        varD1 : float
            Distance from the start of the member to loading starting point.
        varD2 : float
            Distance from the start of the member to loading stopping point.

            Notes:
            - If  varD1 and varD2 are not given, the load is assumed to cover the full member length.

        Returns
        -------
        bool
            True if the trapezoidal load was added successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.Load.AddMemberTrapezoidal([1], 1, 5.0, 10.0, 0, 5)
        >>> print(result)
        """
        safe_MemberIdList = make_safe_array_long_input(memberIds)
        retVal = self._load.AddMemberTrapezoidal(
            safe_MemberIdList, varDirection, varW1, varW2, varD1, varD2
        )
        return bool(retVal)

    def AddMemberAreaLoad(self, beamIds: list, load: float):
        """
        Adds AREA LOAD to beams.

        Parameters
        ----------
        beamIds : list of int
            List of Beam IDs.
        load : float
            Magnitude of the load value.

        Returns
        -------
        bool
            Returns True if successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddMemberAreaLoad([1], 5.0)
        """
        safe_BeamIdList = make_safe_array_long_input(beamIds)
        retVal = self._load.AddMemberAreaLoad(safe_BeamIdList, load)
        return bool(retVal)

    def AddMemberFixedEnd(self, beamIds: list, loadStart: float, loadEnd: float):
        """
        Adds FIXED END LOAD to beams.

        Parameters
        ----------
        beamIds : list of int
            List of Beam IDs.
        loadStart : list of float
            Load at starting point in form of array containing 6 elements corresponding to FX, FY, FZ, MX, MY & MZ respective to each index.
        loadEnd : list of float
            Load at end point in form of array containing 6 elements corresponding to FX, FY, FZ, MX, MY & MZ respective to each index.

        Returns
        -------
        bool
            Returns True if successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddMemberFixedEnd([1], [1.0, 1.0, 1.0, 0, 0, 0], [1.0, 1.0, 1.0, 0, 0, 0])
        """
        safe_BeamIdList = make_safe_array_long_input(beamIds)
        safe_loadStart = make_safe_array_double_input(loadStart)
        safe_loadEnd = make_safe_array_double_input(loadEnd)
        retval = self._load.AddMemberFixedEnd(
            safe_BeamIdList, safe_loadStart, safe_loadEnd
        )
        return bool(retval)

    def AddElementPressure(
        self,
        plateIds: list,
        varDirection: int,
        varPressure: float,
        varX1: float,
        varY1: float,
        varX2: float,
        varY2: float,
    ):
        """
        Adds pressure load to plate elements.

        Parameters
        ----------
        plateIds : list of int
            List of plate IDs.
        varDirection : int
            Load direction: (1 to 9 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ, ProjectedX, ProjectedY and ProjectedZ respectively).
        varPressure : float
            Magnitude of the pressure or concentrate load on the element.
        varX1 : float
            Top-Left coordinate X (local).
        varY1 : float
            Top-Left coordinate Y (local).
        varX2 : float
            Bottom-Right coordinate X (local).
        varY2 : float
            Bottom-Right coordinate Y (local).

        Notes:
        - If X1, Y1, X2 and Y2 are 0, the pressure is applied over the full area of the element.
        - If X1, Y1, X2 and Y2 are not 0: Pressure applied over the area between (X1 , Y1) and (X2 , Y2) measured from the center of plate(s) in the local axis system.
        - If X1 and Y1 are not 0, but X2 and Y2 are 0: Concentrate load applied on (X1 , Y1) measured from the center of plate(s) in the local axis system.

        Returns
        -------
        bool
            True if the pressure load was added successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddElementPressure([1], 3, 5.0, 0.0, 0.0, 1.0, 1.0)
        """
        safe_PlateIdList = make_safe_array_long_input(plateIds)
        retVal = self._load.AddElementPressure(
            safe_PlateIdList, varDirection, varPressure, varX1, varY1, varX2, varY2
        )
        return bool(retVal)

    def AddElementHydrostaticPressure(
        self,
        plateIds: list,
        varLoadDirection: int,
        varInterpolateDirection: int,
        varMinLoad: float,
        varMaxLoad: float,
    ):
        """
        Adds Hydrostatic pressure loading to plate elements.

        Parameters
        ----------
        plateIds : list of int
            List of plate IDs.
        varLoadDirection : int
            Load direction: (= 3 to 6 for LocalZ, GlobalX, GlobalY, GlobalZ, respectively)
        varInterpolateDirection : int
            Interpolate along Global Axis(Int or Long), valid direction codes are 1, 2, 3 for Interpolate along Global X, Y, Z. No other direction is a valid input.
        varMinLoad : float
            Minimum Pressure load
        varMaxLoad : float
            Maximum Pressure load

        Returns
        -------
        bool
            True if successfully.
            False if an error occurred.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddElementHydrostaticPressure([1], 3, 1, 0.0, 10.0)
        """
        safe_PlateIdList = make_safe_array_long_input(plateIds)
        return self._load.AddElementHydrostaticPressure(
            safe_PlateIdList,
            varLoadDirection,
            varInterpolateDirection,
            varMinLoad,
            varMaxLoad,
        )

    def AddTemperatureLoad(
        self,
        elementIds: list,
        varTempAxialElong: float,
        varTempDiffTopAndBtm: float,
        varTemDiffSide: float,
    ):
        """
        Adds TEMPERATURE LOAD to beam or plate elements.

        Parameters
        ----------
        elementIds : list of int
            List of element IDs.
        varTempAxialElong : float
            Change in temperature.
        varTempDiffTopAndBtm : float
            Temperature difference from the top to the bottom of the element (for calculating bending).
        varTemDiffSide : float
            Temperature difference from side to side of the element (local Z axis).

        Returns
        -------
        bool
            True if the temperature load was added successfully.
            False if an error occurred.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddTemperatureLoad([1], 20.0, 30.0, 40.0)
        """
        safe_ElementIdList = make_safe_array_long_input(elementIds)
        retVal = self._load.AddTemperatureLoad(
            safe_ElementIdList, varTempAxialElong, varTempDiffTopAndBtm, varTemDiffSide
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def AddStrainLoad(self, elementIds: list, varAxialElong: float):
        """
        Adds STRAIN LOAD to beam or plate elements.

        Parameters
        ----------
        elementIds : list of int
            List of element IDs.
        varAxialElong : float
            Initial axial elongation (+)/ shrinkage (-) in member due to misfit, etc.

        Returns
        -------
        bool
            True if the strain load was added successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddStrainLoad([1], 0.001)
        """
        safe_ElementIdList = make_safe_array_long_input(elementIds)
        retVal = self._load.AddStrainLoad(safe_ElementIdList, varAxialElong)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def AddLoadAndFactorToCombination(
        self, loadCombNo: int, loadNo: int, factor: float
    ):
        """
        Adds a primary load case with specified multiplication factor to an existing load combination.

        Parameters
        ----------
        loadCombNo : int
            Load Combination Number.
        loadNo : int
            Load Case Reference ID.
        factor : float
            Multiplication factor for the specified primary load case.

        Returns
        -------
        bool
            True if the load case was added to the combination successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddLoadAndFactorToCombination(1, 2, 1.0)
        """
        retVal = self._load.AddLoadAndFactorToCombination(loadCombNo, loadNo, factor)
        return bool(retVal)

    def AddMemberFloorLoad(
        self,
        varPressure: float,
        varYMIN: float,
        varYMAX: float,
        varZMIN: float,
        varZMAX: float,
        varXMIN: float,
        varXMAX: float,
    ):
        """
        Automatically finds enclosed panels in the given boundary (specified using max and min X, Y, Z range inputs) and adds a FLOOR LOAD. Generated floor load is applied only in the Global X direction with YRANGE option.

        Parameters
        ----------
        varPressure : float
            Magnitude of the pressure or concentrate load on the element.
        varYMIN : float
            Y range from which the load start (in global coordinate).
        varYMAX : float
            Y range at which the load end (in global coordinate).
        varZMIN : float
            Z range from which the load start (in global coordinate).
        varZMAX : float
            Z range at which the load end (in global coordinate).
        varXMIN : float
            X range from which the load start (in global coordinate).
        varXMAX : float
            X range at which the load end (in global coordinate).

        Returns
        -------
        bool
            True if the floor load was added successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddMemberFloorLoad(5.0, 0.0, 10.0, 0.0, 10.0, 0.0, 10.0)
        """
        retVal = self._load.AddMemberFloorLoad(
            varPressure, varYMIN, varYMAX, varZMIN, varZMAX, varXMIN, varXMAX
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def AddMemberFloorLoadEx(
        self,
        rangeType: int,
        loadDirection: int,
        pressure: float,
        grpOrOneWay: int,
        yMIN: float,
        yMAX: float,
        zMIN: float,
        zMAX: float,
        xMIN: float,
        xMAX: float,
    ):
        """
        Automatically finds enclosed panels in the given boundary (specified using max and min of X/Y /Z range inputs and varRange) and if varRange is 3 adds member group FLOOR LOAD (specified by GrpOrOneWay input).
        Otherwise adds a FLOOR LOAD with pressure (dPressure) in the Global X/Y/Z direction (as specified by Direction input) with RANGE option.

        Parameters
        ----------
        rangeType : int
            Type of the Range :
                +-------+------------+
                | Value | Range Type |
                +=======+============+
                | 0     | X-RANGE    |
                +-------+------------+
                | 1     | Y-RANGE    |
                +-------+------------+
                | 2     | Z-RANGE    |
                +-------+------------+
                | 3     | Group Load |
                +-------+------------+

        loadDirection:int
            Load direction :
                +-------+-----------+
                | Value | Direction |
                +-------+-----------+
                | 0     | Global X  |
                +-------+-----------+
                | 1     | Global Y  |
                +-------+-----------+
                | 2     | Global Z  |
                +-------+-----------+

        pressure: float
            Magnitude of the pressure or concentrate load on the elemen
        grpOrOneWay: int
            One-Way Load (if it is either "" or "0") or corresponding group name to add Floor Group Load (if it contains Group string name).
            Notes:
            - Group name should be of FLOOR group type.
        yMIN: float
            Y range from which the load start(in global coordinate).
        yMAX: float
            Y range at which the load end(in global coordinate).
        zMIN: float
            Z range from which the load start(in global coordinate).
        zMAX: float
            Z range at which the load end(in global coordinate).
        xMIN: float
            X range from which the load start(in global coordinate).
        xMAX: float
            X range at which the load end(in global coordinate).

        Returns
        -------
        int
            Returns 1 OK.
            Returns 0 General error.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddMemberFloorLoadEx(0, 0, 5.0, 0, 0.0, 10.0, 0.0, 10.0, 0.0, 10.0)
        """
        return self._load.AddMemberFloorLoadEx(
            rangeType,
            loadDirection,
            pressure,
            grpOrOneWay,
            yMIN,
            yMAX,
            zMIN,
            zMAX,
            xMIN,
            xMAX,
        )

    def AddElementTrapPressureEx(
        self,
        PlateIDs: list,
        LoadDirection: int,
        LoadVaryDirection: int,
        StartPressure: float,
        EndPressure: float,
        Pressure3: float,
        Pressure4: float,
    ):
        """
        Adds trapezoidal pressure loading to plate elements.

        Parameters
        ----------
        PlateNo : list of int
            List of Plate IDs.
        LoadDirection : int
            Load direction: (= 3 to 6 for LocalZ, GlobalX, GlobalY, GlobalZ, respectively)
        LoadVaryDirection : int
            Load varying direction: (= 1, 2 ,3 for X, Y and JOINT respectively)
        StartPressure : float
            Pressure at loading starting point.(Node1 when JOINT is selected)
        EndPressure : float
            Pressure at loading ending point.(Node2 when JOINT is selected)
        Pressure3 : float
            Pressure at loading point.(applicable only when JOINT is selected)
        Pressure4 : float
            Pressure at loading point.(applicable only when JOINT is selected)

        Returns
        -------
        bool
            True if OK.
            False if any General error.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddElementTrapPressureEx([1], 3, 1, 100.0, 200.0, 150.0, 250.0)
        """
        safe_PlateIDList = make_safe_array_long_input(PlateIDs)
        return self._load.AddElementTrapPressureEx(
            safe_PlateIDList,
            LoadDirection,
            LoadVaryDirection,
            StartPressure,
            EndPressure,
            Pressure3,
            Pressure4,
        )

    def AddWindDefinition(self, varTypeNo: int, varTypeName: str):
        """
        Adds a Wind Definition named "varTypeName" with number ID varTypeNo.

        Parameters
        ----------
        varTypeNo : int
            Wind Definition Type number ID.
        varTypeName : string
            String name of this new type.

        Returns
        -------
        bool
            True if the wind definition was added successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddWindDefinition(1, "Wind Load 1")
        """
        retVal = self._load.AddWindDefinition(varTypeNo, varTypeName)
        return bool(retVal)

    def AddWindIntensity(self, varTypeNo: int, varIntensity: list, varHeight: list):
        """
        Adds to Wind Definitions Wind Intensity by giving Intensity vs. Height.

        Parameters
        ----------
        varTypeNo : int
            Wind Definition Type number ID.
        varIntensity : list of float
            Intensity values float list
        varHeight : list of float
            Height value float list.

        Returns
        -------
        bool
            True if the wind intensity was added successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddWindIntensity(1,[5.2], [10.0])
        """
        safe_IntensityList = make_safe_array_double_input(varIntensity)
        safe_HeightList = make_safe_array_double_input(varHeight)
        intensity_array_vt = make_variant_vt_ref(
            safe_IntensityList, automation.VT_ARRAY | automation.VT_R8
        )
        height_array_vt = make_variant_vt_ref(
            safe_HeightList, automation.VT_ARRAY | automation.VT_R8
        )
        retVal = self._load.AddWindIntensity(
            varTypeNo, intensity_array_vt, height_array_vt
        )
        return bool(retVal)

    def AddWindExposure(
        self, varTypeNo: int, varExposureFactor: float, varNodeArray: list
    ):
        """
        Adds Wind Exposures factor to Wind Definitions and assign to nodes.

        Parameters
        ----------
        varTypeNo : int
            Wind Definition Type number ID.
        varExposureFactor : float
            Exposure factor.
        varNodeArray : list of int
            Node number ID list.

        Returns
        -------
        bool
            True if the wind exposure was added successfully.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddWindExposure(1, 1.0, [1,2,3])
        """
        safe_NodeList = make_safe_array_long_input(varNodeArray)
        node_array_vt = make_variant_vt_ref(
            safe_NodeList, automation.VT_ARRAY | automation.VT_I4
        )
        retVal = self._load.AddWindExposure(varTypeNo, varExposureFactor, node_array_vt)
        return bool(retVal)

    def AddWindLoad(
        self,
        varTypeNo: int,
        varDirection: int,
        dFraction: float,
        varOpenStructure: int,
        dYMIN: float,
        dYMAX: float,
        dZMIN: float,
        dZMAX: float,
        dXMIN: float,
        dXMAX: float,
    ):
        """
        Adds a wind load.

        Parameters
        ----------
        varTypeNo : int
            Wind Definition Type number ID.
        varDirection : int
            Wind load direction:
                +-------+-----------+
                | Value | Direction |
                +=======+===========+
                | 1     | Global X  |
                +-------+-----------+
                | 3     | Global Z  |
                +-------+-----------+
                | 4     | Global -X |
                +-------+-----------+
                | 6     | Global -Z |
                +-------+-----------+

        dFraction : float
            Factor to be used to multiply the wind loads. Negative signs may be used to indicate opposite direction of resulting load (default=1.0).
        varOpenStructure : int
            For Open-type of structure enter 1  , closed-type of structure 0
        dYMIN : float
            Ymin of GLOBAL Y range in which Wind load applied (assume Y axis is vertical).
        dYMAX : float
            Ymax of GLOBAL Y range in which Wind load applied (assume Y axis is vertical).
        dZMIN : float
            Zmin of GLOBAL Z range in which Wind load applied (assume Y axis is vertical).
        dZMAX : float
            Zmax of GLOBAL Z range in which Wind load applied (assume Y axis is vertical).
        dXMIN : float
            Xmin of GLOBAL X range in which Wind load applied (assume Y axis is vertical).
        dXMAX : float
            Xmax of GLOBAL X range in which Wind load applied (assume Y axis is vertical).

        Returns
        -------
        bool
            True if successful, False otherwise.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddWindLoad(1, 3, 1.0, 1, 0.0, 10.0, 0.0, 10.0, 0.0, 10.0)
        """
        retVal = self._load.AddWindLoad(
            varTypeNo,
            varDirection,
            dFraction,
            varOpenStructure,
            dYMIN,
            dYMAX,
            dZMIN,
            dZMAX,
            dXMIN,
            dXMAX,
        )
        return bool(retVal)

    def AddSeismicDefinition(self, varType: int, varAccidental: int):
        """
        Adds a Seismic Definition with default parameters.

        Parameters
        ----------
        varType : int
            Type of seismic code:
                +-------+------------------------------+
                | Value | Seismic Code                 |
                +=======+==============================+
                | 0     | UBC 1985                     |
                +-------+------------------------------+
                | 1     | UBC 1994                     |
                +-------+------------------------------+
                | 2     | UBC 1997                     |
                +-------+------------------------------+
                | 3     | Indian: IS 1893-1984         |
                +-------+------------------------------+
                | 4     | Indian: IS 1893-2002/2005    |
                +-------+------------------------------+
                | 5     | IBC 2000                     |
                +-------+------------------------------+
                | 6     | IBC 2003                     |
                +-------+------------------------------+
                | 7     | COLOMBIAN: NSR 98            |
                +-------+------------------------------+
                | 8     | JAPANESE (AIJ)               |
                +-------+------------------------------+
                | 9     | ALGERIAN: RPA                |
                +-------+------------------------------+
                | 10    | MEX: CFE-1993                |
                +-------+------------------------------+
                | 11    | MEX: NTC-1987                |
                +-------+------------------------------+
                | 12    | Indian: IS 1893-2016         |
                +-------+------------------------------+
                | 13    | Indian: IS 1893(Part4) 2015  |
                +-------+------------------------------+
                | 14    | IBC 2006                     |
                +-------+------------------------------+
                | 15    | IBC 2012                     |
                +-------+------------------------------+
                | 16    | IBC 2015                     |
                +-------+------------------------------+
                | 17    | IBC 2018                     |
                +-------+------------------------------+
                | 18    | CANADIAN: NRC-2005           |
                +-------+------------------------------+
                | 19    | CANADIAN: NRC-2010           |
                +-------+------------------------------+
                | 20    | CANADIAN: NRC-1995           |
                +-------+------------------------------+
                | 21    | COLOMBIAN: NSR 2010          |
                +-------+------------------------------+
                | 22    | Chinese: GB50011-2001        |
                +-------+------------------------------+
                | 23    | Chinese: GB50011-2010        |
                +-------+------------------------------+
                | 24    | TURKISH                      |
                +-------+------------------------------+

        varAccidental : int
            '1' to consider accidental torsion else '0' ignore

        Returns
        -------
        bool
             True if successful.
             False if unsucessful

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddSeismicDefinition(1, 0)
        """
        return self._load.AddSeismicDefinition(varType, varAccidental)

    def AddSeismicDefSelfWeight(self, varWeightFactor: float):
        """
        Adds self weight to Seismic Definition.

        Parameters
        ----------
        varWeightFactor : float
            Weight Factor to add to self weight

        Returns
        -------
        bool
             True if successful.
             False if unsucessful

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddSeismicDefSelfWeight(1.0)
        """
        return self._load.AddSeismicDefMemberWeight(varWeightFactor)

    def AddSeismicDefMemberWeight(
        self,
        varSeismicType: int,
        loadType: int,
        weight: float,
        startDist: float,
        endDist: float,
        memberList: list,
    ):
        """
        Adds member concentrated/uniform weight to Seismic Definition.

        Parameters
        ----------
        varSeismicType : int
            Type of seismic code:
                +-------+-----------------------------+
                | Value | Seismic Code                |
                +=======+=============================+
                | 0     | AUTO DETECT                 |
                +-------+-----------------------------+
                | 1     | ALGERIAN: RPA               |
                +-------+-----------------------------+
                | 2     | CANADIAN: NRC-1995          |
                +-------+-----------------------------+
                | 3     | CANADIAN: NRC-2005          |
                +-------+-----------------------------+
                | 4     | CANADIAN: NRC-2010          |
                +-------+-----------------------------+
                | 5     | CANADIAN: NRC-2020          |
                +-------+-----------------------------+
                | 6     | Chinese: GB50011-2001       |
                +-------+-----------------------------+
                | 7     | Chinese: GB50011-2010       |
                +-------+-----------------------------+
                | 8     | COLOMBIAN: NSR 95           |
                +-------+-----------------------------+
                | 9     | COLOMBIAN: NSR 2010         |
                +-------+-----------------------------+
                | 10    | IBC 2000                    |
                +-------+-----------------------------+
                | 11    | IBC 2003 ASCE 7-02          |
                +-------+-----------------------------+
                | 12    | IBC 2006/2009 ASCE 7-05     |
                +-------+-----------------------------+
                | 13    | IBC 2012 ASCE 7-10          |
                +-------+-----------------------------+
                | 14    | IBC 2015 ASCE 7-10          |
                +-------+-----------------------------+
                | 15    | IBC 2018 ASCE 7-16          |
                +-------+-----------------------------+
                | 16    | Indian: IS 1893-1984        |
                +-------+-----------------------------+
                | 17    | Indian: IS 1893-2002/2005   |
                +-------+-----------------------------+
                | 18    | Indian: IS 1893-2016        |
                +-------+-----------------------------+
                | 19    | Indian: IS 1893(Part4) 2015 |
                +-------+-----------------------------+
                | 20    | JAPANESE (AIJ)              |
                +-------+-----------------------------+
                | 21    | MEX: CFE-1993               |
                +-------+-----------------------------+
                | 22    | MEX: NTC-1987               |
                +-------+-----------------------------+
                | 23    | TURKISH                     |
                +-------+-----------------------------+
                | 24    | UBC 1985                    |
                +-------+-----------------------------+
                | 25    | UBC 1994                    |
                +-------+-----------------------------+
                | 26    | UBC 1997                    |
                +-------+-----------------------------+

        loadType : int
            1 for uniform loadType and 2 for concentrated loadType
        weight : float
            Uniform weight.
        startDist : float
            Starting distance( = distance from member starting node to weight starting point)
        endDist : float
             Ending distance( =  distance from member starting node to weight ending point)
        memberList : list of int
            List of Member ID to add member concentrated/uniform weight

        Returns
        -------
        bool
             True if successful adds member concentrated/uniform weight to Seismic Definition.
             False if unsucessful

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddSeismicDefMemberWeight(1, 2, 3.0, 4.0, 5.0, [6, 7, 8])
        """
        safe_MemberIdList = make_safe_array_long_input(memberList)
        return self._load.AddSeismicDefMemberWeight(
            varSeismicType, loadType, weight, startDist, endDist, safe_MemberIdList
        )

    def AddSeismicDefJointWeight(self, weight: float, nodeList: list):
        """
         Adds joint self weight to Seismic Definition.

        Parameters
        ----------
        weight : float
            Weight value.
        nodeList : list of int
            List of Node number IDs

        Returns
        -------
        bool
            True if successful, False otherwise.

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddSeismicDefJointWeight(1.0, [1, 2, 3])
        """
        safe_NodeIdList = make_safe_array_long_input(nodeList)
        retVal = self._load.AddSeismicDefJointWeight(weight, safe_NodeIdList)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def AddSeismicDefElementWeight(self, pressure: float, elementList: list):
        """
        Adds a pressure to Seismic Definition.

        Parameters
        ----------
        pressure : float
            Pressure Value
        elementList : List of int
            'List of element ID list.

        Returns
        -------
        int
             True if successful.
             False if unsucessful

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddSeismicDefElementWeight(1.0, [1, 2, 3])
        """
        return self._load.AddSeismicDefElementWeight(pressure, elementList)

    def AddSeismicDefFloorWeight(
        self,
        rangeType: int,
        loadDirection: int,
        pressure: float,
        grpOrOneWay: int,
        yMIN: float,
        yMAX: float,
        zMIN: float,
        zMAX: float,
        xMIN: float,
        xMAX: float,
    ):
        """
        Adds a floor weight to Seismic Definition.

        Parameters
        ----------
        rangeType : int
            Type of the Range :
                +-------+------------+
                | Value | Range Type |
                +=======+============+
                | 0     | X-RANGE    |
                +-------+------------+
                | 1     | Y-RANGE    |
                +-------+------------+
                | 2     | Z-RANGE    |
                +-------+------------+
                | 3     | Group Load |
                +-------+------------+

        loadDirection:int
            Load direction :
                +-------+-----------+
                | Value | Direction |
                +-------+-----------+
                | 0     | Global X  |
                +-------+-----------+
                | 1     | Global Y  |
                +-------+-----------+
                | 2     | Global Z  |
                +-------+-----------+

        pressure: float
            Magnitude of the pressure or concentrate load on the elemen
        grpOrOneWay: int
            One-Way Load (if it is either "" or "0") or corresponding group name to add Floor Group Load (if it contains Group string name).
            Notes:
            - Group name should be of FLOOR group type.
        yMIN: float
            Y range from which the load start(in global coordinate).
        yMAX: float
            Y range at which the load end(in global coordinate).
        zMIN: float
            Z range from which the load start(in global coordinate).
        zMAX: float
            Z range at which the load end(in global coordinate).
        xMIN: float
            X range from which the load start(in global coordinate).
        xMAX: float
            X range at which the load end(in global coordinate).

        Returns
        -------
        int
             True if successful.
             False if unsucessful

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddSeismicDefFloorWeight(0, 1, 2.0, 3, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0)
        """
        return self._load.AddSeismicDefFloorWeight(
            rangeType,
            loadDirection,
            pressure,
            grpOrOneWay,
            yMIN,
            yMAX,
            zMIN,
            zMAX,
            xMIN,
            xMAX,
        )

    def AddSeismicLoad(self, loadDirection: int, factor: float):
        """
        Adds a Seismic Definition with default parameters.

        Parameters
        ----------
        loadDirection : int
            Load direction: (= 0 to 2 for global X, Y and Z, respectively).
        factor : float
            'Multiplication factor to be used to multiply the seismic load.

        Returns
        -------
        bool
            True if successful.

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddSeismicLoad(0, 1.0)
        """
        retVal = self._load.AddSeismicLoad(loadDirection, factor)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def AddAutoLoadCombinations(
        self, loadCombCode: str, loadCombCategory: str, loadList: list
    ):
        """
        Automatically adds load combination based on assigned design code and Category.

        Parameters
        ----------
        loadCombCode : str
            Load Combination Code string name (refer to "Codes.ini")
        loadCombCategory : str
            Load Combination Category string name (refer to corresponding rule ini file defined in "Codes.ini")
        loadList : list of int
            Load case reference ID(s), Array of Load case numbers. If the array is either null or empty then all load cases in current model will be considered

        Returns
        -------
        int
            Returns load case reference ID with which automatically load combination generation starts. If nStartLoadCaseNo is valid, auto load combinations will be created from the provided ID.
            IfnStartLoadCaseNo is invalid Load Case ID already present Load Case ID, load combinations would automatically generated from next available Load Case ID and nStartLoadCaseNo will be returned/updatedwith this ID.

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddAutoLoadCombinations("AISC 9th Ed","2.3 LRFD General", [1, 2, 3])
        """
        load_list_vt = make_safe_array_long_input(loadList)
        start_load_case = create_variant_int(0)
        start_load_case_vt = make_variant_vt_ref(start_load_case, automation.VT_I4)
        result = self._load.AddAutoLoadCombinations(
            loadCombCode, loadCombCategory, load_list_vt, start_load_case_vt
        )
        if result < 0:
            raise_os_error_if_error_code(result)
        return start_load_case_vt[0]

    def AddRepeatLoad(self, varLoadCaseList: list, varFactorList: list):
        """
        Creates a primary load case using combinations of previously defined primary load cases.

        Parameters
        ----------
        varLoadCaseList : list of int
            (Primary) load case reference number ID(s) array.
        varFactorList : list of float
            Multiplication factor array.

        Returns
        -------
        int
            Returns 1 if Load Case is added successfully, 0 otherwise.

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddRepeatLoad([1, 2, 3], [1.0, 2.0, 3.0])
        """
        loadCaseIdList_safe_list = make_safe_array_long_input(varLoadCaseList)
        multiplicationList_safe_list = make_safe_array_double_input(varFactorList)
        return self._load.AddRepeatLoad(
            loadCaseIdList_safe_list, multiplicationList_safe_list
        )

    def AddLoadCasesToEnvelop(self, varEnvNo: int, varLoadCaseList: list):
        """
        Adds a list of primary load case(s) to an existed load envelop.

        Parameters
        ----------
        varEnvNo : int
            Load Envelop reference ID
        varLoadCaseList :list of int
            Load cases reference IDs list.

        Returns
        -------
        int
             Returns 1 if OK.
             Returns 0 if general error.

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddLoadCasesToEnvelop(1, [1, 2, 3])
        """
        safe_LoadCaseList = make_safe_array_long_input(varLoadCaseList)
        return self._load.AddLoadCasesToEnvelop(varEnvNo, safe_LoadCaseList)

    def AddReferenceLoad(
        self, varRefLoadCaseNoIds: list[int], varFactorList: list[float]
    ):
        """
        Adds a reference load item to the currently active load case.

        Parameters
        ----------
        varRefLoadCaseNoIds : list of int
            List of reference load case number IDs from Reference Load Definitions.
        varFactorList : list of float
            List of corresponding load factors.

        Returns
        -------
        int
            Reference load case number ID.

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddReferenceLoad([1, 2, 3], [1.0, 2.0, 3.0])
        """
        refLoadCaseNoId_safe_list = make_safe_array_long_input(varRefLoadCaseNoIds)
        ref_factors_safe_list = make_safe_array_double_input(varFactorList)

        retval = self._load.AddReferenceLoad(
            refLoadCaseNoId_safe_list, ref_factors_safe_list
        )
        if retval < 0:
            raise_os_error_if_error_code(retval)
        return retval

    def AddSeismicDefWallArea(self, nTypeNo: int, direction: str, sizeArray: list):
        """
        Adds wall area to Seismic Definition.
        Note:
            - Wall Area is only available in IS1893-2016 seismic code.

        Parameters
        ----------
        nTypeNo : int
            Type of seismic code:
                +-------+---------------------------------------+
                | Value | Seismic Code                          |
                +=======+=======================================+
                | 15    | Indian: IS 1893-2016                  |
                +-------+---------------------------------------+
        direction : string
            Direction value. [X directin or Z direction]
        sizeArray : List of float
            Length and Width list consisting of consecutive length and width measurements

        Returns
        -------
        bool
            True if successful, False otherwise.

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddSeismicDefWallArea(15, "X", [10.0, 20.0])
        """
        ref_size_safe_list = make_safe_array_double_input(sizeArray)
        refSizeArray_vt = make_variant_vt_ref(
            ref_size_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        retVal = self._load.AddSeismicDefWallArea(nTypeNo, direction, refSizeArray_vt)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def AddWindDefinitionASCE7Parameters(
        self,
        varTypeNo: int,
        code: int,
        windSpeed: float,
        heightAboveSeaLvl: float,
        bldgclass: int,
        bldgtype: int,
        expCat: int,
        varEscarpment: bool,
        wallType: int,
        varIsFlexible: bool,
        varEscarpmentData: list,
        varbldgData: list,
        varUnitsData: list,
        varFactorsUserInput: list,
        varFactors: list,
    ):
        """
        Generates the Wind Definition Parameters using ASCE CODE.

        Parameters
        ----------
        varTypeNo : int
            Wind Definition Type number ID
        code : int
            +-------+------------+
            | Value | ASCE CODE  |
            +=======+============+
            | 0     | ACSE 7-1995|
            +-------+------------+
            | 1     | ACSE 7-2002|
            +-------+------------+
            | 2     | ACSE 7-2010|
            +-------+------------+
            | 3     | ACSE 7-2016|
            +-------+------------+

        windSpeed : float
            Wind speed.
        heightAboveSeaLvl : float
            Ground Height above sea level. [Required only for ASCE7-2016. For other versions give 0.0]
        bldgclass : int
            +-------+----------------------------------+
            | value | Building Classification Category |
            +=======+==================================+
            | 0     |Category I                        |
            +-------+----------------------------------+
            | 1     |Category II                       |
            +-------+----------------------------------+
            | 2     |Category III                      |
            +-------+----------------------------------+
            | 3     |Category IV                       |
            +-------+----------------------------------+

        bldgtype : int
            +-------+--------------------------------------+
            | Value | Building Type                        |
            +=======+======================================+
            | 0     | Building Structures                  |
            +-------+--------------------------------------+
            | 1     | Chimney, Tank and similar structures |
            +-------+--------------------------------------+
            | 2     | Solid Signs                          |
            +-------+--------------------------------------+
            | 3     | Open Signs                           |
            +-------+--------------------------------------+
            | 4     | Lattice Framework                    |
            +-------+--------------------------------------+
            | 5     | Trussed Tower                        |
            +-------+--------------------------------------+

        expCat : int
            +-------+-------------------+
            | Value | Exposure Category |
            +=======+===================+
            | 0     | Exposure A        |
            +-------+-------------------+
            | 1     | Exposure B        |
            +-------+-------------------+
            | 2     | Exposure C        |
            +-------+-------------------+
            | 3     | Exposure D        |
            +-------+-------------------+

        varEscarpment : bool
             Consider Wind Speed-up over Hills (FALSE) or Escarpment (TRUE).
        wallType : int
            +-------+-----------+
            | Value | Wall Type |
            +=======+===========+
            | 0     | WindWard  |
            +-------+-----------+
            | 1     | Leeward   |
            +-------+-----------+
            | 2     | SideWall  |
            +-------+-----------+

        varIsFlexible : bool
             Consider structure is Flexible (TRUE) or RIGID (FALSE).

        varEscarpmentData : List of float
            Float list of size 4 containing information describing Hills or Escarpment
            +-------+----------------------------------------------------------+
            | Index | Data                                                     |
            +=======+==========================================================+
            | 0     | Type: 2D Ridge (0)/ 2D Escarpment (1)/ 3D Escarpment (2) |
            +-------+----------------------------------------------------------+
            | 1     | Height (H)                                               |
            +-------+----------------------------------------------------------+
            | 2     | Distance upwind of crest (Lh)                            |
            +-------+----------------------------------------------------------+
            | 3     | Distance from the crest to the building (x)              |
            +-------+----------------------------------------------------------+

        varbldgData : List of float
            Float list of size 7 containing information describing the building based on structure type :
                - Building Data :
                    +-------+------------------------------------------------------------------------------------------+
                    | Index | Item                                                                                     |
                    +=======+==========================================================================================+
                    | 0     | Enclosure Classification :                                                               |
                    |       |   - Before 2016 :                                                                        |
                    |       |     Open Building (0)/ Partially Enclosed (1)/ Enclosed Building (2)                     |
                    |       |   - [2016] :                                                                             |
                    |       |     Open Building (0)/ Partially Open (1)/ Partially Enclosed (2)/ Enclosed Building (3) |
                    +-------+------------------------------------------------------------------------------------------+
                    | 1     | Building Height                                                                          |
                    +-------+------------------------------------------------------------------------------------------+
                    | 2     | Building length long the direction of Wind (L)                                           |
                    +-------+------------------------------------------------------------------------------------------+
                    | 3     | Building length normal to the direction of Wind (B)                                      |
                    +-------+------------------------------------------------------------------------------------------+
                    | 4     | Building Natural Frequency                                                               |
                    +-------+------------------------------------------------------------------------------------------+
                    | 5     | Building Damping Ratio                                                                   |
                    +-------+------------------------------------------------------------------------------------------+

                - Tank Data :
                    +-------+----------------------------------------------------------------------------------------+
                    | Index | Item                                                                                   |
                    +=======+========================================================================================+
                    | 0     | Horizontal Cross-section Type :                                                        |
                    |       |   - Before 2016 :                                                                      |
                    |       |     Square (0)/ Square Diagonal (1)/ Hexagonal or Octagonal (2)/ Round (3)             |
                    |       |   - [2016] :                                                                           |
                    |       |     Square (0)/ Square Diagonal (1)/ Hexagonal (2)/ Octagonal Non-axisymmetric (3)     |
                    |       |     / Octagonal Axisymmetric (4) / Round Non-axisymmetric (5) / Round Axisymmetric (6) |
                    +-------+----------------------------------------------------------------------------------------+
                    | 1     | Tank Height                                                                            |
                    +-------+----------------------------------------------------------------------------------------+
                    | 2     | Least Horizontal Dimension (W)                                                         |
                    +-------+----------------------------------------------------------------------------------------+
                    | 3     | Depth of producing elements like Spoilers and Ribs (D')                                |
                    +-------+----------------------------------------------------------------------------------------+
                    | 4     | Structure Natural Frequency                                                            |
                    +-------+----------------------------------------------------------------------------------------+
                    | 5     | Structure Damping Ratio                                                                |
                    +-------+----------------------------------------------------------------------------------------+

                - Solid Sign Data :
                    +-------+----------------------------------------------------------------------------------------+
                    | Index | Item                                                                                   |
                    +=======+========================================================================================+
                    | 0     | Solid Sign Height (H)                                                                  |
                    +-------+----------------------------------------------------------------------------------------+
                    | 1     | Solid Sign M Dimension (M)                                                             |
                    +-------+----------------------------------------------------------------------------------------+
                    | 2     | Solid Sign N Dimension (N)                                                             |
                    +-------+----------------------------------------------------------------------------------------+
                    | 3     | Structure Natural Frequency                                                            |
                    +-------+----------------------------------------------------------------------------------------+
                    | 4     | Structure Damping Ratio                                                                |
                    +-------+----------------------------------------------------------------------------------------+

                - Open Sign/Lattice Framework Data :
                    +-------+----------------------------------------------------------------------------------------+
                    | Index | Item                                                                                   |
                    +=======+========================================================================================+
                    | 0     | Orientation Type: Flat (0)/ Rounded (1)                                                |
                    +-------+----------------------------------------------------------------------------------------+
                    | 1     | Height (H)                                                                             |
                    +-------+----------------------------------------------------------------------------------------+
                    | 2     | Width                                                                                  |
                    +-------+----------------------------------------------------------------------------------------+
                    | 3     | Diameter of typical round member                                                       |
                    +-------+----------------------------------------------------------------------------------------+
                    | 4     | Ratio of Solid Area to Gross Area                                                      |
                    +-------+----------------------------------------------------------------------------------------+
                    | 5     | Structure Natural Frequency                                                            |
                    +-------+----------------------------------------------------------------------------------------+
                    | 6     | Structure Damping Ratio                                                                |
                    +-------+----------------------------------------------------------------------------------------+

                - Trussed Tower Data :
                    +-------+----------------------------------------------------------------------------------------+
                    | Index | Item                                                                                   |
                    +=======+========================================================================================+
                    | 0     | Horizontal Cross Sectio Type: Triangle (0)/ Square (1)                                 |
                    +-------+----------------------------------------------------------------------------------------+
                    | 1     | Tank Height (H)                                                                        |
                    +-------+----------------------------------------------------------------------------------------+
                    | 2     | Width                                                                                  |
                    +-------+----------------------------------------------------------------------------------------+
                    | 3     | Ratio of Solid Area to Gross Area(in percetage)                                        |
                    +-------+----------------------------------------------------------------------------------------+
                    | 4     | Structure Natural Frequency                                                            |
                    +-------+----------------------------------------------------------------------------------------+
                    | 5     | Structure Damping Ratio                                                                |
                    +-------+----------------------------------------------------------------------------------------+

        varUnitsData : List of int
            Float list of size 7 containing Units of data inputs
            +-------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
            | Index | Data                                                                                                                                                                                                                                                                                                                                                                                                                                    |
            +=======+=========================================================================================================================================================================================================================================================================================================================================================================================================================================+
            | 0     | Unit of Wind Speed {mph(VelocityUnit::mph or 0) or m/sec(VelocityUnit::metersec or 1) or cm/sec(VelocityUnit::cmsec or 2) or mm/sec(VelocityUnit::mmsec or 3) or kmph(VelocityUnit::kmph or 4) or in/sec(VelocityUnit::inchsec or 5) or ft/sec(VelocityUnit::ftsec or 6) or Yd/sec(VelocityUnit::yardsec or 7)}                                                                                                                         |
            +-------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
            | 1     | Unit of Height above sea level {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}[Required only for ASCE7-2016. For other versions give any length unit.]                                            |
            +-------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
            | 2     | [Escarpment] Unit of Height (H) {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                                                   |
            +-------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
            | 3     | [Escarpment] Unit of Distance upwind of crest (Lh) {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                                |
            +-------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
            | 4     | [Escarpment] Unit of Distance from the crest to the building (x) {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                  |
            +-------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
            | 5     | [Building]Unit of Height/ [Tank]Unit of Height/ [Solid Sign]Unit of Height/ [Open Sign/Lattice]Unit of Height/ [Trusses Tower]Unit of Height {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}      |
            +-------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
            | 6     | [Building]Unit of Length/ [Tank]Unit of Width/ [Solid Sign]Unit of M dimension/ [Open Sign/Lattice]Unit of Width/ [Trusses Tower]Unit of Width {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}    |
            +-------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
            | 7     | [Building]Unit of Width/ [Tank]Unit of Depth/ [Solid Sign]Unit of N dimension/ [Open Sign/Lattice]Unit of Diameter/ [Trusses Tower]Not Applicable {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)} |
            +-------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+

        varFactorsUserInput : List of int
            Float list of size 7 containing information describing whether Factors are User Input or Calculated
            +-------+---------------------------------------------------------------------+
            | Index | Data                                                                |
            +=======+=====================================================================+
            | 0     | Kz is User Input(1) or Calculated(0)                                |
            +-------+---------------------------------------------------------------------+
            | 1     | Kzt is User Input(1) or Calculated(0)                               |
            +-------+---------------------------------------------------------------------+
            | 2     | I is User Input(1) or Calculated(0)                                 |
            +-------+---------------------------------------------------------------------+
            | 3     | Kd is User Input(1) or Calculated(0)                                |
            +-------+---------------------------------------------------------------------+
            | 4     | Ke is User Input(1) or Calculated(0) [Required only for ASCE7-2016] |
            +-------+---------------------------------------------------------------------+
            | 5     | G is User Input(1) or Calculated(0)                                 |
            +-------+---------------------------------------------------------------------+
            | 6     | Cp is User Input(1) or Calculated(0)                                |
            +-------+---------------------------------------------------------------------+
            | 7     | Gcpi is User Input(1) or Calculated(0)                              |
            +-------+---------------------------------------------------------------------+

        varFactors : List of int
            Float list of size 7 containing information describing whether Factors are User Input or Calculated
            +-------+------------------------------------------+
            | Index | Data                                     |
            +=======+==========================================+
            | 0     | Factor Kz                                |
            +-------+------------------------------------------+
            | 1     | Factor Kzt                               |
            +-------+------------------------------------------+
            | 2     | Factor I                                 |
            +-------+------------------------------------------+
            | 3     | Factor Kd                                |
            +-------+------------------------------------------+
            | 4     | Factor Ke [Required only for ASCE7-2016] |
            +-------+------------------------------------------+
            | 5     | Factor G                                 |
            +-------+------------------------------------------+
            | 6     | Factor Cp                                |
            +-------+------------------------------------------+
            | 7     | Factor Gcpi                              |
            +-------+------------------------------------------+

        Returns
        -------
        bool
            Returns True if succesful

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddWindDefinitionASCE7Parameters(1, "ASCE7-10", 100.0, 10.0, 1, 1, 1, 1, [1.0, 2.0, 3.0], [1.0, 2.0, 3.0])
        """
        EscarpmentData = make_safe_array_double_input(varEscarpmentData)

        bldgData_safe_list = make_safe_array_double_input(varbldgData)

        UnitsData_safe_list = make_safe_array_double_input(varUnitsData)

        FactorsUserInput_safe_list = make_safe_array_double_input(varFactorsUserInput)

        Factors_safe_list = make_safe_array_double_input(varFactors)

        retval = self._load.AddWindDefinitionASCE7Parameters(
            varTypeNo,
            code,
            windSpeed,
            heightAboveSeaLvl,
            bldgclass,
            bldgtype,
            expCat,
            varEscarpment,
            wallType,
            varIsFlexible,
            EscarpmentData,
            bldgData_safe_list,
            UnitsData_safe_list,
            FactorsUserInput_safe_list,
            Factors_safe_list,
        )

        return bool(retval)

    def AddNotionalLoad(
        self,
        varPrimaryLoadCaseList: list[int],
        varPLFactorList: list[float],
        varPLDirectionList: list[int],
        varReferenceLoadCaseList: list[int],
        varRLFactorList: list[float],
        varRLDirectionList: list[int],
    ):
        """
        Creates a Notional load case using combinations of previously defined primary load cases and Reference load cases.

        Parameters
        ----------
        varPrimaryLoadCaseList : list of int
            List of Primary load case reference number IDs
        varPLFactorList : list of int
            List of Multiplication factor of Primary load cases
        varPLDirectionList : list of int
            List of Directions of Primary load cases. Directions can be passed as following:
                +-----------------------+-----------------------+
                | Direction             | Integer               |
                +=======================+=======================+
                | X OR                  | 1                     |
                | GlobalLoadDirection X | 4                     |
                +-----------------------+-----------------------+
                | Y OR                  | 2                     |
                | GlobalLoadDirection Y | 5                     |
                +-----------------------+-----------------------+
                | Z OR                  | 3                     |
                | GlobalLoadDirection Z | 6                     |
                +-----------------------+-----------------------+

        varReferenceLoadCaseList : list of int
            List of Reference load case reference number IDs
        varRLFactorList : list of int
            List of Multiplication factor of Reference load cases
        varRLDirectionList : list of int
            List of Directions of Reference load cases. Directions can be passed as following:
                +-----------------------+-----------------------+
                | Direction             | Integer               |
                +=======================+=======================+
                | X OR                  | 1                     |
                | GlobalLoadDirection X | 4                     |
                +-----------------------+-----------------------+
                | Y OR                  | 2                     |
                | GlobalLoadDirection Y | 5                     |
                +-----------------------+-----------------------+
                | Z OR                  | 3                     |
                | GlobalLoadDirection Z | 6                     |
                +-----------------------+-----------------------+

        Returns
        -------
        bool
            return True if successful

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddNotionalLoad([1, 2], [1.0, 1.2], [1, 2], [3], [0.8], [3])
        """

        ref_PrimaryLoadCase_safe_list = make_safe_array_long_input(
            varPrimaryLoadCaseList
        )
        refPrimaryLoadCaseArray_vt = make_variant_vt_ref(
            ref_PrimaryLoadCase_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        ref_PLFactor_safe_list = make_safe_array_double_input(varPLFactorList)
        refPLFactorArray_vt = make_variant_vt_ref(
            ref_PLFactor_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        ref_PLDirection_safe_list = make_safe_array_long_input(varPLDirectionList)
        refPLDirectionArray_vt = make_variant_vt_ref(
            ref_PLDirection_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        ref_ReferenceLoadCase_safe_list = make_safe_array_long_input(
            varReferenceLoadCaseList
        )
        refReferenceLoadCaseArray_vt = make_variant_vt_ref(
            ref_ReferenceLoadCase_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        ref_RLFactor_safe_list = make_safe_array_double_input(varRLFactorList)
        refRLFactorArray_vt = make_variant_vt_ref(
            ref_RLFactor_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        ref_RLDirection_safe_list = make_safe_array_long_input(varRLDirectionList)
        refRLDirectionArray_vt = make_variant_vt_ref(
            ref_RLDirection_safe_list, automation.VT_ARRAY | automation.VT_I4
        )

        retval = self._load.AddNotionalLoad(
            refPrimaryLoadCaseArray_vt,
            refPLFactorArray_vt,
            refPLDirectionArray_vt,
            refReferenceLoadCaseArray_vt,
            refRLFactorArray_vt,
            refRLDirectionArray_vt,
        )
        if retval < 0:
            raise_os_error_if_error_code(retval)
        return True

    def AddDirectAnalysisDefinitionParameter(
        self, pParamType: int, members: list, param: float
    ):
        """
        Adds Direct Analysis Definition (FLEX,AXIAL parameters).

        Parameters
        ----------
        pParamType : int
            Integer indicating type of direct analysis parameter to be added. Integer value should be taken from following table :
                +--------------+-----------------------------------------+
                | Value        | AnalysisCommand                         |
                +==============+=========================================+
                | FLEX = 0     | DirectAnalysisParameterTypes.FLEX       |
                +--------------+-----------------------------------------+
                | AXIAL = 2    | DirectAnalysisParameterTypes.AXIAL      |
                +--------------+-----------------------------------------+

        members : list of int
            List of Member IDs
        param : float
            FLEX parameter value. [For AXIAL this value is Not Applicable. Should pass 0(zero)]

        Returns
        -------
        bool
            Returns TRUE if successful
            Returns FALSE if unsuccessful

        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddDirectAnalysisDefinitionParameter(0, [1, 2], 0.5)
        True
        """

        refMembersArray_vt = make_safe_array_long_input(members)

        result = self._load.AddDirectAnalysisDefinitionParameter(
            pParamType, refMembersArray_vt, param
        )
        return bool(result)

    def AddResponseSpectrumLoad(
        self,
        rsaCode: int,
        rsaCombination: int,
        varSet1Names: list,
        varSet1Vals: list,
        varSet2Names: list,
        varSet2Vals: list,
        varDataPairs: list,
    ):
        """
        Adds Response Spectrum load item to the currently active load case.

        Parameters
        ----------
        rsaCode : int
            Response Spectrum Loading Code. Refer to the following table for the integers corresponding to different codes.
        rsaCombination : int
            Modal combination rule. (SRSS = 0, ABS = 1, CQC = 2, ASCE = 3, TEN = 4, CSM = 5, GRP = 6)
        varSet1Names : list of string
            List of string containing parameter key words. Refer to the Technical Reference sections as indicated below.
        varSet1Vals : list of float
            List of Parameters values corresponding to the keywords supplied in varSet1Names array.
        varSet2Names : list of string
            List of string containing parameter key words for the spectrum generation data command. NULL can be used if not needed.
        varSet2Vals : list of float
            List of Parameters values corresponding to the keywords supplied in varSet1Names array. NULL can be used if not needed.
        varDataPairs : list of float
            List of containing pairs of time period and acceleration data. NULL can be used if not needed.
            Inputs (varSet2Names, varSet2Vals) and (varDataPairs) are mutually exclusive, i.e. if one is specified, other should not specified

        Notes:
        - Techincal Reference sections :
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | nTypeNo | Seismic Code        | Parameters                                                                                                  | Ref. Sec.     |
            +=========+=====================+=============================================================================================================+===============+
            | 0       | Generic or Custom   | DEC, ECC, X, Y, Z, ACC, DIS, SCA, DAM, CDA, MDA, LIN, LOG, MIS, ZPA, FF1, FF2, DOM, SIG, SAV, IMR, STA      | TR.32.10.1.1  |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 1       | IS:1893 Part 1 2002 | TOR, DEC, ECC, X, Y, Z, ACC, DIS, SCA, DAM, CDA, MDA, MIS, ZPA, IGN, DOM, SIG, SAV, IMR, STA                | TR.32.10.1.7  |
            |         |                     | SOI, CHE, RF                                                                                                |               |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 2       | IS:1893 2016        | TOR, DEC, ECC, X, Y, Z, ACC, DIS, SCA, DAM, CDA, MDA, LIN, LOG, MIS, ZPA, IGN, DOM, SIG, SAV, IMR, STA      | TR.32.10.1.8  |
            |         |                     | SOI, CHE, RF                                                                                                |               |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 4       | ENV 1998-1:1994     | ELA, DES, X, Y, Z, ACC, DAM, CDA, MDA, LIN, LOG, MIS, ZPA, DOM, SIG, SAV, IMR, STA                          | TR.32.10.1.4  |
            |         |                     | SOI, ALP, Q                                                                                                 |               |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 5       | EN 1998-1:2004      | ELA, DES, RS1, RS2, X, Y, Z, ACC, DAM, CDA, MDA, LIN, LOG, MIS, ZPA, DOM, SIG, SAV, IMR, STA                | TR.32.10.1.5  |
            |         |                     | SOI, ALP, Q                                                                                                 |               |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 6       | IBC 2006            | X, Y, Z, ACC, DAM, CDA, MDA, LIN, LOG, MISC, ZPA, DOM, SIG, SAV IMR, STA                                    | TR.32.10.1.10 |
            |         |                     | ZIP, LAT, LON, SS, S1, SCA, FA, FV, TL                                                                      |               |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 7       | IBC 2012            | X, Y, Z, ACC, DAM, CDA, MDA, LIN, LOG, MISC, ZPA, DOM, SIG, SAV IMR, STA                                    | TR.32.10.1.11 |
            |         |                     | ZIP, LAT, LON, SS, S1, SCA, FA, FV, TL                                                                      |               |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 8       | IBC 2015            | X, Y, Z, ACC, DAM, CDA, MDA, LIN, LOG, MISC, ZPA, DOM, SIG, SAV IMR, STA                                    | TR.32.10.1.12 |
            |         |                     | ZIP, LAT, LON, SS, S1, SCA, FA, FV, TL                                                                      |               |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 10      | SNiP II-7-81        | A, X, KWX, KX1, Y, KWY, KY1, Z, KWZ, KZ1, ACC, SCA, DAM, CDA, MDA, LIN, LOG, MIS, ZPA, DOM, SIG, SOI, SAV   | TR.32.10.1.14 |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 11      | SP 14.13330.2011    | ECC, A, X, Y, Z, ACC, SCA, DAM, LOG, MIS, ZPA, DOM, SIG, SOI                                                | TR.32.10.1.15 |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 12      | CANADIAN: NRC-2005  | TOR, DEC, ECC, X, Y, Z, ACC, DIS, SCA, DAM, CDA, MDA, LIN, LOG, MIS, ZPA, DOM, SIG, SAV, IMR, STA           | TR.32.10.1.2  |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 13      | CANADIAN: NRC-2010  | TOR, DEC, ECC, X, Y, Z, ACC, DIS, SCA, DAM, CDA, MDA, LIN, LOG, MIS, ZPA, DOM, SIG, SAV, IMR, STA           | TR.32.10.1.3  |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+
            | 14      | GB 50011 2010       | X, Y, Z, ALP, DAM, CDA, MDA, LIN, LOG, MISS, ZPA, DOM, SIG, INT, FRE, FOR,RAR, GRO, SCL                     | TR.32.10.1.6  |
            +---------+---------------------+-------------------------------------------------------------------------------------------------------------+---------------+

        - Following values should be specified for INT parameter of GB 50011 2010 code:
            +---------------+-------+
            | Fortification | Value |
            | Intensity     |       |
            +===============+=======+
            | 6             | 0     |
            +---------------+-------+
            | 7             | 1     |
            +---------------+-------+
            | 7A            | 2     |
            +---------------+-------+
            | 8             | 3     |
            +---------------+-------+
            | 8A            | 4     |
            +---------------+-------+
            | 9             | 5     |
            +---------------+-------+

        Returns
        -------
        bool
            Returns TRUE if successful
            Returns FALSE if unsuccessful

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.AddResponseSpectrumLoad(0, 0, ['DEC', 'ECC'], [1.0, 2.0], [], [], [])
        """

        Set1Names_safe_list = make_safe_array_string_input(varSet1Names)
        Set1Vals_safe_list = make_safe_array_double_input(varSet1Vals)
        Set2Names_safe_list = make_safe_array_string_input(varSet2Names)
        Set2Vals_safe_list = make_safe_array_double_input(varSet2Vals)
        DataPairs_safe_list = make_safe_array_double_input(varDataPairs)

        return self._load.AddResponseSpectrumLoadEx(
            rsaCode,
            rsaCombination,
            Set1Names_safe_list,
            Set1Vals_safe_list,
            Set2Names_safe_list,
            Set2Vals_safe_list,
            DataPairs_safe_list,
        )

    def AddAutoCombinationRepeat(
        self,
        varCode: str,
        varCategory: str,
        varLoadList: list,
        varStartLoadCaseNo: int,
        varGeneratedLCS: int,
        bVarReference: bool,
        bVarNotional: bool,
        dVarNotionalLoadFactor: float,
        bVarGB50017: bool,
        nVarFloor: int,
        bVarX: bool,
        bVarNegtiveX: bool,
        bVarZ: bool,
        bVarNegtiveZ: bool,
    ):
        """
        Automatically adds repeat load based on assigned design code and Category.

        Parameters
        ----------
        varCode : bool
            Load Combination Code string name (refer to "Codes.ini")
        varCategory : bool
            Load Combination Category string name (refer to corresponding rule ini file defined in "Codes.ini")
        varLoadList : List of int
            List of Load case reference IDs. If the array is either null or empty then all load cases in current model will be considered.
        varStartLoadCaseNo : int
            (Repeat Load) Load case reference ID with which automatically generation starts.
            If nStartLoadCaseNo is valid, auto repeat load will be created from the provided ID.
            If nStartLoadCaseNo is invalid Load Case ID/already present Load Case ID, the repeat load would automatically generated from next available Load Case ID and nStartLoadCaseNo will be returned/updated with this ID.
        varGeneratedLCS : int
            (Repeat Load) The counts of automatically generated repeat loads.
        bVarReference : bool
            Whether include Reference load
        bVarNotional : bool
            Whether include Notional load. If it's True but all Directions are Flase, return -1.
        dVarNotionalLoadFactor : float
            If bVarNotional is valid, the value of Notional load factor
        bVarGB50017 : bool
            Consider Notional load factor per GB 50017 Design code
        nVarFloor : int
            The count of floor, it is valid when bVarGB50017 is True only
        bVarX : bool
            Consider X Direction of Notional Load
        bVarNegtiveX : bool
            Consider -X Direction of Notional Load
        bVarZ : bool
            Consider Z Direction of Notional Load
        bVarNegtiveZ : bool
            Consider -Z Direction of Notional Load

        Returns
        -------
        bool
            True if successful, False otherwise.

        Notes:
            - The default path of Codes.ini under "%localappdata%\\Bentley\\Engineering\\STAAD.Pro <version>\\Default\\Language\\en".
        """
        loadList_safe_list = make_safe_array_long_input(varLoadList)
        retVal = self._load.AddAutoCombinationRepeat(
            varCode,
            varCategory,
            loadList_safe_list,
            varStartLoadCaseNo,
            varGeneratedLCS,
            bVarReference,
            bVarNotional,
            dVarNotionalLoadFactor,
            bVarGB50017,
            nVarFloor,
            bVarX,
            bVarNegtiveX,
            bVarZ,
            bVarNegtiveZ,
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def RemoveLoadCasesFromEnvelop(self, varEnvNo: int, varLoadCaseList: list):
        """
        Removes a list of primary load case(s) from an existed load envelop.

        Parameters
        ----------
        varEnvNo : int
            Load Envelop reference ID
        varLoadCaseList :list of int
            Load cases reference IDs list.

        Returns
        -------
        bool
            True if successful, False otherwise.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.RemoveLoadCasesFromEnvelop(1, [2, 3])
        """
        safe_LoadCaseList = make_safe_array_long_input(varLoadCaseList)
        retval = self._load.RemoveLoadCasesFromEnvelop(varEnvNo, safe_LoadCaseList)
        return bool(retval)

    def RemoveAttribute(self, lLoadCase: int):
        """
        Removes the load attribute specified by lLoadCase.

        Parameters
        ----------
        lLoadCase : int
            Load case reference ID

        Returns
        -------
        bool
            True if successful, False otherwise.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.RemoveAttribute(1)
        """
        retVal = self._load.RemoveAttribute(lLoadCase)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def ClearPrimaryLoadCase(self, varLoadCaseNos: list, isReferenceLoad: bool):
        """
        Clears the load items in a specified Primary Load cases or Reference Load cases.

        Parameters
        ----------
        varLoadCaseNos : list
            Primary load case reference ID(s) list.
        isReferenceLoad : bool
            If reference load case(s): True or False.

        Returns
        -------
        int
            Returns 1 if OK
            Returns 0 if failed to delete load(s)

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.ClearPrimaryLoadCase([1, 2, 3], False)
        """
        safe_LoadCaseList = make_safe_array_long_input(varLoadCaseNos)
        return self._load.ClearPrimaryLoadCase(safe_LoadCaseList, isReferenceLoad)

    def ClearReferenceLoadCase(self, varLoadCaseNos: list):
        """
        Clears the load items in a specified Primary Load cases or Reference Load cases.

        Parameters
        ----------
        varLoadCaseNos : list
            Primary load case reference ID(s) list.

        Returns
        -------
        int
            Returns 1 if OK
            Returns 0 if failed to delete load(s)

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.ClearReferenceLoadCase([1, 2, 3])
        """
        safe_LoadCaseList = make_safe_array_long_input(varLoadCaseNos)
        return self._load.ClearReferenceLoadCase(safe_LoadCaseList)

    def IsDynamicLoadIncluded(self, nLoadCase: int):
        """
        Checks if dynamic load included in specified load case.

        Parameters
        ----------
        nLoadCase : int
            Load case reference ID

        Returns
        -------
        bool
            True if dynamic load included in specified load case, False otherwise.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.IsDynamicLoadIncluded(1)
        """
        retVal = self._load.IsDynamicLoadIncluded(nLoadCase)
        return bool(retVal)

    def IsCombinationCase(self, nLoadCase: int):
        """
        Checks if specified load case is combination load case.

        Parameters
        ----------
        nLoadCase : int
            Load case reference ID

        Returns
        -------
        bool
            True if specified load case is combination load case, False otherwise.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.IsCombinationCase(1)
        """
        retVal = self._load.IsCombinationCase(nLoadCase)
        return bool(retVal)

    def SplitLoadsOnBeam(self, varBeamOld: int, varBeamNew: int):
        """
        Split Load from BeamOld to BeamNew.

        Parameters
        ----------
        varBeamOld : int
            Old Beam Id
        varBeamNew : int
            New Beam Id

        Returns
        -------
        bool
            True if Successful
            False if General Error.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.SplitLoadsOnBeam(1, 2)
        """
        return bool(self._load.SplitLoadsOnBeam(varBeamOld, varBeamNew))

    def MergeLoadsOnBeam(self, varBeamToKeep: int, varBeamToMerge: int):
        """
        Merge Load from beam to merge.

        Parameters
        ----------
        varBeamToKeep : int
            Beam Id where load to not merge.
        varBeamToMerge : int
            Beam Id to where load to merge.

        Returns
        -------
        bool
            True if Successful
            False if General Error.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.MergeLoadsOnBeam(1, 2)
        """
        return bool(self._load.MergeLoadsOnBeam(varBeamToKeep, varBeamToMerge))

    def BeginLoadMerging(self):
        """
        Begin Load Merging

        Returns
        -------
        None

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.BeginLoadMerging()
        """
        return self._load.BeginLoadMerging()

    def EndLoadMerging(self):
        """
        End Load Merging

        Returns
        -------
        None

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.EndLoadMerging()
        """
        return self._load.EndLoadMerging()

    def ModifySeismicDefinitionParams(self, varParamName: str, varValue: float):
        """
        Modifies or adds a seismic parameter in the existing seismic definition.

        Parameters
        ----------
        varParamName : string
            Parameter name for the corresponding code in the seismic definition.
        varValue : float
             Value corresponding to the above parameter:
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | Seismic Code                  | Parameters                                                                                                  |
                +===============================+=============================================================================================================+
                | ALGERIAN: RPA                 | A Q RX RZ STYPE CT CRDAMP PX PZ                                                                             |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | CANADIAN: NRC-1995            | V ZA ZV RX RZ I F CT PX PZ                                                                                  |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | CANADIAN: NRC-2005            | SA1 SA2 SA3 SA4 IE SCLASS MVX MVZ JX JZ RDX RDZ ROX ROZ CT PX PZ FA FV                                      |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | CANADIAN: NRC-2010            | SA1 SA2 SA3 SA4 I SCLASS MVX MVZ RDX RDZ ROX ROZ CTX CTZ PX PZ FA FV STX STZ MD                             |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | CHINESE: GB50011-2001         | INTENSITY FREQUENT RARE GROUP SCLASS DAMP DELN SF PX PZ GFACTOR                                             |
                |                               | Note: For CHINESE: GB50011-2001 FREQUENT/RARE parameter value will be 0, 1 respectively                     |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | CHINESE: GB50011-2010         | INTENSITY FREQUENT FORTIFIED RARE GROUP SCLASS DAMP GFACTOR DELN SF PX PZ                                   |
                |                               | Note: For CHINESE: GB50011-2010 FREQUENT/FORTIFIED/RARE parameter value will be 0, 1 and 2 respectively     |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | COLOMBIAN: NSR 98             | ZONE I S                                                                                                    |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | COLOMBIAN: NSR 2010           | AA AV FA FV I CT PX PZ ALPHA                                                                                |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | IBC 2000                      | SDS SD1 S1 I RX RZ SCLASS CT PX PZ                                                                          |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | IBC 2003                      | SDS SD1 S1 I RX RZ SCLASS CT PX PZ                                                                          |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | IBC 2006                      | SS S1 ZIP I RX RZ SCLASS CTX CTZ PX PZ LAT LONG TL FA FV XX XZ                                              |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | IBC 2012                      | SS S1 ZIP I RX RZ SCLASS CTX CTZ PX PZ LAT LONG TL FA FV XX XZ                                              |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | IBC 2015                      | SS S1 ZIP I RX RZ SCLASS CTX CTZ PX PZ LAT LONG TL FA FV XX XZ                                              |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | IBC 2018                      | SS S1 ZIP I RX RZ SCLASS CTX CTZ PX PZ LAT LONG TL FA FV XX XZ                                              |
                |                               | Note: For IBC 2006 - 2018 Please provide any one of ZIP OR LAT LONG OR SS S1                                |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | INDIAN: IS 1893-1984          | ZONE K I B PX PZ                                                                                            |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | INDIAN: IS 1893-2002/2005     | ZONE RF I SS ST DM PX PZ DT GL SA DF CS AX ES CV DV                                                         |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | INDIAN: IS 1893-2016          | ZONE RF I SS ST DM PX PZ DT GL SA DF HT DX DZ                                                               |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | INDIAN: IS 1893(Part4) 2015   | ZONE RF I SS ST DM PX PZ SA DF                                                                              |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | JAPANESE (AIJ)                | ZONE CO TC ALPHA                                                                                            |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | MEX: CFE-1993                 | ZONE QX QZ GROUP STYPE REGULAR TS PX PZ                                                                     |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | MEX: NTC-1987                 | ZONE QX QZ GROUP SHADOWED REGULAR REDUCE PX PZ                                                              |
                |                               | Note: For SHADOWED, REGULAR and REDUCE parameter value will be 0 or 1 respectively                          |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | TURKISH                       | A TA TB I RX RZ CT PX PZ                                                                                    |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | UBC 1985                      | ZONE I K TS                                                                                                 |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | UBC 1994                      | ZONE I RWX RWZ S CT PX PZ                                                                                   |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+
                | UBC 1997                      | ZONE I RWX RWZ STYPE CT PX PZ NA NV                                                                         |
                +-------------------------------+-------------------------------------------------------------------------------------------------------------+

        Returns
        -------
        bool
            True if OK

        Example
        -------
        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.ModifySeismicDefinitionParams('ZONE', 0.2)
        """
        retVal = self._load.ModifySeismicDefinitionParams(varParamName, varValue)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def ComputeWallWindPressureProfile(
        self,
        loadingCode: int,
        windSpeed: float,
        bldgClass: int,
        bldgType: int,
        expCat: int,
        bEscarpment: bool,
        varUnitsData: list,
        varEscarpmentData: list,
        varBldgData: list,
        wallType: int,
    ):
        """
        Generates the wall wind pressure profile using ASCE CODE.

        Parameters
        ----------
        loadingCode : int
            ASCE CODE:
                +------------------+-------------------+
                | Value            | ASCE CODE         |
                +==================+===================+
                | ASCE7Y95 = 0     | ACSE 7-1995       |
                +------------------+-------------------+
                | ACSE702 = 1      | ACSE 7-2002       |
                +------------------+-------------------+
                | ACSE705_10 = 2   | ACSE 7-2010       |
                +------------------+-------------------+

        windSpeed : float
            Wind speed. Default value 85 mph.
        bldgClass : int
            Building Classification Category:
                +--------------+------------------------------------+
                | Value        | Building Classification Category   |
                +==============+====================================+
                | TypeI = 0    | Category I                         |
                +--------------+------------------------------------+
                | TypeII = 1   | Category II                        |
                +--------------+------------------------------------+
                | TypeIII = 2  | Category III                       |
                +--------------+------------------------------------+
                | TypeIV = 3   | Category IV                        |
                +--------------+------------------------------------+

        bldgtype : int
            Structure Type:
                +------------------+--------------------------------------+
                | Value            | Structure Type                       |
                +==================+======================================+
                | Building = 0     | Building Structures                  |
                +------------------+--------------------------------------+
                | Chimney = 1      | Chimney, Tank and similar structures |
                +------------------+--------------------------------------+
                | Solidsign = 2    | Solid Signs                          |
                +------------------+--------------------------------------+
                | Opensign = 3     | Open Signs                           |
                +------------------+--------------------------------------+
                | Laticeframe = 4  | Lattice Framework                    |
                +------------------+--------------------------------------+
                | Trusstower = 5   | Trussed Tower                        |
                +------------------+--------------------------------------+

        expCat : int
            Exposure Category:
                +------------------+--------------------------------------+
                | Value            | Exposure Category                    |
                +==================+======================================+
                | ExpA = 0         | Exposure A                           |
                +------------------+--------------------------------------+
                | ExpB = 1         | Exposure B                           |
                +------------------+--------------------------------------+
                | ExpC = 2         | Exposure C                           |
                +------------------+--------------------------------------+
                | ExpD = 3         | Exposure D                           |
                +------------------+--------------------------------------+

        bEscarpment : bool
            Consider Wind Speed-up over Hills (FALSE) or Escarpment (TRUE).
        varUnitsData : list of long
            Integer list of size 8 containing Units of data inputs:
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | Index            | Data                                                                                                                                                                                                                                                                                                                                                                                                                                    |
                +==================+=========================================================================================================================================================================================================================================================================================================================================================================================================================================+
                | 0                | Unit of Wind Speed {mph(VelocityUnit::mph or 0) or m/sec(VelocityUnit::metersec or 1) or cm/sec(VelocityUnit::cmsec or 2) or mm/sec(VelocityUnit::mmsec or 3) or kmph(VelocityUnit::kmph or 4) or in/sec(VelocityUnit::inchsec or 5) or ft/sec(VelocityUnit::ftsec or 6) or Yd/sec(VelocityUnit::yardsec or 7)}                                                                                                                         |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 1                | Unit of Ground height above sea level {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                                             |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 2                | [Escarpment] Unit of Height (H) {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                                                   |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 3                | [Escarpment] Unit of Distance upwind of crest (Lh) {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                                |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 4                | [Escarpment] Unit of Distance from the crest to the building (x) {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                  |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 5                | [Building]Unit of Height/ [Tank]Unit of Height/ [Solid Sign]Unit of Height/ [Open Sign/Lattice]Unit of Height/ [Trusses Tower]Unit of Height {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}      |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 6                | [Building]Unit of Length/ [Tank]Unit of Width/ [Solid Sign]Unit of M dimension/ [Open Sign/Lattice]Unit of Width/ [Trusses Tower]Unit of Width {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}    |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 7                | [Building]Unit of Width/ [Tank]Unit of Depth/ [Solid Sign]Unit of N dimension/ [Open Sign/Lattice]Unit of Diameter/ [Trusses Tower]Not Applicable {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)} |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+

        varescarpmentData : list
            Information describing Hills or Escarpment:
                +-------+----------------------------------------------------------+
                | Index | Data                                                     |
                +=======+==========================================================+
                | 0     | Type: 2D Ridge (0), 2D Escarpment (1), 3D Escarpment (2) |
                +-------+----------------------------------------------------------+
                | 1     | Height (H)                                               |
                +-------+----------------------------------------------------------+
                | 2     | Distance upwind of crest (Lh)                            |
                +-------+----------------------------------------------------------+
                | 3     | Distance from the crest to the building (x)              |
                +-------+----------------------------------------------------------+

        varbldgData : list
            List of size 7 containing information describing the building based on structure type.
                - Building Data :
                    +-------+--------------------------------------------------------------------------------------------+
                    | Index | Item                                                                                       |
                    +=======+============================================================================================+
                    | 0     | Enclosure Classification: Open Building (0)/ Partially Enclosed (1)/ Enclosed Building (2) |
                    +-------+--------------------------------------------------------------------------------------------+
                    | 1     | Building Height                                                                            |
                    +-------+--------------------------------------------------------------------------------------------+
                    | 2     | Building length long the direction of Wind (L)                                             |
                    +-------+--------------------------------------------------------------------------------------------+
                    | 3     | Building length normal to the direction of Wind (B)                                        |
                    +-------+--------------------------------------------------------------------------------------------+
                    | 4     | Building Natural Frequency                                                                 |
                    +-------+--------------------------------------------------------------------------------------------+
                    | 5     | Building Damping Ratio                                                                     |
                    +-------+--------------------------------------------------------------------------------------------+

                - OR Tank Data :
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | Index | Item                                                                                                   |
                    +=======+========================================================================================================+
                    | 0     | Horizontal Cross-section Type:- Square (0)/ Square Diagonal (1)/ Hexagonal or Octagonal (2)/ Round (3) |
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | 1     | Tank Height (H)                                                                                        |
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | 2     | Least Horizontal Dimension (W)                                                                         |
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | 3     | Depth of producing elements like Spoilers and Ribs (D')                                                |
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | 4     | Structure Natural Frequency                                                                            |
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | 5     | Structure Damping Ratio                                                                                |
                    +-------+--------------------------------------------------------------------------------------------------------+

                - OR Solid Sign Data :
                    +-------+-----------------------------+
                    | Index | Item                        |
                    +=======+=============================+
                    | 0     | Solid Sign Height  (H)      |
                    +-------+-----------------------------+
                    | 1     | Solid Sign M Dimension (M)  |
                    +-------+-----------------------------+
                    | 2     | Solid Sign N Dimension (N)  |
                    +-------+-----------------------------+
                    | 3     | Structure Natural Frequency |
                    +-------+-----------------------------+
                    | 4     | Structure Damping Ratio     |
                    +-------+-----------------------------+

                - OR Open Sign/Lattice Framework Data :
                    +-------+-----------------------------------------+
                    | Index | Item                                    |
                    +=======+=========================================+
                    | 0     | Orientation Type: Flat (0)/ Rounded (1) |
                    +-------+-----------------------------------------+
                    | 1     | Height (H)                              |
                    +-------+-----------------------------------------+
                    | 2     | Width                                   |
                    +-------+-----------------------------------------+
                    | 3     | Diameter of typical round member        |
                    +-------+-----------------------------------------+
                    | 4     | Structure Natural Frequency             |
                    +-------+-----------------------------------------+
                    | 5     | Structure Damping Ratio                 |
                    +-------+-----------------------------------------+
                    | 6     | Ratio of Solid Area to Gross Area       |
                    +-------+-----------------------------------------+

                - OR Trussed Tower Data :
                    +-------+--------------------------------------------------------+
                    | Index | Item                                                   |
                    +=======+========================================================+
                    | 0     | Horizontal Cross Sectio Type: Triangle (0)/ Square (1) |
                    +-------+--------------------------------------------------------+
                    | 1     | Height (H)                                             |
                    +-------+--------------------------------------------------------+
                    | 2     | Width                                                  |
                    +-------+--------------------------------------------------------+
                    | 3     | Structure Natural Frequency                            |
                    +-------+--------------------------------------------------------+
                    | 4     | Structure Damping Ratio                                |
                    +-------+--------------------------------------------------------+
                    | 5     | Ratio of Solid Area to Gross Area(in percetage)        |
                    +-------+--------------------------------------------------------+

        wallType : int
            Building wall to generate Wind Load on:
                +------------------+--------------------------------------+
                | Value            | Wall Type                            |
                +==================+======================================+
                | WindWard = 0     | WindWard                             |
                +------------------+--------------------------------------+
                | LeeWard = 1      | Leeward                              |
                +------------------+--------------------------------------+
                | SideWall = 2     | SideWall                             |
                +------------------+--------------------------------------+

                - (0 to 2 for WindWard, Leeward and SideWall, respectively).

        Returns
        -------
        int
            Returns number of Height or Intensity data.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.ComputeWallWindPressureProfile(2, 90.0, 1, 0, 2, False, [0]*8, [0.0]*4, [0.0]*7, 0)
        """
        safe_UnitsDataList = make_safe_array_long_input(varUnitsData)
        safe_EscarpmentDataList = make_safe_array_double_input(varEscarpmentData)
        safe_BldgDataList = make_safe_array_double_input(varBldgData)
        retval = self._load.ComputeWallWindPressureProfile(
            loadingCode,
            windSpeed,
            bldgClass,
            bldgType,
            expCat,
            int(bEscarpment),
            safe_UnitsDataList,
            safe_EscarpmentDataList,
            safe_BldgDataList,
            wallType,
        )
        if retval < 0:
            raise_os_error_if_error_code(retval)
        return retval

    def ComputeWallWindPressureProfileASCE72016(
        self,
        windSpeed: float,
        heightAboveSeaLvl: float,
        bldgClass: int,
        bldgType: int,
        expCat: int,
        bEscarpment: bool,
        varUnitsData: list,
        varEscarpmentData: list,
        varBldgData: list,
        wallType: int,
    ):
        """
        Modifies or adds a seismic parameter in the existing seismic definition.

        Parameters
        ----------
        windSpeed : float
            Wind speed. Default value 85 mph.
        heightAboveSeaLvl : float
            Ground height above sea level. Used only for ASCE7-2016 Wind. Default value 0.0 ft.
        bldgClass : int
            Building Classification Category:
                +--------------+------------------------------------+
                | Value        | Building Classification Category   |
                +==============+====================================+
                | TypeI = 0    | Category I                         |
                +--------------+------------------------------------+
                | TypeII = 1   | Category II                        |
                +--------------+------------------------------------+
                | TypeIII = 2  | Category III                       |
                +--------------+------------------------------------+
                | TypeIV = 3   | Category IV                        |
                +--------------+------------------------------------+

        bldgtype : int
            Structure Type:
                +------------------+--------------------------------------+
                | Value            | Structure Type                       |
                +==================+======================================+
                | Building = 0     | Building Structures                  |
                +------------------+--------------------------------------+
                | Chimney = 1      | Chimney, Tank and similar structures |
                +------------------+--------------------------------------+
                | Solidsign = 2    | Solid Signs                          |
                +------------------+--------------------------------------+
                | Opensign = 3     | Open Signs                           |
                +------------------+--------------------------------------+
                | Laticeframe = 4  | Lattice Framework                    |
                +------------------+--------------------------------------+
                | Trusstower = 5   | Trussed Tower                        |
                +------------------+--------------------------------------+

        expCat : int
            Exposure Category:
                +------------------+--------------------------------------+
                | Value            | Exposure Category                    |
                +==================+======================================+
                | ExpA = 0         | Exposure A                           |
                +------------------+--------------------------------------+
                | ExpB = 1         | Exposure B                           |
                +------------------+--------------------------------------+
                | ExpC = 2         | Exposure C                           |
                +------------------+--------------------------------------+
                | ExpD = 3         | Exposure D                           |
                +------------------+--------------------------------------+

        bEscarpment : bool
            Consider Wind Speed-up over Hills (FALSE) or Escarpment (TRUE).
        varUnitsData : list of long
            Integer list of size 8 containing Units of data inputs:
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | Index            | Data                                                                                                                                                                                                                                                                                                                                                                                                                                    |
                +==================+=========================================================================================================================================================================================================================================================================================================================================================================================================================================+
                | 0                | Unit of Wind Speed {mph(VelocityUnit::mph or 0) or m/sec(VelocityUnit::metersec or 1) or cm/sec(VelocityUnit::cmsec or 2) or mm/sec(VelocityUnit::mmsec or 3) or kmph(VelocityUnit::kmph or 4) or in/sec(VelocityUnit::inchsec or 5) or ft/sec(VelocityUnit::ftsec or 6) or Yd/sec(VelocityUnit::yardsec or 7)}                                                                                                                         |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 1                | Unit of Ground height above sea level {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                                             |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 2                | [Escarpment] Unit of Height (H) {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                                                   |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 3                | [Escarpment] Unit of Distance upwind of crest (Lh) {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                                |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 4                | [Escarpment] Unit of Distance from the crest to the building (x) {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}                                                                                  |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 5                | [Building]Unit of Height/ [Tank]Unit of Height/ [Solid Sign]Unit of Height/ [Open Sign/Lattice]Unit of Height/ [Trusses Tower]Unit of Height {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}      |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 6                | [Building]Unit of Length/ [Tank]Unit of Width/ [Solid Sign]Unit of M dimension/ [Open Sign/Lattice]Unit of Width/ [Trusses Tower]Unit of Width {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)}    |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+
                | 7                | [Building]Unit of Width/ [Tank]Unit of Depth/ [Solid Sign]Unit of N dimension/ [Open Sign/Lattice]Unit of Diameter/ [Trusses Tower]Not Applicable {inch(LengthUnit::In or 0) or ft(LengthUnit::Ft or 1) or foot(LengthUnit::foot or 2) or cm(LengthUnit::cm or 3) or m(LengthUnit::M or 4) or mm(LengthUnit::Mm or 5) or dm(LengthUnit::Dm or 6) or Km(LengthUnit::Km or 7) or yard(LengthUnit::Yd or 8) or mile(LengthUnit::mil or 9)} |
                +------------------+-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------+

        varescarpmentData : list
            Information describing Hills or Escarpment:
                +-------+----------------------------------------------------------+
                | Index | Data                                                     |
                +=======+==========================================================+
                | 0     | Type: 2D Ridge (0), 2D Escarpment (1), 3D Escarpment (2) |
                +-------+----------------------------------------------------------+
                | 1     | Height (H)                                               |
                +-------+----------------------------------------------------------+
                | 2     | Distance upwind of crest (Lh)                            |
                +-------+----------------------------------------------------------+
                | 3     | Distance from the crest to the building (x)              |
                +-------+----------------------------------------------------------+

        varbldgData : list
            List of size 7 containing information describing the building based on structure type.
                - Building Data :
                    +-------+--------------------------------------------------------------------------------------------+
                    | Index | Item                                                                                       |
                    +=======+============================================================================================+
                    | 0     | Enclosure Classification: Open Building (0)/ Partially Enclosed (1)/ Enclosed Building (2) |
                    +-------+--------------------------------------------------------------------------------------------+
                    | 1     | Building Height                                                                            |
                    +-------+--------------------------------------------------------------------------------------------+
                    | 2     | Building length long the direction of Wind (L)                                             |
                    +-------+--------------------------------------------------------------------------------------------+
                    | 3     | Building length normal to the direction of Wind (B)                                        |
                    +-------+--------------------------------------------------------------------------------------------+
                    | 4     | Building Natural Frequency                                                                 |
                    +-------+--------------------------------------------------------------------------------------------+
                    | 5     | Building Damping Ratio                                                                     |
                    +-------+--------------------------------------------------------------------------------------------+

                - OR Tank Data :
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | Index | Item                                                                                                   |
                    +=======+========================================================================================================+
                    | 0     | Horizontal Cross-section Type:- Square (0)/ Square Diagonal (1)/ Hexagonal or Octagonal (2)/ Round (3) |
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | 1     | Tank Height (H)                                                                                        |
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | 2     | Least Horizontal Dimension (W)                                                                         |
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | 3     | Depth of producing elements like Spoilers and Ribs (D')                                                |
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | 4     | Structure Natural Frequency                                                                            |
                    +-------+--------------------------------------------------------------------------------------------------------+
                    | 5     | Structure Damping Ratio                                                                                |
                    +-------+--------------------------------------------------------------------------------------------------------+

                - OR Solid Sign Data :
                    +-------+-----------------------------+
                    | Index | Item                        |
                    +=======+=============================+
                    | 0     | Solid Sign Height  (H)      |
                    +-------+-----------------------------+
                    | 1     | Solid Sign M Dimension (M)  |
                    +-------+-----------------------------+
                    | 2     | Solid Sign N Dimension (N)  |
                    +-------+-----------------------------+
                    | 3     | Structure Natural Frequency |
                    +-------+-----------------------------+
                    | 4     | Structure Damping Ratio     |
                    +-------+-----------------------------+

                - OR Open Sign/Lattice Framework Data :
                    +-------+-----------------------------------------+
                    | Index | Item                                    |
                    +=======+=========================================+
                    | 0     | Orientation Type: Flat (0)/ Rounded (1) |
                    +-------+-----------------------------------------+
                    | 1     | Height (H)                              |
                    +-------+-----------------------------------------+
                    | 2     | Width                                   |
                    +-------+-----------------------------------------+
                    | 3     | Diameter of typical round member        |
                    +-------+-----------------------------------------+
                    | 4     | Structure Natural Frequency             |
                    +-------+-----------------------------------------+
                    | 5     | Structure Damping Ratio                 |
                    +-------+-----------------------------------------+
                    | 6     | Ratio of Solid Area to Gross Area       |
                    +-------+-----------------------------------------+

                - OR Trussed Tower Data :
                    +-------+--------------------------------------------------------+
                    | Index | Item                                                   |
                    +=======+========================================================+
                    | 0     | Horizontal Cross Sectio Type: Triangle (0)/ Square (1) |
                    +-------+--------------------------------------------------------+
                    | 1     | Height (H)                                             |
                    +-------+--------------------------------------------------------+
                    | 2     | Width                                                  |
                    +-------+--------------------------------------------------------+
                    | 3     | Structure Natural Frequency                            |
                    +-------+--------------------------------------------------------+
                    | 4     | Structure Damping Ratio                                |
                    +-------+--------------------------------------------------------+
                    | 5     | Ratio of Solid Area to Gross Area(in percetage)        |
                    +-------+--------------------------------------------------------+

        wallType : int
            Building wall to generate Wind Load on:
                +------------------+--------------------------------------+
                | Value            | Wall Type                            |
                +==================+======================================+
                | WindWard = 0     | WindWard                             |
                +------------------+--------------------------------------+
                | LeeWard = 1      | Leeward                              |
                +------------------+--------------------------------------+
                | SideWall = 2     | SideWall                             |
                +------------------+--------------------------------------+

                - (0 to 2 for WindWard, Leeward and SideWall, respectively).

        Returns
        -------
        int
            Returns number of Height or Intensity data.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.ComputeWallWindPressureProfileASCE72016(100, 10, 1, 1, 1, 1, [1], [1.0], [1.0], 0)
        """
        safe_UnitsDataList = make_safe_array_long_input(varUnitsData)
        safe_EscarpmentDataList = make_safe_array_double_input(varEscarpmentData)
        safe_BldgDataList = make_safe_array_double_input(varBldgData)
        retval = self._load.ComputeWallWindPressureProfileASCE72016(
            windSpeed,
            heightAboveSeaLvl,
            bldgClass,
            bldgType,
            expCat,
            int(bEscarpment),
            safe_UnitsDataList,
            safe_EscarpmentDataList,
            safe_BldgDataList,
            wallType,
        )
        if retval < 0:
            raise_os_error_if_error_code(retval)
        return retval

    def DeleteLoadEnvelop(self, varEnvNo: int):
        """
        Deletes a specified load envelop.

        Parameters
        ----------
        varEnvNo : int
            Load Envelop reference ID.

        Returns
        -------
        bool
            True if OK
            False if general error.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.DeleteLoadEnvelop(1)
        """
        return bool(self._load.DeleteLoadEnvelop(varEnvNo))

    def DeleteLoadList(self, varLoadListIndex: int):
        """
        Deletes specified load list.

        Parameters
        ----------
        varLoadListIndex : int
            Load list index.

        Returns
        -------
        bool
            True if OK

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.DeleteLoadList(1)
        """
        retVal = self._load.DeleteLoadList(varLoadListIndex)
        return bool(retVal)

    def DeletePrimaryLoadCases(self, varLoadCaseNos: list, varIsReferenceLoads: bool):
        """
        Deletes specified Primary/Reference Load Cases.

        Parameters
        ----------
        varLoadCaseNos : List of int
            List of Primary/Reference load case reference ID.
        varIsReferenceLoads : bool
            If reference load case(s): TRUE or FALSE

        Returns
        -------
        bool
            True if OK

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.DeletePrimaryLoadCases([1, 2, 3], False)
        """
        loadCaseNoList = make_safe_array_long_input(varLoadCaseNos)
        return bool(
            self._load.DeletePrimaryLoadCases(loadCaseNoList, varIsReferenceLoads)
        )

    def DeleteReferenceLoadCases(self, varLoadCaseNos: list):
        """
        Deletes specified Reference Load Cases.

        Parameters
        ----------
        varLoadCaseNos : List of int
            List of Reference load case reference ID.

        Returns
        -------
        bool
            True if OK

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.DeleteReferenceLoadCases([1, 2, 3])
        """
        loadCaseNoList = make_safe_array_long_input(varLoadCaseNos)
        return bool(self._load.DeleteReferenceLoadCases(loadCaseNoList))

    def DeleteWindDefinition(self, nTypeNo: int):
        """
        Deletes Wind definition. All defintions will be deleted if this input is set as 0.

        Parameters
        ----------
        nTypeNo : int
            Type of Wind.

        Returns
        -------
        bool
            True if OK

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.DeleteWindDefinition(1)
        """
        retVal = self._load.DeleteWindDefinition(nTypeNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def DeleteDirectAnalysisDefinitionParameter(self, pParamType: int):
        """
        Deletes respective parameters from Direct Analysis Definition based on the Parameter Type passed as argument (FLEX/AXIAL).

        Parameters
        ----------
        pParamType : int
            Integer indicating type of direct analysis parameter to be added. Integer value should be taken from following table :
                +----------+------------------------------------+
                | Value    | AnalysisCommand                    |
                +==========+====================================+
                |FLEX = 0  | DirectAnalysisParameterTypes.FLEX  |
                +----------+------------------------------------+
                |AXIAL = 2 | DirectAnalysisParameterTypes.AXIAL |
                +----------+------------------------------------+

        Returns
        -------
        int
            Returns True if successful.
            Returns False if unsuccessful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.DeleteDirectAnalysisDefinitionParameter(0)
        """
        return self._load.DeleteDirectAnalysisDefinitionParameter(pParamType)

    def DeleteDirectAnalysisDefinition(self):
        """
        Deletes whole Direct Analysis Definition.

        Returns
        -------
        int
            Returns True if successful.
            Returns False if unsuccessful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.DeleteDirectAnalysisDefinition()
        """
        return self._load.DeleteDirectAnalysisDefinition()

    def GetPrimaryLoadCaseCount(self):
        """
        Returns the total number of primary load cases in the current structure.

        Returns
        -------
        int
            Total number of primary load cases.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetPrimaryLoadCaseCount()
        """
        return self._load.GetPrimaryLoadCaseCount()

    def GetPrimaryLoadCaseNumbers(self):
        """
        Retrieves all primary load case numbers.

        Returns
        -------
        list of int
            List of load case reference number IDs.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetPrimaryLoadCaseNumbers()
        """
        primaryLoadCaseCount = self._load.GetPrimaryLoadCaseCount()
        primaryLoadCaseIdList_safe_list = make_safe_array_long(primaryLoadCaseCount)
        primaryLoadCaseIdArray = make_variant_vt_ref(
            primaryLoadCaseIdList_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        retVal = self._load.GetPrimaryLoadCaseNumbers(primaryLoadCaseIdArray)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return list(primaryLoadCaseIdArray[0])

    def GetLoadCombinationCaseCount(self):
        """
        Returns the total number of load combination cases in the current structure.

        Returns
        -------
        int
            Total number of load combination cases.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadCombinationCaseCount()
        """
        return self._load.GetLoadCombinationCaseCount()

    def GetLoadCombinationCaseNumbers(self):
        """
        Retrieves all load combination case numbers.

        Returns
        -------
        list of int
            List of load case reference number IDs.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadCombinationCaseNumbers()
        """
        loadCombinationCaseCount = self._load.GetLoadCombinationCaseCount()
        loadCombinationCaseId_safe_list = make_safe_array_long(loadCombinationCaseCount)
        loadCombinationLoadCaseIdArray = make_variant_vt_ref(
            loadCombinationCaseId_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        retVal = self._load.GetLoadCombinationCaseNumbers(
            loadCombinationLoadCaseIdArray
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return loadCombinationLoadCaseIdArray[0]

    def GetReferenceLoadCount(self):
        """
        Returns the number of reference load items in the currently active load case.

        Returns
        -------
        int
            Number of reference load items.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetReferenceLoadCount()
        """
        retVal = self._load.GetReferenceLoadCount()
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetReferenceLoadCaseCount(self):
        """
        Returns the number of reference load case items in the currently active load case.

        Returns
        -------
        int
            Number of reference load case items.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetReferenceLoadCaseCount()
        """
        retVal = self._load.GetReferenceLoadCaseCount()
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetReferenceLoadCaseNumbers(self):
        """
        Retrieves reference load case number IDs from Reference Load Definitions.

        Returns
        -------
        list of int
            List of reference load case IDs.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetReferenceLoadCaseNumbers()
        """
        refLoadCaseCount = self._load.GetReferenceLoadCaseCount()
        refLoadCaseIdList_safe_list = make_safe_array_long(refLoadCaseCount)
        refLoadCaseIdArray = make_variant_vt_ref(
            refLoadCaseIdList_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        retVal = self._load.GetReferenceLoadCaseNumbers(refLoadCaseIdArray)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return list(refLoadCaseIdArray[0])

    def GetNoOfSetsInReferenceLoad(self, nIndex: int):
        """
        Returns the number of reference load case-factor sets in a specified reference load item.

        Parameters
        ----------
        nIndex : int
            Index of the reference load case item.

        Returns
        -------
        int
            Number of sets in the reference load item.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetNoOfSetsInReferenceLoad(1)
        """
        retVal = self._load.GetNoOfSetsInReferenceLoad(nIndex)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetReferenceLoadByIndex(self, nIndex: int):
        """
        Retrieves a dictionary of load case numbers and their corresponding factors for a given reference load case.

        Parameters
        ----------
        nIndex : int
            Index of the reference load.

        Returns
        -------
        tuple of lists
            tuple of load case number factor lists. [[loadcase1, loadcase2, ...], [factor1, factor2, ...]]

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetReferenceLoadByIndex(1)
        """
        refLoadCaseCount = self.GetNoOfSetsInReferenceLoad(nIndex)
        if refLoadCaseCount <= 0:
            return
        refLoad_safe_array = make_safe_array_long(refLoadCaseCount)
        refloadArray_vt = make_variant_vt_ref(
            refLoad_safe_array, automation.VT_ARRAY | automation.VT_I4
        )
        factor_safe_array = make_safe_array_double(refLoadCaseCount)
        refFactorArray_vt = make_variant_vt_ref(
            factor_safe_array, automation.VT_ARRAY | automation.VT_R8
        )
        retval = self._load.GetReferenceLoadByIndex(
            nIndex, refloadArray_vt, refFactorArray_vt
        )
        if retval < 0:
            raise_os_error_if_error_code(retval)
        return list(refloadArray_vt[0]), list(refFactorArray_vt[0])

    def GetReferenceLoadType(self, varLoadNo: int):
        """
        Returns the type of a reference load.

        Parameters
        ----------
        varLoadNo : int
            Reference load number.

        Returns
        -------
        int
            Reference load type (0 to 23).

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetReferenceLoadType(1)
        """
        retVal = self._load.GetReferenceLoadType(varLoadNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetReferenceLoadCaseTitle(self, varLoadNo: int):
        """
        Returns the title of a reference load case.

        Parameters
        ----------
        varLoadNo : int
            Reference load number.

        Returns
        -------
        str
            Title of the reference load case.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetReferenceLoadCaseTitle(1)
        """
        return self._load.GetReferenceLoadCaseTitle(varLoadNo)

    def GetBeamCountAtFloor(
        self,
        varfMinX: float,
        varfMaxX: float,
        varfMinY: float,
        varfMaxY: float,
        varfMinZ: float,
        varfMaxZ: float,
        varnDirection: int,
    ):
        """
        Get the beam count at the specific floor.

        Parameters
        ----------
        varfMinX : float
            varfMinX X range start (in global coordinate)
        varfMaxX : float
            varfMaxX X range end (in global coordinate)
        varfMinY : float
            varfMinY Y range start (in global coordinate)
        varfMaxY : float
            varfMaxY Y range end (in global coordinate)
        varfMinZ : float
            varfMinZ Z range start (in global coordinate)
        varfMaxZ : float
            varfMaxZ Z range end (in global coordinate)
        varnDirection : int
            varnDirection Direction(1 for XRange, 2 for YRange, 3 for ZRange).

        Returns
        -------
        int
            The beam count at the specific floor.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetBeamCountAtFloor(0, 10, 0, 10, 0, 10, 1)
        """
        return self._load.GetBeamCountAtFloor(
            varfMinX, varfMaxX, varfMinY, varfMaxY, varfMinZ, varfMaxZ, varnDirection
        )

    def GetInfluenceArea(
        self,
        varfMinX: float,
        varfMaxX: float,
        varfMinY: float,
        varfMaxY: float,
        varfMinZ: float,
        varfMaxZ: float,
        varnDirection: int,
    ):
        """
        Returns a dictionary of beam to influence area at the specific floor.

        Parameters
        ----------
        varfMinX : float
            X range start (in global coordinate).
        varfMaxX : float
            X range end (in global coordinate).
        varfMinY : float
            Y range start (in global coordinate).
        varfMaxY : float
            Y range end (in global coordinate).
        varfMinZ : float
            Z range start (in global coordinate).
        varfMaxZ : float
            Z range end (in global coordinate).
        varnDirection : int
            Direction(1 for XRange, 2 for YRange, 3 for ZRange).

        Returns
        -------
        Dictionary
            Returns dictionary have beam id to influence area data.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetInfluenceArea(0, 10, 0, 10, 0, 10, 1)
        """
        beamCount = self._load.GetBeamCountAtFloor(
            varfMinX, varfMaxX, varfMinY, varfMaxY, varfMinZ, varfMaxZ, varnDirection
        )
        beamIdList_safe_list = make_safe_array_long(beamCount)
        influenceAreaList_safe_list = make_safe_array_double(beamCount)
        beamIdList = make_variant_vt_ref(
            beamIdList_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        influenceAreaList = make_variant_vt_ref(
            influenceAreaList_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        self._load.GetInfluenceArea(
            varfMinX,
            varfMaxX,
            varfMinY,
            varfMaxY,
            varfMinZ,
            varfMaxZ,
            varnDirection,
            beamIdList,
            influenceAreaList,
        )
        beamToAreaInfluence = {}
        for i in range(0, beamCount):
            beamToAreaInfluence[beamIdList[0][i]] = influenceAreaList[0][i]
        return beamToAreaInfluence

    def GetActiveLoad(self):
        """
        Returns the current load case number.

        Returns
        -------
        int
            Returns active load case number ID.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetActiveLoad()
        """
        retVal = self._load.GetActiveLoad()
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetNodalLoadCount(self, nNodeNo: int):
        """
        Returns number of nodal loads present for the specified node.

        Parameters
        ----------
        nNodeNo : int
            Node Id

        Returns
        -------
        int
            Returns the number of node(s).

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetNodalLoadCount(1)
        """
        retVal = self._load.GetNodalLoadCount(nNodeNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetNodalLoads(self, nNodeNo: int):
        """
        Returns tuple of list of forces in X direction, forces in Y direction, forces in Z direction, moments in X direction, moments in Y direction and moments in Z direction respectively.

        Parameters
        ----------
        nNodeNo : int
            Node Id

        Returns
        -------
        Tuple
            Returns a tuple of list of forces in X direction, forces in Y direction, forces in Z direction, moments in X direction, moments in Y direction and moments in Z direction respectively.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetNodalLoads(1)
        """
        nodeCount = self._load.GetNodalLoadCount(nNodeNo)
        varFX_safe_list = make_safe_array_double(nodeCount)
        varFY_safe_list = make_safe_array_double(nodeCount)
        varFZ_safe_list = make_safe_array_double(nodeCount)
        varMX_safe_list = make_safe_array_double(nodeCount)
        varMY_safe_list = make_safe_array_double(nodeCount)
        varMZ_safe_list = make_safe_array_double(nodeCount)
        varFXList = make_variant_vt_ref(
            varFX_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varFYList = make_variant_vt_ref(
            varFY_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varFZList = make_variant_vt_ref(
            varFZ_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varMXList = make_variant_vt_ref(
            varMX_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varMYList = make_variant_vt_ref(
            varMY_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varMZList = make_variant_vt_ref(
            varMZ_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        retVal = self._load.GetNodalLoads(
            nodeCount, varFXList, varFYList, varFZList, varMXList, varMYList, varMZList
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return (
            varFXList[0],
            varFYList[0],
            varFZList[0],
            varMXList[0],
            varMYList[0],
            varMZList[0],
        )

    def GetUDLLoadCount(self, nBeamNo: int):
        """
        Returns the number of uniformly distributed load(s) present for the specified beam.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        int
            Returns the number of uniformly distributed load item(s) applied.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetUDLLoadCount(1)
        """
        retVal = self._load.GetUDLLoadCount(nBeamNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetUDLLoads(self, nBeamNo: int):
        """
        Gets the uniformly distributed load(s) with all the parameters for the specified member.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        tuple
            Return a tuple of lists in which each list consist of
            - load directions (1 to 9 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ, ProjectedX, ProjectedY, ProjectedZ, respectively.)
            - magnitude of uniform force
            - distance from start of member to the start of load
            - distance from start of member to the end of load
            - perpendicular distance from the member shear center to the local plane of loading respectively.
            [[dirL1, dirL2,..], [forceL1, forceL2,..], [dst_startL1, dst_startL2,..], [dst_endL1, dst_endL2,..], [dst_perpendicularL1, dst_perpendicularL2,..]]

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetUDLLoads(1)
        """
        UDLLoadCount = self.GetUDLLoadCount(nBeamNo)
        if UDLLoadCount <= 0:
            return ([], [], [], [], [])
        varDirection_safe_list = make_safe_array_long(UDLLoadCount)
        varForce_safe_list = make_safe_array_double(UDLLoadCount)
        varD1_safe_list = make_safe_array_double(UDLLoadCount)
        varD2_safe_list = make_safe_array_double(UDLLoadCount)
        varD3_safe_list = make_safe_array_double(UDLLoadCount)
        varDirectionList = make_variant_vt_ref(
            varDirection_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        varForceList = make_variant_vt_ref(
            varForce_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD1List = make_variant_vt_ref(
            varD1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD2List = make_variant_vt_ref(
            varD2_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD3List = make_variant_vt_ref(
            varD3_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retval = self._load.GetUDLLoads(
            nBeamNo, varDirectionList, varForceList, varD1List, varD2List, varD3List
        )
        if retval < 0:
            raise_os_error_if_error_code(retval)
        if not bool(retval):
            return ([], [], [], [], [])
        return (
            list(varDirection_safe_list[0]),
            list(varForce_safe_list[0]),
            list(varD1_safe_list[0]),
            list(varD2_safe_list[0]),
            list(varD3_safe_list[0]),
        )

    def GetUNIMomentCount(self, nBeamNo: int):
        """
        Returns the count of uniformly distributed (UNI) moment applied to the specified member.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        int
            Returns the number of uniformly distributed (UNI) moment item(s) applied.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetUNIMomentCount(1)
        """
        retVal = self._load.GetUNIMomentCount(nBeamNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetUNIMoments(self, nBeamNo: int):
        """
        Returns the uniformly distributed (UNI) moments with all the parameters for the specified member.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        List of tuple
            Return a list of tuple in which tuple consist of load direction, magnitude of uniform moment, distance from start of member to the start of load, distance from start of member to the end of load, perpendicular distance from the member shear center to the local plane of loading respectively. Load direction will be represented numerically - 1 to 9 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ, ProjectedX, ProjectedY, ProjectedZ, respectively.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetUNIMoments(1)
        """

        UNILoadCount = self._load.GetUNIMomentCount(nBeamNo)
        varDirection_safe_list = make_safe_array_long(UNILoadCount)
        varMoment_safe_list = make_safe_array_double(UNILoadCount)
        varD1_safe_list = make_safe_array_double(UNILoadCount)
        varD2_safe_list = make_safe_array_double(UNILoadCount)
        varD3_safe_list = make_safe_array_double(UNILoadCount)
        varDirectionList = make_variant_vt_ref(
            varDirection_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        varMomentList = make_variant_vt_ref(
            varMoment_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD1List = make_variant_vt_ref(
            varD1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD2List = make_variant_vt_ref(
            varD2_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD3List = make_variant_vt_ref(
            varD3_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retVal = self._load.GetUNIMoments(
            nBeamNo, varDirectionList, varMomentList, varD1List, varD2List, varD3List
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        UNILoads = []
        for i in range(0, UNILoadCount):
            UNILoads.append(
                (
                    varDirectionList[0][i],
                    varMomentList[0][i],
                    varD1List[0][i],
                    varD2List[0][i],
                    varD3List[0][i],
                )
            )
        return UNILoads

    def GetTrapLoadCount(self, nBeamNo: int):
        """
        Returns number of trapezoidal load(s) present for the specified beam.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        int
            Returns the number of trapezoidal load item(s) applied.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetTrapLoadCount(1)
        """
        retVal = self._load.GetTrapLoadCount(nBeamNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetTrapLoads(self, nBeamNo: int):
        """
        Returns the trapezoidal load(s) with all the parameters for the specified member.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        List of tuple
            Return a list of tuple in which tuple consist of load direction, Load at the start of the member, Load at the end of the member, distance from the start of the member to loading starting point, distance from the end of the member to loading stopping point respectively. Load direction will be represented numerically - 1 to 9 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ, ProjectedX, ProjectedY, ProjectedZ, respectively.
        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetTrapLoads(1)
        """

        TrapezodialLoadCount = self._load.GetTrapLoadCount(nBeamNo)
        varDirection_safe_list = make_safe_array_long(TrapezodialLoadCount)
        varW1_safe_list = make_safe_array_double(TrapezodialLoadCount)
        varW2_safe_list = make_safe_array_double(TrapezodialLoadCount)
        varD1_safe_list = make_safe_array_double(TrapezodialLoadCount)
        varD2_safe_list = make_safe_array_double(TrapezodialLoadCount)
        varDirectionList = make_variant_vt_ref(
            varDirection_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        varW1List = make_variant_vt_ref(
            varW1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varW2List = make_variant_vt_ref(
            varW2_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD1List = make_variant_vt_ref(
            varD1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD2List = make_variant_vt_ref(
            varD2_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retVal = self._load.GetTrapLoads(
            nBeamNo, varDirectionList, varW1List, varW2List, varD1List, varD2List
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        TrapezodialLoads = []
        for i in range(0, TrapezodialLoadCount):
            TrapezodialLoads.append(
                (
                    varDirectionList[0][i],
                    varW1List[0][i],
                    varW2List[0][i],
                    varD1List[0][i],
                    varD2List[0][i],
                )
            )
        return TrapezodialLoads

    def GetConcForceCount(self, nBeamNo: int):
        """
        Get number of concentrated force(s) present for the specified beam.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        int
            Returns the number of concentrated force item(s) applied.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetConcForceCount(1)
        """
        retVal = self._load.GetConcForceCount(nBeamNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetConcForces(self, nBeamNo: int):
        """
        Returns the concentrated force(s) with all the parameters for the specified member.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        List of tuple
            Return a list of tuple in which tuple consist of load direction, Magnitude of the concentrate force, distance from the start of the member to concentrated force or moment, Perpendicular distance from the member shear center to the local plane of loading respectively. Load direction will be represented numerically - 1 to 6 for LocalX, LocalY, LocalZ, GlobalX, GlobalY and GlobalZ, respectively

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetConcForces(1)
        """

        ConcForceCount = self._load.GetConcForceCount(nBeamNo)
        varDirection_safe_list = make_safe_array_long(ConcForceCount)
        varForce_safe_list = make_safe_array_double(ConcForceCount)
        varD1_safe_list = make_safe_array_double(ConcForceCount)
        varD2_safe_list = make_safe_array_double(ConcForceCount)
        varDirectionList = make_variant_vt_ref(
            varDirection_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        varForceList = make_variant_vt_ref(
            varForce_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD1List = make_variant_vt_ref(
            varD1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD2List = make_variant_vt_ref(
            varD2_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retVal = self._load.GetConcForces(
            nBeamNo, varDirectionList, varForceList, varD1List, varD2List
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        ConcForces = []
        for i in range(0, ConcForceCount):
            ConcForces.append(
                (
                    varDirectionList[0][i],
                    varForceList[0][i],
                    varD1List[0][i],
                    varD2List[0][i],
                )
            )
        return ConcForces

    def GetConcMomentCount(self, nBeamNo: int):
        """
        Gets number of concentrated moment(s) present for the specified beam.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        int
            Returns the number of concentrated moment item(s) applied.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetConcMomentCount(1)
        """
        retVal = self._load.GetConcMomentCount(nBeamNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetConcMoments(self, nBeamNo: int):
        """
        Returns the concentrated moments(s) with all the parameters for the specified member.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        List of tuple
            Return a list of tuple in which tuple consist of load direction, Magnitude of the concentrate moment, distance from the start of the member to concentrated force or moment, Perpendicular distance from the member shear center to the local plane of loading respectively. Load direction will be represented numerically - 1 to 6 for LocalX, LocalY, LocalZ, GlobalX, GlobalY and GlobalZ, respectively

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetConcMoments(1)
        """

        ConcForceCount = self._load.GetConcForceCount(nBeamNo)
        if ConcForceCount < 0:
            raise_os_error_if_error_code(ConcForceCount)
        varDirection_safe_list = make_safe_array_long(ConcForceCount)
        varMoment_safe_list = make_safe_array_double(ConcForceCount)
        varD1_safe_list = make_safe_array_double(ConcForceCount)
        varD2_safe_list = make_safe_array_double(ConcForceCount)
        varDirectionList = make_variant_vt_ref(
            varDirection_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        varMomentList = make_variant_vt_ref(
            varMoment_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD1List = make_variant_vt_ref(
            varD1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varD2List = make_variant_vt_ref(
            varD2_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retVal = self._load.GetConcForces(
            nBeamNo, varDirectionList, varMomentList, varD1List, varD2List
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        ConcForces = []
        for i in range(0, ConcForceCount):
            ConcForces.append(
                (
                    varDirectionList[0][i],
                    varMomentList[0][i],
                    varD1List[0][i],
                    varD2List[0][i],
                )
            )
        return ConcForces

    def GetNoOfLoadAndFactorPairsForCombination(self, varLoadCombNo: int):
        """
        Gets the number of load case(s) applied with multiplication factor in specified load combination.

        Parameters
        ----------
        varLoadCombNo : int
            Combination Load case reference number ID.

        Returns
        -------
        int
            Returns the number of load cases in specified load combination.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetNoOfLoadAndFactorPairsForCombination(1)
        """
        return self._load.GetNoOfLoadAndFactorPairsForCombination(varLoadCombNo)

    def GetLoadAndFactorForCombination(self, varLoadCombNo: int):
        """
        Get number of concentrated force(s) present for the specified beam.

        Parameters
        ----------
        varLoadCombNo : int
            Combination Load case reference number ID.

        Returns
        -------
        Tuple
            Returns a Tuple consisting of a list of load case reference number IDs and list of multiplication factors. For SRSS, in multiplication factor list, an extra element is added, at the end. This factor represents overall multiplication factor for SRSS combination.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadAndFactorForCombination(1)
        """
        LoadCaseCount = self._load.GetNoOfLoadAndFactorPairsForCombination(
            varLoadCombNo
        )
        varLoadCaseID_safe_list = make_safe_array_long(LoadCaseCount)
        varMultiplicationFactor_safe_list = make_safe_array_double(LoadCaseCount + 1)
        varLoadCaseIDList = make_variant_vt_ref(
            varLoadCaseID_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        varMultiplicationFactorList = make_variant_vt_ref(
            varMultiplicationFactor_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        self._load.GetLoadAndFactorForCombination(
            varLoadCombNo, varLoadCaseIDList, varMultiplicationFactorList
        )
        return (varLoadCaseIDList[0], varMultiplicationFactorList[0])

    def GetLoadCaseTitle(self, varLoadNo: int):
        """
        Returns title of the specified load case as a text string. Input 0 to retrieve title of current active load case or reference load case

        Parameters
        ----------
        varLoadNo : int
            The load case string title.

        Returns
        -------
        str
            Returns the load case string title.
            Returns "NONE" if load case varLoadNo not found.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadCaseTitle(1)
        """
        return self._load.GetLoadCaseTitle(varLoadNo)

    def GetElementPressureLoadCount(self, varPlateNo: int):
        """
        Gets the number pressure load(s) for the specified plate.

        Parameters
        ----------
        varPlateNo : int
            Plate number ID.

        Returns
        -------
        str
            Returns the number of pressure load(s).

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetElementPressureLoadCount(1)
        """
        retVal = self._load.GetElementPressureLoadCount(varPlateNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetElementPressureLoads(self, varPlateNo: int):
        """
        Returns the pressure load(s) with all the parameters for the specified plate.

        Parameters
        ----------
        varPlateNo : int
            Plate number ID.

        Returns
        -------
        List of tuple
            Returns a list of tuple in which tuple consist of load direction, Magnitude of the pressure load(s), Top-Left coordinate X (local), Top-Left coordinate Y (local), Bottom-Right coordinate X (local) and Bottom-Right coordinate Y (local) respectively. Load direction will be represented numerically - 1 to 9 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ, ProjectedX, ProjectedY and ProjectedZ respectively.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetElementPressureLoads(1)
        """

        PressureLoadCount = self._load.GetElementPressureLoadCount(varPlateNo)
        varDirection_safe_list = make_safe_array_long(PressureLoadCount)
        varW1_safe_list = make_safe_array_double(PressureLoadCount)
        varX1_safe_list = make_safe_array_double(PressureLoadCount)
        varY1_safe_list = make_safe_array_double(PressureLoadCount)
        varX2_safe_list = make_safe_array_double(PressureLoadCount)
        varY2_safe_list = make_safe_array_double(PressureLoadCount)
        varDirectionList = make_variant_vt_ref(
            varDirection_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        varW1List = make_variant_vt_ref(
            varW1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varX1List = make_variant_vt_ref(
            varX1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varY1List = make_variant_vt_ref(
            varY1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varX2List = make_variant_vt_ref(
            varX2_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varY2List = make_variant_vt_ref(
            varY2_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retVal = self._load.GetElementPressureLoads(
            PressureLoadCount,
            varDirectionList,
            varW1List,
            varX1List,
            varY1List,
            varX2List,
            varY2List,
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        PressureLoads = []
        for i in range(0, PressureLoadCount):
            PressureLoads.append(
                (
                    varDirectionList[0][i],
                    varW1List[0][i],
                    varX1List[0][i],
                    varY1List[0][i],
                    varX2List[0][i],
                    varY2List[0][i],
                )
            )
        return PressureLoads

    def GetElementConcLoadCount(self, varPlateNo: int):
        """
        Returns the number of concentrated load for specified plate.

        Parameters
        ----------
        varPlateNo : int
            Plate number ID.

        Returns
        -------
        int
            Returns the number of concentrated load on specified plate.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetElementConcLoadCount(1)
        """
        retVal = self._load.GetElementConcLoadCount(varPlateNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetElementConcLoads(self, varPlateNo: int):
        """
        Returns the concentrated load(s) with all the parameters for the specified plate.

        Parameters
        ----------
        varPlateNo : int
            Plate number ID.

        Returns
        -------
        List of tuple
            Returns a list of tuple in which tuple consist of load direction, pressure, the first d x coordinate, the first d y coordinate respectively. Load direction will be represented numerically - 1 to 9 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ, ProjectedX, ProjectedY and ProjectedZ respectively.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetElementConcLoads(1)
        """

        ConcentratedLoadCount = self.GetElementPressureLoadCount(varPlateNo)
        varDirection_safe_list = make_safe_array_long(ConcentratedLoadCount)
        varW1_safe_list = make_safe_array_double(ConcentratedLoadCount)
        varX1_safe_list = make_safe_array_double(ConcentratedLoadCount)
        varY1_safe_list = make_safe_array_double(ConcentratedLoadCount)
        varDirectionList = make_variant_vt_ref(
            varDirection_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        varW1List = make_variant_vt_ref(
            varW1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varX1List = make_variant_vt_ref(
            varX1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varY1List = make_variant_vt_ref(
            varY1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retVal = self._load.GetElementConcLoads(
            ConcentratedLoadCount, varDirectionList, varW1List, varX1List, varY1List
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        ConcentratedLoads = []
        for i in range(0, ConcentratedLoadCount):
            ConcentratedLoads.append(
                (
                    varDirectionList[0][i],
                    varW1List[0][i],
                    varX1List[0][i],
                    varY1List[0][i],
                )
            )
        return ConcentratedLoads

    def GetLoadType(self, varLoadNo: int):
        """
        Returns primary load case category(s) as an long value.

        Parameters
        ----------
        varLoadNo : int
            Primary load case reference ID. Pass in 0 to get information about current active load case or reference load case

        Returns
        -------
        int
            Returns 0 if Dead.
            Returns 1 if Live.
            Returns 2 if Roof Live.
            Returns 3 if Wind.
            Returns 4 if Seismic-H.
            Returns 5 if Seismic-V.
            Returns 6 if Snow.
            Returns 7 if Fluids.
            Returns 8 if Soil.
            Returns 9 if Rain.
            Returns 10 if Ponding.
            Returns 11 if Dust.
            Returns 12 if Traffic.
            Returns 13 if Temperature.
            Returns 14 if Imperfection.
            Returns 15 if Accidental.
            Returns 16 if Flood.
            Returns 17 if Ice.
            Returns 18 if Wind Ice.
            Returns 19 if Crane Hook.
            Returns 20 if Mass.
            Returns 21 if Gravity.
            Returns 22 if Push.
            Returns 23 if None.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadType(1)
        """
        retVal = self._load.GetLoadType(varLoadNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetLoadListCount(self):
        """
        Gets the number of existing load list(s)

        Returns
        -------
        int
            Returns the number of load list(s).

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadListCount()
        """
        retVal = self._load.GetLoadListCount()
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetLoadCountInLoadList(self, varLoadListIndex: int):
        """
        Gets the number of load case(s) in specified load list.

        Parameters
        ----------
        varLoadListIndex : int
            Load list index.

        Returns
        -------
        int
            The number of Load Case(s) in specified Load List.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadCountInLoadList(1)
        """
        return self._load.GetLoadCountInLoadList(varLoadListIndex)

    def GetLoadsInLoadList(self, varLoadListIndex: int):
        """
        Gets the load case(s) in specified load list.

        Parameters
        ----------
        varLoadListIndex : int
            Load list index(Starts from one).

        Returns
        -------
        list of int
            Load  Case reference IDs list.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadsInLoadList(1)
        """
        loadListCount = self.GetLoadCountInLoadList(varLoadListIndex)
        if loadListCount == 0:
            return []
        varLoad_safe_list = make_safe_array_long(loadListCount)
        varLoadList = make_variant_vt_ref(
            varLoad_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        self._load.GetLoadsInLoadList(varLoadListIndex, varLoadList)
        return varLoadList[0]

    def GetAttribute(self, lLoadCase: int):
        """
        Gets load attribute information of specified load case.

        Parameters
        ----------
        lLoadCase : int
            Load case reference ID.

        Returns
        -------
        bool
            True if OK.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetAttribute(1)
        """
        retVal = self._load.GetAttribute(lLoadCase)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return True

    def GetRepeatLoadCount(self):
        """
        Returns the number of repeat load commands in the active load case.

        Returns
        -------
        int
            Returns the number of repeat load commands in the active load case.
            Returns 0 if General Error

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetRepeatLoadCount()
        """
        return self._load.GetRepeatLoadCount()

    def GetNoLoadFactorInRepeatLoad(self, nIndex: int):
        """
        Returns the number of load and factor pairs associated with a given repeat load command in the active load case.

        Parameters
        ----------
        nIndex : int
            The index(One based) for repeat load.

        Returns
        -------
        int
            Returns number of load and factor pairs associated with a given repeat load command in the active load case.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetNoLoadFactorInRepeatLoad(1)
        """
        retVal = self._load.GetNoLoadFactorInRepeatLoad(nIndex)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetRepeatLoadByIndex(self, nIndex: int):
        """
        Returns the dictionary of load case IDs to load factors for a given repeat load command in the active load case.

        Parameters
        ----------
        nIndex : int
            The index(One based) for repeat load.

        Returns
        -------
        dictionary
            Returns a dictionary of load case ID to load factor.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetRepeatLoadByIndex(1)
        """

        loadSizeCount = self._load.GetNoLoadFactorInRepeatLoad(nIndex)
        varLoadCase_safe_list = make_safe_array_long(loadSizeCount)
        varLoadFactor_safe_list = make_safe_array_long(loadSizeCount)
        varLoadCaseList = make_variant_vt_ref(
            varLoadCase_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        varLoadFactorList = make_variant_vt_ref(
            varLoadFactor_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        retVal = self._load.GetRepeatLoadByIndex(
            nIndex, varLoadCaseList, varLoadFactorList
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)

        loadCaseToFactor = {}
        for i in range(0, loadSizeCount):
            loadCaseToFactor[varLoadCaseList[0][i]] = varLoadFactorList[0][i]

        return loadCaseToFactor

    def GetLinearVaryingLoadCount(self, nBeamNo: int):
        """
        Returns number of linear varying load(s) present for the specified beam.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        int
            Returns the number of linear varying load item(s) applied.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLinearVaryingLoadCount(1)
        """
        retVal = self._load.GetLinearVaryingLoadCount(nBeamNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetLinearVaryingLoads(self, nBeamNo: int):
        """
        Returns parameters for defining linear varying loads for specified beam.

        Parameters
        ----------
        nBeamNo : int
            Beam number ID.

        Returns
        -------
        List of tuple
            Returns a list of tuple in which tuple consist of load direction, load at the start of the member, load at the end of the member, Load in the middle of the member (for triangular load) respectively. Load direction will be represented numerically - 1 to 3 for local X, Y and Z, respectively

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLinearVaryingLoads(1)
        """

        LinearVaryingLoadCount = self._load.GetLinearVaryingLoadCount(nBeamNo)
        varDirection_safe_list = make_safe_array_long(LinearVaryingLoadCount)
        varW1_safe_list = make_safe_array_double(LinearVaryingLoadCount)
        varW2_safe_list = make_safe_array_double(LinearVaryingLoadCount)
        varW3_safe_list = make_safe_array_double(LinearVaryingLoadCount)
        varDirectionList = make_variant_vt_ref(
            varDirection_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        varW1List = make_variant_vt_ref(
            varW1_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varW2List = make_variant_vt_ref(
            varW2_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varW3List = make_variant_vt_ref(
            varW3_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retVal = self._load.GetLinearVaryingLoads(
            LinearVaryingLoadCount, varDirectionList, varW1List, varW2List, varW3List
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        LinearVaryingLoads = []
        for i in range(0, LinearVaryingLoadCount):
            LinearVaryingLoads.append(
                (
                    varDirectionList[0][i],
                    varW1List[0][i],
                    varW2List[0][i],
                    varW3List[0][i],
                )
            )
        return LinearVaryingLoads

    def GetLoadTypeCount(self, loadType: int):
        """
        Gets the number of load(s) with specified Load Type in active Load Case.

        Parameters
        ----------
        loadType : int
            +-------+-----------------------------------+---+----------------------------------------+
            | Value | LoadType                          | * | Value | LoadType                       |
            +=======+===================================+===+=======+================================+
            | 4000  | SelfWeight                        |   | 3275  | Uniform Force (Physical)       |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3110  | Nodal Load (Node)                 |   | 3280  | Uniform Moment (Physical)      |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3120  | Nodal Load (Inclined)             |   | 3285  | Concentrated Force (Physical)  |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3910  | Nodal Load (Support Displacement) |   | 3290  | Concentrated Moment (Physical) |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3210  | Uniform Force                     |   | 3295  | Trapezoidal (Physical)         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3220  | Uniform Moment                    |   | 3310  | Pressure on full plate         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3230  | Concentrated Force                |   | 3310  | Concentrated Load (Plate)      |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3240  | Concentrated Moment               |   | 3310  | Partial plate pressure load    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3250  | Linear Varying                    |   | 3320  | Trapezoidal (Plate)            |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3260  | Trapezoidal                       |   | 3322  | Solid                          |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3260  | Hydrostatic                       |   | 3710  | Temperature                    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3620  | Pre/Post Stress                   |   | 3720  | Strain                         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3810  | Fixed End                         |   | 3721  | Strain Rate                    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3530  | FloorLoadGroup                    |   | 3410  | Area                           |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3554  | OneWayFloorLoadGroup              |   |       |                                |
            +-------+-----------------------------------+---+-------+--------------------------------+

        Returns
        -------
        int
            Returns the number of load(s).
            Returns 0 if loadCaseNo not found.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadTypeCount(1)
        """
        return self._load.GetLoadTypeCount(loadType)

    def GetListSizeForLoadType(self, loadType: int, loadIndex: int):
        """
        Gets number of entities to vwhich specified Load Type and load index.

        Parameters
        ----------
        loadType : int
            +-------+-----------------------------------+---+-----+----------------------------------+
            | Value | LoadType                          | * | Value | LoadType                       |
            +=======+===================================+===+=======+================================+
            | 4000  | SelfWeight                        |   | 3275  | Uniform Force (Physical)       |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3110  | Nodal Load (Node)                 |   | 3280  | Uniform Moment (Physical)      |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3120  | Nodal Load (Inclined)             |   | 3285  | Concentrated Force (Physical)  |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3910  | Nodal Load (Support Displacement) |   | 3290  | Concentrated Moment (Physical) |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3210  | Uniform Force                     |   | 3295  | Trapezoidal (Physical)         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3220  | Uniform Moment                    |   | 3310  | Pressure on full plate         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3230  | Concentrated Force                |   | 3310  | Concentrated Load (Plate)      |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3240  | Concentrated Moment               |   | 3310  | Partial plate pressure load    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3250  | Linear Varying                    |   | 3320  | Trapezoidal (Plate)            |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3260  | Trapezoidal                       |   | 3322  | Solid                          |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3260  | Hydrostatic                       |   | 3710  | Temperature                    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3620  | Pre/Post Stress                   |   | 3720  | Strain                         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3810  | Fixed End                         |   | 3721  | Strain Rate                    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3530  | FloorLoadGroup                    |   | 3410  | Area                           |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3554  | OneWayFloorLoadGroup              |   |       |                                |
            +-------+-----------------------------------+---+-------+--------------------------------+

        loadIndex : int
            Load item index of specified load type (Zero based). Program returns the first one start from loadIndex if exist the same load type in specified load case.

        Returns
        -------
        int
            Returns the number of entities.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetListSizeForLoadType(1, 0)
        """
        return self._load.GetListSizeForLoadType(loadType, loadIndex)

    def GetAssignmentListForLoadType(self, loadType: int, loadIndex: int):
        """
        Return the list of entities that have been assigned to a load command in the active load case (see function SetLoadActive). This command is identified as the LoadType, which has a defined reference number (see below) and an index number starting from 0. That is, if a load command has been defined 10 times in the load case, then index 0 identifies the first instance of the command and index 9 identifies the 10th.

        Parameters
        ----------
        loadType : int
            +-------+-----------------------------------+-+-------+----------------------------------+
            | Value | LoadType                          | * | Value | LoadType                       |
            +=======+===================================+===+=======+================================+
            | 4000  | SelfWeight                        |   | 3275  | Uniform Force (Physical)       |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3110  | Nodal Load (Node)                 |   | 3280  | Uniform Moment (Physical)      |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3120  | Nodal Load (Inclined)             |   | 3285  | Concentrated Force (Physical)  |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3910  | Nodal Load (Support Displacement) |   | 3290  | Concentrated Moment (Physical) |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3210  | Uniform Force                     |   | 3295  | Trapezoidal (Physical)         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3220  | Uniform Moment                    |   | 3310  | Pressure on full plate         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3230  | Concentrated Force                |   | 3310  | Concentrated Load (Plate)      |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3240  | Concentrated Moment               |   | 3310  | Partial plate pressure load    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3250  | Linear Varying                    |   | 3320  | Trapezoidal (Plate)            |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3260  | Trapezoidal                       |   | 3322  | Solid                          |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3260  | Hydrostatic                       |   | 3710  | Temperature                    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3620  | Pre/Post Stress                   |   | 3720  | Strain                         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3810  | Fixed End                         |   | 3721  | Strain Rate                    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3530  | FloorLoadGroup                    |   | 3410  | Area                           |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3554  | OneWayFloorLoadGroup              |   |       |                                |
            +-------+-----------------------------------+---+-------+--------------------------------+

        loadIndex : int
            Load item index of specified load type (Zero based). Program returns the first one start from loadIndex if exist the same load type in specified load case.

        Returns
        -------
        list of int
            Returns List of Entities number ID(s)

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetAssignmentListForLoadType(1, 0)
        """
        size = self.GetListSizeForLoadType(loadType, loadIndex)
        if size < 1:
            return []
        entities = make_safe_array_long(size)
        entities_ref = make_variant_vt_ref(
            entities, automation.VT_ARRAY | automation.VT_I4
        )
        retval = self._load.GetAssignmentListForLoadType(
            loadType, loadIndex, entities_ref
        )
        if retval == 0:
            return []
        return list(entities[0])

    def GetNodalLoadInfo(self, loadIndex: int):
        """
        Gets nodal load(s) generated by specified load item in specified load case.

        Parameters
        ----------
        loadIndex : int
            Load item index.

        Returns
        -------
        bool
            Returns a list of 5 nodal forces - FX, FY, FZ, MX, MY and MZ which are placed respectively.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetNodalLoadInfo(1)
        """
        nodalForce_safe_list = make_safe_array_double(6)
        nodalForceList = make_variant_vt_ref(
            nodalForce_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        self._load.GetNodalLoadInfo(loadIndex, nodalForceList)
        return nodalForceList[0]

    def GetMemberLoadInfo(self, loadIndex: int):
        """
        Gets member load(s) information generated by specified load item in specified load case.

        Parameters
        ----------
        loadIndex : int
            Load item index (Zero based)

        Returns
        -------
        tuple containing
            - direction: int  (Load direction will be represented numerically -  1 to 9 for LocalX, LocalY, LocalZ, GlobalX, GlobalY, GlobalZ, ProjectedX, ProjectedY, ProjectedZ, respectively.)
            - member force parameters: List [dW1, dW2, dW3]
            - member force distances: List [dD1, dD2, dD3]

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetMemberLoadInfo(1)
        """
        loadCount = 3
        varDirection_safe_list = make_safe_array_long(1)
        varForce_safe_list = make_safe_array_double(loadCount)
        varDistance_safe_list = make_safe_array_double(loadCount)
        varDirectionList = make_variant_vt_ref(varDirection_safe_list, automation.VT_I4)
        varForceList = make_variant_vt_ref(
            varForce_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varDistanceList = make_variant_vt_ref(
            varDistance_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retval = self._load.GetMemberLoadInfo(
            loadIndex, varDirectionList, varForceList, varDistanceList
        )
        if not bool(retval):
            return 0, [0, 0, 0], [0, 0, 0]
        return varDirectionList[0], varForceList[0], varDistanceList[0]

    def GetElementLoadInfo(self, loadIndex: int):
        """
        Gets element load information generated by specified load item in specified load case. use setLoadActive API to active load case, or select the loadCase from UI (ElementLoadPressure only supported).

        Parameters
        ----------
        loadIndex : int
            Load item index (Zero based)

        Returns
        -------
        List of tuple
            Returns a list of tuple in which tuple consist of load direction, element pressures : dW1, dW2, dW3 and dW4, element force distances: dX1, dY1, dX2 and dY2 respectively. Load direction will be represented numerically -  = 0 to 8 for LocalX, LocalY, LocalZ, GlobalX, GlobalY and GlobalZ, ProjectedX, ProjectedY, ProjectedZ respectively.

        see Also
        --------
        setLoadActive : To activate a specific load case.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetElementLoadInfo(1)
        """
        loadCount = 4
        safe_dir = make_safe_array_long(1)
        varDirection = make_variant_vt_ref(safe_dir, automation.VT_I4)
        varForce_safe_list = make_safe_array_double(loadCount)
        varDistance_safe_list = make_safe_array_double(loadCount)
        varForceList = make_variant_vt_ref(
            varForce_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        varDistanceList = make_variant_vt_ref(
            varDistance_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retVal = self._load.GetElementLoadInfo(
            loadIndex, varDirection, varForceList, varDistanceList
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        Loads = [varDirection[0], varForceList[0], varDistanceList[0]]
        return Loads

    def GetNotionalLoadCount(self):
        """
        Returns the number of Notional load.

        Returns
        -------
        int
            Returns the number of Notional load

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetNotionalLoadCount()
        """
        retVal = self._load.GetNotionalLoadCount()
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetNoLoadFactorDirectionInNotionalLoad(self, nIndex: int):
        """
        Gets the no of factor for specified Notional load.

        Parameters
        ----------
        nIndex : int
            The index for Notional load.

        Returns
        -------
        int
            Returns the factor for specified Notional load.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetNoLoadFactorDirectionInNotionalLoad(1)
        """
        retVal = self._load.GetNoLoadFactorDirectionInNotionalLoad(nIndex)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetNotionalLoadByIndex(self, nIndex: int):
        """
        Gets load case(s), direction(s) and factor(s) for specified Notional load.

        Parameters
        ----------
        nIndex : int
            The index for Notional load.

        Returns
        -------
        List of tuple
            Returns a list of tuple in which tuple consist of load direction, load case reference ID(s) in VARIANT array - (if +ve values = Primary Load Cases or -ve values = Reference Load Cases) , load factors respectively. Load direction will be represented numerically -  = 1 to 3 for X, Y and Z direction respectively.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetNotionalLoadByIndex(1)
        """
        notionalloadCount = self._load.GetNoLoadFactorDirectionInNotionalLoad(nIndex)
        if notionalloadCount < 0:
            raise_os_error_if_error_code(notionalloadCount)
        Direction_safe_list = make_safe_array_long(notionalloadCount)
        LoadCase_safe_list = make_safe_array_double(notionalloadCount)
        Factor_safe_list = make_safe_array_double(notionalloadCount)
        DirectionList = make_variant_vt_ref(
            Direction_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        LoadCaseList = make_variant_vt_ref(
            LoadCase_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        FactorList = make_variant_vt_ref(
            Factor_safe_list, automation.VT_ARRAY | automation.VT_R8
        )

        retVal = self._load.GetElementLoadInfo(
            notionalloadCount, LoadCaseList, FactorList, DirectionList
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        Loads = []
        for i in range(0, notionalloadCount):
            Loads.append((DirectionList[0][i], LoadCaseList[0][i], FactorList[0][i]))
        return Loads

    def GetLoadItemsCount(self, loadCaseNo: int):
        """
        Returns the number of loaditems in the specified load case.

        Parameters
        ----------
        loadCaseNo : int
            Load case number.

        Returns
        -------
        int
            Returns the number of loaditems in the specified load case.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadItemsCount(1)
        """
        retVal = self._load.GetLoadItemsCount(loadCaseNo)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetLoadItemType(self, loadCaseNo: int, loadItemIndex: int):
        """
        Returns the load item type for the specified loadIndex and loadCase.

        Parameters
        ----------
        loadCaseNo : int
            Load case number.
        loadItemIndex : int
            Load item index (Zero based).

        Returns
        -------
        int
            Returns LoadItemType for the specified loadIndex and loadCase.
            Returns 0 if LoadCase/LoadItemIndex Not Found.
            Return Values and LoadItem Type
            +-------+-----------------------------------+---+-------+--------------------------------+
            | Value | LoadItem Type                     | * | Value | LoadItem Type                  |
            +=======+===================================+===+=======+================================+
            | 4000  | SelfWeight                        |   | 3520  | FloorLoadZrange                |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3110  | Nodal Load (Node)                 |   | 3530  | FloorLoadGroup                 |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3120  | Nodal Load (Inclined)             |   | 3551  | OneWayFloorLoadXrange          |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3910  | Nodal Load (Support Displacement) |   | 3552  | OneWayFloorLoadYrange          |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3312  | Nodal Load (Region node load)     |   | 3553  | OneWayFloorLoadZrange          |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3210  | Uniform Force                     |   | 3554  | OneWayFloorLoadGroup           |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3220  | Uniform Moment                    |   | 3310  | Pressure on full plate         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3230  | Concentrated Force                |   | 3311  | Concentrated Load (Plate)      |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3240  | Concentrated Moment               |   | 3312  | Partial plate pressure load    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3250  | Linear Varying                    |   | 3320  | Trapezoidal (Plate)            |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3260  | Trapezoidal                       |   | 3322  | Solid                          |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3261  | Hydrostatic                       |   | 3710  | Temperature                    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3620  | Pre/Post Stress                   |   | 3720  | Strain                         |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3810  | Fixed End                         |   | 3721  | Strain Rate                    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3275  | Uniform Force (Physical)          |   | 4400  | UBC Load                       |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3280  | Uniform Moment (Physical)         |   | 4600  | Wind Load                      |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3285  | Concentrated Force (Physical)     |   | 4610  | Wind Load Dynamic              |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3290  | Concentrated Moment (Physical)    |   | 4405  | IbcLoad                        |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3295  | Trapezoidal (Physical)            |   | 4410  | 1893Load                       |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3410  | Area                              |   | 4500  | AijLoad                        |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3510  | FloorLoadYrange                   |   | 4510  | ColombianLoad                  |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 3511  | FloorLoadXrange                   |   | 4520  | CFELoad                        |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4570  | TurkishLoad                       |   | 4530  | RPALoad                        |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4575  | GB50011Load                       |   | 4540  | NTCLoad                        |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4576  | Colombian2010Load                 |   | 4550  | NRCLoad                        |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4820  | TimeHistoryLoad                   |   | 4560  | NRCLoad2005                    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4651  | Snow Load Data                    |   | 4561  | NRCLoad2010                    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4201  | Repeat load data                  |   | 4100  | Spectrum Load                  |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4223  | Notional Load Data                |   | 4700  | Calulate Natural Frequency     |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4220  | Reference Load                    |   | 4710  | Modal Calculation Requested    |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4101  | Spectrum Data                     |   | 4222  | Notional Load                  |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4701  | Calulate Rayleigh Frequency       |   | 4650  | Snow Load                      |
            +-------+-----------------------------------+---+-------+--------------------------------+
            | 4200  | Repeat load                       |   |       |                                |
            +-------+-----------------------------------+---+-------+--------------------------------+

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadItemType(1, 1)

        """
        return self._load.GetLoadItemType(loadCaseNo, loadItemIndex)

    def GetEnvelopeCount(self):
        """
        Returns number of Envelopes defined.

        Returns
        -------
        int
            Total Number of load Envelopes present.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetEnvelopeCount()
        """
        return self._load.GetEnvelopeCount()

    def GetLoadEnvelopeDetails(self, EnvNo: int):
        """
        Returns

        Parameters
        ----------
        EnvNo : int
            Load Envelope reference ID.

        Returns
        -------
        Tuple
            Returns a tuple containing EnvelopeType and NumberofLoadCasesInEnvelope information respectively.
            Type of Load Envelope
            +-------+-------------------+
            | Value | Load Envelop Type |
            +=======+===================+
            | 0     | NONE              |
            +-------+-------------------+
            | 1     | STRESS            |
            +-------+-------------------+
            | 2     | SERVICEABILITY    |
            +-------+-------------------+
            | 3     | COLUMN            |
            +-------+-------------------+
            | 4     | CONNECTION        |
            +-------+-------------------+
            | 5     | STRENGTH          |
            +-------+-------------------+
            | 6     | TEMPORARY         |
            +-------+-------------------+

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetEnvelopeDetails(1)
        """
        safe_EnvelopeType = make_safe_array_long(0)
        EnvelopeType = make_variant_vt_ref(safe_EnvelopeType, automation.VT_I4)

        safe_NumberofLoadCasesInEnvelope = make_safe_array_long(0)
        NumberofLoadCasesInEnvelope = make_variant_vt_ref(
            safe_NumberofLoadCasesInEnvelope, automation.VT_I4
        )

        retVal = self._load.GetLoadEnvelopeDetails(
            EnvNo, EnvelopeType, NumberofLoadCasesInEnvelope
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return (EnvelopeType[0], NumberofLoadCasesInEnvelope[0])

    def GetLoadListfromLoadEnvelope(self, EnvNo: int):
        """
        Gets the list of primary load case reference Ids present in the load envelope passed.

        Parameters
        ----------
        EnvNo : int
            Load Envelope reference ID

        Returns
        -------
        List of int
            (Primary) load case(s) reference ID(s).

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetLoadListfromLoadEnvelope(1)
        """
        LoadEnvelopeDetails = self.GetLoadEnvelopeDetails(EnvNo)
        LoadCase_safe_list = make_safe_array_long(LoadEnvelopeDetails[1])
        LoadCaseList = make_variant_vt_ref(
            LoadCase_safe_list, automation.VT_ARRAY | automation.VT_I4
        )

        retval = self._load.GetLoadListfromLoadEnvelope(EnvNo, LoadCaseList)
        if retval < 0:
            raise_os_error_if_error_code(retval)
        if retval <= 0:
            return []
        return list(retval)

    def GetEnvelopeIDs(self):
        """
        Gets the list of Loads Envelope IDs present in the staad file.

        Returns
        -------
        List of int
            Envelope ID(s)

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Load.GetEnvelopeIDs()
        """
        EnvelopeIDCount = self._load.GetEnvelopeCount()
        EnvelopeId_safe_list = make_safe_array_long(EnvelopeIDCount)
        EnvelopeIdList = make_variant_vt_ref(
            EnvelopeId_safe_list, automation.VT_ARRAY | automation.VT_I4
        )
        retVal = self._load.GetEnvelopeIDs(EnvelopeIdList)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal
