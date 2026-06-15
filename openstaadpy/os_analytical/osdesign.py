# ---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
# ---------------------------------------------------------------------------------------------
from .openStaadHelper import (
    make_safe_array_long_input,
    make_variant_vt_ref,
)
from comtypes import automation, CoInitialize
import comtypes.client as cc
from .oserrors import raise_os_error_if_error_code


class OSDesign:
    CoInitialize()

    def __init__(self, staadObj):
        self._staad = staadObj
        self._design = self._staad.Design

        self._functions = [
            "CreateDesignBrief",
            "AssignDesignParameter",
            "AssignDesignCommand",
            "AssignDesignGroup",
            "GetDesignBriefCode",
        ]

        for function_name in self._functions:
            self._design._FlagAsMethod(function_name)

    # Design FUNCTIONS

    def CreateDesignBrief(self, design_code: int):
        """
        Create a new design brief with the specified design code.

        Parameters
        ----------
        design_code : int
            Design code index.

        Returns
        -------
        int
            Reference ID of the created design brief.

        Example
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> ref_id = staad_obj.Design.CreateDesignBrief(1001)
        """
        retVal = self._design.CreateDesignBrief(design_code)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def AssignDesignParameter(
        self,
        design_ref_id: int,
        design_param: str,
        design_param_value: str,
        member_ids: list | int,
    ):
        """
        Assign design parameters to a specified design brief.

        Parameters
        ----------
        design_ref_id : int
            Design brief reference ID.
        design_param : str
            Name of the design parameter.
        design_param_value : str
            Value for the design parameter.
        member_ids : list of int or int
            List of member numbers.

        Returns
        -------
        bool
            True if successful, False otherwise.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> ref_id = staad_obj.Design.CreateDesignBrief(1001)
        >>> result = staad_obj.Design.AssignDesignParameter(ref_id, "BEAM", "1", [1, 3])
        """
        if isinstance(member_ids, int):
            member_ids = [member_ids]
        safe_members = make_safe_array_long_input(member_ids)
        members_variant = make_variant_vt_ref(
            safe_members, automation.VT_ARRAY | automation.VT_I4
        )
        retVal = self._design.AssignDesignParameter(
            design_ref_id, design_param, design_param_value, members_variant
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def AssignDesignCommand(
        self,
        design_ref_id: int,
        design_command_name: str,
        design_command_value: str,
        member_ids: list | int,
    ):
        """
        Assign a design command to specified members in a design brief.

        Parameters
        ----------
        design_ref_id : int
            Design brief reference ID.
        design_command_name : str
            Name of the design command.
        design_command_value : str
            Value for the design command.
        member_ids : list of int
            List of member numbers.

        Returns
        -------
        bool
            True if successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> ref_id = staad_obj.Design.CreateDesignBrief(1001)
        >>> result = staad_obj.Design.AssignDesignCommand(ref_id, "CHECK CODE", "", [1, 3])
        """
        if isinstance(member_ids, int):
            member_ids = [member_ids]
        safe_members = make_safe_array_long_input(member_ids)
        members_variant = make_variant_vt_ref(
            safe_members, automation.VT_ARRAY | automation.VT_I4
        )
        retVal = self._design.AssignDesignCommand(
            design_ref_id, design_command_name, design_command_value, members_variant
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def AssignDesignGroup(
        self,
        design_ref_id: int,
        design_group_name: str,
        design_group_value: str,
        same_as_member: int,
        member_ids: list | int,
    ):
        """
        Assign physical members to a design group using a design command.

        Parameters
        ----------
        design_ref_id : int
            Design brief reference ID.
        design_group_name : str
            Name of the design group.
        design_group_value : str
            Value for the design group.
        same_as_member : int
            Reference member for the group.
        member_ids : list of int
            List of member numbers.

        Returns
        -------
        bool
            True if successful, False otherwise.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> ref_id = staad_obj.Design.CreateDesignBrief(1001)
        >>> result = staad_obj.Design.AssignDesignGroup(ref_id, "scSteelGroup", "Ax", 2, [1, 3])
        """
        if isinstance(member_ids, int):
            member_ids = [member_ids]
        safe_members = make_safe_array_long_input(member_ids)
        members_variant = make_variant_vt_ref(
            safe_members, automation.VT_ARRAY | automation.VT_I4
        )
        retVal = self._design.AssignDesignGroup(
            design_ref_id,
            design_group_name,
            design_group_value,
            same_as_member,
            members_variant,
        )
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal == 0

    def GetDesignBriefCode(self, design_ref_id: int):
        """
        Get the design code for a specified design brief.

        Parameters
        ----------
        design_ref_id : int
            Design brief reference ID.

        Returns
        -------
        int
            Design code.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> ref_id = staad_obj.Design.CreateDesignBrief(1001)
        >>> result = staad_obj.Design.GetDesignBriefCode(ref_id)
        """
        return self._design.GetDesignBriefCode(design_ref_id)

    def GetMemberDesignParameters(self, design_ref_id: int, member_no: int):  # noqa: C901
        """
        Get the design parameters for a specified member in a design brief.

        Parameters
        ----------
        design_ref_id : int
            Design brief reference ID.
        member_no : int
            Member number.

        Returns
        -------
        dict
            Dictionary containing:

            status : int
                Return code from COM (0 success, -1 failure or other).
            count : int | None
                Number of parameters available for the member (from COM object's Count property) or None.
            _raw : COM object
                Original COM object (StaadPro.MembSteelDgnParams).
            parameters : dict[str, list]
                Mapping of parameter name -> [value, unit, description, default]. If any component is
                unavailable it is set to None. Parameter names are returned exactly as provided by COM.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> brief_id = 1
        >>> params = staad_obj.Design.GetMemberDesignParameters(brief_id, 1)
        """
        design_params = cc.CreateObject("StaadPro.MembSteelDgnParams")
        status = self._design.GetMemberDesignParameters(
            design_ref_id, member_no, design_params
        )
        if status < 0:
            raise_os_error_if_error_code(status)

        # Attempt to read count (number of parameters for the member)
        try:
            count = getattr(design_params, "Count")
        except Exception:
            count = None

        def _get_attr(name):
            try:
                return getattr(design_params, name)
            except Exception:
                return None

        # Raw COM attributes (may be SAFEARRAY, sequence, or callable indexer)
        raw_names = _get_attr("Name")
        raw_values = _get_attr("Value")
        raw_units = _get_attr("Unit")
        raw_descriptions = _get_attr("Description")
        raw_defaults = _get_attr("Default")

        def _materialize(obj):
            """Turn a COM collection or callable indexer into a Python list."""
            if obj is None:
                return []
            # Avoid splitting a single string into characters
            if isinstance(obj, str):
                return [obj]
            # Already a list/tuple
            if isinstance(obj, (list, tuple)):
                return list(obj)
            # Callable indexer pattern (obj(i)) if count known
            if callable(obj) and count is not None:
                result = []
                for i in range(int(count)):
                    try:
                        result.append(obj(i))
                    except Exception:
                        result.append(None)
                return result
            # Generic sequence protocol
            if hasattr(obj, "__len__") and hasattr(obj, "__getitem__"):
                try:
                    return [obj[i] for i in range(len(obj))]
                except Exception:
                    return []
            # Fallback: wrap scalar
            return [obj]

        names_list = _materialize(raw_names)
        values_list = _materialize(raw_values)
        units_list = _materialize(raw_units)
        descriptions_list = _materialize(raw_descriptions)
        defaults_list = _materialize(raw_defaults)

        parameters = {}
        for i, nm in enumerate(names_list):
            if nm is None:
                continue
            if isinstance(nm, str) and nm.strip() == "":
                continue
            parameters[str(nm)] = [
                values_list[i] if i < len(values_list) else None,
                units_list[i] if i < len(units_list) else None,
                descriptions_list[i] if i < len(descriptions_list) else None,
                defaults_list[i] if i < len(defaults_list) else None,
            ]

        return {
            "status": status,
            "count": count,
            "_raw": design_params,
            "parameters": parameters,
        }
