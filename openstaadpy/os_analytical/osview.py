# ---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
# ---------------------------------------------------------------------------------------------
from __future__ import annotations

from comtypes import CoInitialize, automation

from .openstaadhelper import (
    make_safe_array_double,
    make_safe_array_double_input,
    make_safe_array_int,
    make_safe_array_long,
    make_safe_array_long_input,
    make_variant_vt_ref,
)
from .oserrors import OsErrorBase, raise_os_error_if_error_code
from .osgeometry import OSGeometry


class OSView:
    CoInitialize()

    def __init__(self, staadObj):
        self._staad = staadObj
        self._view = self._staad.View
        self._geometry = OSGeometry(staadObj)

        self._functions = [
            "RefreshView",
            "ShowAllMembers",
            "HideAllMembers",
            "ZoomExtentsMainView",
            "ShowMembers",
            "HideMember",
            "HideMembers",
            "ShowBack",
            "ShowBottom",
            "ShowFront",
            "ShowIsometric",
            "ShowLeft",
            "ShowPlan",
            "ShowRight",
            "SpinLeft",
            "SpinRight",
            "ZoomAll",
            "GetApplicationDesktopSize",
            "SetWindowPosition",
            "RotateUp",
            "RotateDown",
            "RotateLeft",
            "RotateRight",
            "CreateNewViewForSelections",
            "SetLabel",
            "SetSectionView",
            "SetDiagramMode",
            "SetStressType",
            "SetNodeAnnotationMode",
            "SetReactionAnnotationMode",
            "GetInterfaceMode",
            "SetInterfaceMode",
            "SetModeSectionPage",
            "SetBeamAnnotationMode",
            "ShowMember",
            "SetUnits",
            "HidePlate",
            "HideSolid",
            "HideSurface",
            "HideEntity",
            "SelectMembersParallelTo",
            "SelectGroup",
            "SelectInverse",
            "SelectByItemList",
            "SelectByMissingAttribute",
            "SelectEntitiesConnectedToNode",
            "SelectEntitiesConnectedToMember",
            "SelectEntitiesConnectedToPlate",
            "SelectEntitiesConnectedToSolid",
            "GetNoOfBeamsInView",
            "GetBeamsInView",
            "CreateNewViewForSelectionsEx",
            "ExportView",
            "CopyPicture",
            "GetScaleValues",
            "SetScaleValues",
            "GetScaleValueByType",
            "SetScaleValueByType",
            "GetScaleCount",
            "DetachView",
            "RenameView",
            "OpenView",
            "SaveView",
            "GetWindowTitle",
            "GetWindowCount",
            "CloseActiveWindow",
            "SetActiveWindow",
            "SetDesignResults",
        ]

        for function_name in self._functions:
            self._view._FlagAsMethod(function_name)

    def RefreshView(self):
        """
        Refresh the STAAD view window.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.RefreshView()
        """
        self._view.RefreshView()

    def ShowAllMembers(self):
        """
        Show all members in the STAAD view.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ShowAllMembers()
        """
        self._view.ShowAllMembers()

    def HideAllMembers(self):
        """
        Hide all members in the STAAD view.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.HideAllMembers()
        """
        self._view.HideAllMembers()

    def ZoomExtentsMainView(self):
        """
        Zoom to extents in the main STAAD view.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ZoomExtentsMainView()
        """
        self._view.ZoomExtentsMainView()

    def ShowMembers(self, member_list: list | int):
        """
        Show specific members in the STAAD view.

        Parameters
        -----------
        member_list : list of int
            List of member numbers to show.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ShowMembers([1, 2])
        """
        if isinstance(member_list, int):
            member_list = [member_list]
        safe_list = make_safe_array_long_input(member_list)
        lista_variant = make_variant_vt_ref(
            safe_list, automation.VT_ARRAY | automation.VT_I4
        )

        self._view.ShowAllMembers()
        self._view.HideAllMembers()
        self._geometry.ClearMemberSelection()
        self._view.ShowMembers(len(member_list), lista_variant)
        self._view.ShowIsometric()
        self._view.ZoomExtentsMainView()
        self._view.RefreshView()

    def HideMember(self, IDMember):
        """
        Hide a specific member in the STAAD view.

        Parameters
        ----------
        IDMember : int
            Member number to hide.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.HideMember(1)
        """
        self._view.HideMember(IDMember)
        self._view.RefreshView()

    def HideMembers(self, NaMemberNos):
        """
        Hide specific members in the STAAD view.

        Parameters
        ----------
        NaMemberNos : list of int
            List of member numbers to hide.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.HideMembers(2, [1, 2])
        """
        safe_list = make_safe_array_long_input(NaMemberNos)
        lista_variant = make_variant_vt_ref(
            safe_list, automation.VT_ARRAY | automation.VT_I4
        )

        self._view.HideMembers(len(NaMemberNos), lista_variant)
        self._view.RefreshView()

    def ShowBack(self):
        """
        Set the view to the back orientation.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ShowBack()
        """
        self._view.ShowBack()
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def ShowBottom(self):
        """
        Set the view to the bottom orientation.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ShowBottom()
        """
        self._view.ShowBottom()
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def ShowFront(self):
        """
        Set the view to the front orientation.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ShowFront()
        """
        self._view.ShowFront()
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def ShowIsometric(self):
        """
        Set the view to isometric orientation.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ShowIsometric()
        """
        self._view.ShowIsometric()
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def ShowLeft(self):
        """
        Set the view to the left orientation.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ShowLeft()
        """
        self._view.ShowLeft()
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def ShowPlan(self):
        """
        Set the view to the plan (top) orientation.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ShowPlan()
        """
        self._view.ShowPlan()
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def ShowRight(self):
        """
        Set the view to the right orientation.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ShowRight()
        """
        self._view.ShowRight()
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SpinLeft(self, Degrees: float):
        """
        Spin the view to the left by a specified number of degrees.

        Parameters
        ----------
        Degrees : float or int
            Number of degrees to spin left.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SpinLeft(15)
        """
        Degrees = float(Degrees)
        self._view.SpinLeft(Degrees)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SpinRight(self, Degrees: float):
        """
        Spin the view to the right by a specified number of degrees.

        Parameters
        ----------
        Degrees : float or int
            Number of degrees to spin right.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SpinRight(15)
        """
        Degrees = float(Degrees)
        self._view.SpinRight(Degrees)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def ZoomAll(self):
        """
        Zoom to show all objects in the STAAD view.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ZoomAll()
        """
        self._view.ZoomAll()

    def GetApplicationDesktopSize(self):
        """
        Get the size of the application desktop.

        Returns
        -------
        tuple of int
            Tuple of width and height of the application desktop.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> width, height = staad_obj.View.GetApplicationDesktopSize()
        """
        safe_n1 = make_safe_array_int(1)
        L = make_variant_vt_ref(safe_n1, automation.VT_I4)

        safe_n2 = make_safe_array_int(1)
        W = make_variant_vt_ref(safe_n2, automation.VT_I4)

        retVal = self._view.GetApplicationDesktopSize(L, W)
        if not retVal:
            raise_os_error_if_error_code(-1)
        elif retVal < 0:
            raise_os_error_if_error_code(retVal)

        L = int(L[0])
        W = int(W[0])

        return (L, W)

    def SetWindowPosition(self, xTop, yTop, xWindow, yWindow):
        """
        Set the position and size of the STAAD application window.

        Parameters
        ----------
        xTop : int
            X coordinate of the top-left corner.
        yTop : int
            Y coordinate of the top-left corner.
        xWindow : int
            Width of the window.
        yWindow : int
            Height of the window.

        Returns
        -------
        bool
            True if successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SetWindowPosition(100, 100, 800, 600)
        """
        retVal = self._view.SetWindowPosition(xTop, yTop, xWindow, yWindow)
        if not retVal:
            raise_os_error_if_error_code(-1)
        elif retVal < 0:
            raise_os_error_if_error_code(retVal)
        return bool(retVal)

    def RotateUp(self, dDegrees: float):
        """
        Rotates the structure through Degrees about the Global X-Axis.

        Parameters
        ----------
        dDegrees : float
            Variable providing the degree of rotation.

        Examples
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.RotateUp(30)
        """
        self._view.RotateUp(dDegrees)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def RotateDown(self, dDegrees: float):
        """
        Rotates the structure through Degrees about the Global X-Axis.

        Parameters
        ----------
        dDegrees : float
            Variable providing the degree of rotation.

        Examples
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.RotateDown(30)
        """
        self._view.RotateDown(dDegrees)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def RotateLeft(self, dDegrees: float):
        """
        Rotates the structure through Degrees about the Global Y-Axis.

        Parameters
        ----------
        dDegrees : float
            Variable providing the degree of rotation.

        Examples
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.RotateLeft(30)
        """
        self._view.RotateLeft(dDegrees)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def RotateRight(self, dDegrees: float):
        """
        Rotates the structure through Degrees about the Global Y-Axis.

        Parameters
        ----------
        dDegrees : float
            Variable providing the degree of rotation.

        Examples
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.RotateRight(30)
        """
        self._view.RotateRight(dDegrees)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def CreateNewViewForSelections(self):
        """
        Creates a new view in new window for the selected objects displayed in the active window.

        Examples
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.CreateNewViewForSelections()
        """
        self._view.CreateNewViewForSelections()

    def SetLabel(self, which: int, showFlag: bool):
        """
        Sets the label on the structure diagram on or off.

        Parameters
        ----------
        which : int
            Variable identifying the diagram type. It may be one of the following values:
                +------+--------------------------------------------+
                | ID   | Label Type                                 |
                +======+============================================+
                | 0    | Node number label                          |
                +------+--------------------------------------------+
                | 1    | Member number label                        |
                +------+--------------------------------------------+
                | 2    | Member property reference label            |
                +------+--------------------------------------------+
                | 3    | Material property reference label          |
                +------+--------------------------------------------+
                | 4    | Support label                              |
                +------+--------------------------------------------+
                | 5    | Member release label                       |
                +------+--------------------------------------------+
                | 6    | Member orientation label                   |
                +------+--------------------------------------------+
                | 7    | Member section label                       |
                +------+--------------------------------------------+
                | 8    | Load value label                           |
                +------+--------------------------------------------+
                | 9    | Axes label                                 |
                +------+--------------------------------------------+
                | 10   | Node position label                        |
                +------+--------------------------------------------+
                | 11   | Member specification label                 |
                +------+--------------------------------------------+
                | 12   | Member ends                                |
                +------+--------------------------------------------+
                | 13   | Plate element number label                 |
                +------+--------------------------------------------+
                | 14   | Plate element orientation label            |
                +------+--------------------------------------------+
                | 15   | Solid element number label                 |
                +------+--------------------------------------------+
                | 16   | Dimension label                            |
                +------+--------------------------------------------+
                | 17   | Floor load label                           |
                +------+--------------------------------------------+
                | 18   | Floor load distribution diagram label      |
                +------+--------------------------------------------+
                | 19   | Wind load label                            |
                +------+--------------------------------------------+
                | 20   | Wind load influence area diagram label     |
                +------+--------------------------------------------+
                | 21   | Diagram Info                               |
                +------+--------------------------------------------+
        showFlag : bool
            Variable to set label mode on (True) or off (False).


        Examples
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SetLabel(20, True)
        """
        self._view.SetLabel(which, showFlag)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SetSectionView(self, plane: int, minVal: float, maxVal: float):
        """
        Creates a section view of the structure.

        Parameters
        ----------
        plane : int
            Variable identifying the section plane.  It may be one of the following values:
                +-----+------------------+
                | ID  | Values for plane |
                +=====+==================+
                | 0   | XY Plane         |
                +-----+------------------+
                | 1   | YZ Plane         |
                +-----+------------------+
                | 2   | XZ Plane         |
                +-----+------------------+
        minVal : float
            Minimum range of the cutting plane.
        maxVal : float
            Maximum range of the cutting plane.

        Examples
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SetSectionView(1, 0.4, 0.6)
        """
        self._view.SetSectionView(plane, minVal, maxVal)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SetDiagramMode(self, which: int, showFlag: bool, refreshFlag: bool):
        """
        Sets the label on the structure diagram on or off.

        Parameters
        ----------
        which : int
            Variable identifying the diagram type. It may be one of the following values

            +-----+--------------------------------------------------------+
            | ID  | Diagram Type                                           |
            +=====+========================================================+
            | 0   | Load                                                   |
            +-----+--------------------------------------------------------+
            | 1   | Displacement                                           |
            +-----+--------------------------------------------------------+
            | 2   | MY                                                     |
            +-----+--------------------------------------------------------+
            | 3   | MZ                                                     |
            +-----+--------------------------------------------------------+
            | 4   | FY                                                     |
            +-----+--------------------------------------------------------+
            | 5   | FZ                                                     |
            +-----+--------------------------------------------------------+
            | 6   | AX                                                     |
            +-----+--------------------------------------------------------+
            | 7   | TR                                                     |
            +-----+--------------------------------------------------------+
            | 8   | Structure                                              |
            +-----+--------------------------------------------------------+
            | 9   | Full Section                                           |
            +-----+--------------------------------------------------------+
            | 10  | Section Outline                                        |
            +-----+--------------------------------------------------------+
            | 11  | Stress                                                 |
            +-----+--------------------------------------------------------+
            | 12  | Shrink                                                 |
            +-----+--------------------------------------------------------+
            | 13  | Perspective                                            |
            +-----+--------------------------------------------------------+
            | 14  | Hide Structure                                         |
            +-----+--------------------------------------------------------+
            | 15  | Fill Plates & Solids                                   |
            +-----+--------------------------------------------------------+
            | 16  | Hide Plates & Solids                                   |
            +-----+--------------------------------------------------------+
            | 17  | Hide Piping                                            |
            +-----+--------------------------------------------------------+
            | 18  | Sort Geometry                                          |
            +-----+--------------------------------------------------------+
            | 19  | Sort Nodes                                             |
            +-----+--------------------------------------------------------+
            | 20  | Plate Stress                                           |
            +-----+--------------------------------------------------------+
            | 21  | Solid Stress                                           |
            +-----+--------------------------------------------------------+
            | 22  | Mode Shape                                             |
            +-----+--------------------------------------------------------+
            | 23  | Stress Animation                                       |
            +-----+--------------------------------------------------------+
            | 24  | Plate reinforcement                                    |
            +-----+--------------------------------------------------------+
            | 25  | Deck Influence Diagram*                                |
            +-----+--------------------------------------------------------+
            | 26  | Deck Carriageways*                                     |
            +-----+--------------------------------------------------------+
            | 27  | Deck Triangulation*                                    |
            +-----+--------------------------------------------------------+
            | 28  | Deck Loads*                                            |
            +-----+--------------------------------------------------------+
            | 29  | Deck Vehicles*                                         |
            +-----+--------------------------------------------------------+
            |     | (*) Requires the STAAD.beava component                 |
            +-----+--------------------------------------------------------+

        showFlag : bool
            Variable to set label mode on (True) or off (False).
        refreshFlag : bool
            Variable (True or False). If True, STAAD.Pro viewing windows refresh.

        Notes
        -----
        For Plate Stress/Solid Stress (which = 20/21), a stress type must already be set via
        :meth:`SetStressType` before calling this with showFlag = True, otherwise the diagram is
        not turned on (matches the Diagrams>Plate/Solid Stress dialog, which also requires a type
        to be picked before the contour can be shown).

        Examples
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SetDiagramMode(1, True, True)
        """
        self._view.SetDiagramMode(which, showFlag, refreshFlag)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SetStressType(self, entityType: int, stressType: int, refreshFlag: bool):
        """
        Sets the plate or solid stress type used for the contour diagram, as shown in the
        Diagrams>Plate/Solid Stress dialog, and recomputes its legend range for the active load.
        Call this before :meth:`SetDiagramMode` so the contour has a type to display when the
        diagram is turned on; it can also be called afterwards to change the type of an
        already-shown diagram.

        Parameters
        ----------
        entityType : int
            Variable identifying which stress diagram to set. It may be one of the following values:
                +-----+--------------+
                | ID  | Entity Type  |
                +=====+==============+
                | 20  | Plate Stress |
                +-----+--------------+
                | 21  | Solid Stress |
                +-----+--------------+

        stressType : int
            Variable identifying the stress type.

            When entityType = 20 (Plate Stress), it may be one of the following values:
                +-----+------------------------------+
                | ID  | Plate Stress Type            |
                +=====+==============================+
                | 1   | Max Absolute                 |
                +-----+------------------------------+
                | 2   | Top Max                      |
                +-----+------------------------------+
                | 3   | Top Min                      |
                +-----+------------------------------+
                | 4   | Top Max Shear                |
                +-----+------------------------------+
                | 5   | Bottom Max                   |
                +-----+------------------------------+
                | 6   | Bottom Min                   |
                +-----+------------------------------+
                | 7   | Bottom Max Shear             |
                +-----+------------------------------+
                | 8   | Max Von Mises                |
                +-----+------------------------------+
                | 9   | Von Mises Top Max            |
                +-----+------------------------------+
                | 10  | Von Mises Bottom Max         |
                +-----+------------------------------+
                | 11  | Max Tresca                   |
                +-----+------------------------------+
                | 12  | Top Tresca                   |
                +-----+------------------------------+
                | 13  | Bottom Tresca                |
                +-----+------------------------------+
                | 14  | FX                           |
                +-----+------------------------------+
                | 15  | FY                           |
                +-----+------------------------------+
                | 16  | FXY                          |
                +-----+------------------------------+
                | 17  | MX                           |
                +-----+------------------------------+
                | 18  | MY                           |
                +-----+------------------------------+
                | 19  | MZ                           |
                +-----+------------------------------+
                | 20  | QX                           |
                +-----+------------------------------+
                | 21  | QY                           |
                +-----+------------------------------+
                | 22  | Global                       |
                +-----+------------------------------+
                | 23  | Global Membrane Stresses     |
                +-----+------------------------------+
                | 24  | Global Shear Stresses        |
                +-----+------------------------------+
                | 25  | Base Pressure                |
                +-----+------------------------------+
                | 26  | Combined X Top               |
                +-----+------------------------------+
                | 27  | Combined Y Top               |
                +-----+------------------------------+
                | 28  | Combined XY Top              |
                +-----+------------------------------+
                | 29  | Combined X Bottom            |
                +-----+------------------------------+
                | 30  | Combined Y Bottom            |
                +-----+------------------------------+
                | 31  | Combined XY Bottom           |
                +-----+------------------------------+

            When entityType = 21 (Solid Stress), it may be one of the following values:
                +-----+------------------------------+
                | ID  | Solid Stress Type            |
                +=====+==============================+
                | 1   | SXX                          |
                +-----+------------------------------+
                | 2   | SYY                          |
                +-----+------------------------------+
                | 3   | SZZ                          |
                +-----+------------------------------+
                | 4   | SXY                          |
                +-----+------------------------------+
                | 5   | SYZ                          |
                +-----+------------------------------+
                | 6   | SXZ                          |
                +-----+------------------------------+
                | 7   | S11                          |
                +-----+------------------------------+
                | 8   | S22                          |
                +-----+------------------------------+
                | 9   | S33                          |
                +-----+------------------------------+
                | 10  | Sigma Effective (Von Mises)  |
                +-----+------------------------------+

        refreshFlag : bool
            Variable (True or False). If True, STAAD.Pro viewing windows refresh.

        Returns
        -------
        bool
            True if successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SetStressType(20, 8, True)
        """
        retVal = self._view.SetStressType(entityType, stressType, refreshFlag)
        if not retVal:
            raise_os_error_if_error_code(-1)
        elif retVal < 0:
            raise_os_error_if_error_code(retVal)
        return bool(retVal)

    def SetNodeAnnotationMode(self, dFlag: bool, refreshFlag: bool):
        """
        Sets the node displacement annotation mode.  This function works only in the post-processing mode of STAAD.Pro.

        Parameters
        ----------
        dFlag : bool
            Variable controlling the annotation type. It may be one of the following values:
                +-----+--------------------------+
                | ID  | Annotation Type          |
                +=====+==========================+
                | 1   | X Displacement           |
                +-----+--------------------------+
                | 2   | Y Displacement           |
                +-----+--------------------------+
                | 3   | Z Displacement           |
                +-----+--------------------------+
                | 4   | Resultant Displacement   |
                +-----+--------------------------+
        refreshFlag : bool
            Variable (True or False). If True, STAAD.Pro viewing windows refresh with the current annotation mode.

        Examples
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SetNodeAnnotationMode(1, True)
        """
        self._view.SetNodeAnnotationMode(dFlag, refreshFlag)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SetReactionAnnotationMode(self, dFlag: bool, refreshFlag: bool):
        """
        Sets the node displacement annotation mode.  This function works only in the post-processing mode of STAAD.Pro.

        Parameters
        ----------
        dFlag : bool
            Variable controlling the annotation type. It may be one of the following values:
                +-----+-----------------------+
                | ID  | Annotation Type       |
                +=====+=======================+
                | 1   | X Reaction            |
                +-----+-----------------------+
                | 2   | Y Reaction            |
                +-----+-----------------------+
                | 3   | Z Reaction            |
                +-----+-----------------------+
                | 4   | X Rotation            |
                +-----+-----------------------+
                | 5   | Y Rotation            |
                +-----+-----------------------+
                | 6   | Z Rotation            |
                +-----+-----------------------+
                | 7   | Reaction Value Only   |
                +-----+-----------------------+
        refreshFlag : bool
            Variable (True or False). If True, STAAD.Pro viewing windows refresh with the current annotation mode.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SetReactionAnnotationMode(1, True)
        """
        self._view.SetReactionAnnotationMode(dFlag, refreshFlag)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def GetInterfaceMode(self):
        """
        Gets the interface mode type.
        Note:
            This function returns the current visual mode in the STAAD.Pro environment.

        Returns
        -------
        int
            Returns 0 if Pre-processor or modeling mode.
            Returns 1 if Post-processing mode.
            Returns 2 if Interactive design mode for STAAD.etc interoperability.
            Returns 4 if Piping mode.
            Returns 5 if BEAVA (i.e., Bridge Deck ) mode.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> mode = staad_obj.View.GetInterfaceMode()
        """
        return self._view.GetInterfaceMode()

    def SetInterfaceMode(self, interfaceMode: int):
        """
        Sets the interface mode type.
        Note:
            This function sets the current visual mode in the STAAD.Pro environment.

        Parameters
        ----------
        interfaceMode : int
            Variable to set the current visual mode in STAAD.Pro environment. Followings are the valid values for mode:
                +-----+--------------------------------+
                |ID   | Mode Type                      |
                +=====+================================+
                | 0   | Pre-processor or modeling mode |
                +-----+--------------------------------+
                | 1   | Physical modeling mode         |
                +-----+--------------------------------+
                | 2   | Building planner mode          |
                +-----+--------------------------------+
                | 3   | Piping mode                    |
                +-----+--------------------------------+
                | 5   | Post Processing mode           |
                +-----+--------------------------------+
                | 6   | FoundationDesign mode          |
                +-----+--------------------------------+
                | 7   | ConnectionDesign mode          |
                +-----+--------------------------------+
                | 9   | AdvancedConcreteDesign mode    |
                +-----+--------------------------------+
                | 10  | AdvancedSlabDesign mode        |
                +-----+--------------------------------+
                | 11  | Earthquake mode                |
                +-----+--------------------------------+
                | 12  | SteelAutoDrafter mode          |
                +-----+--------------------------------+
                | 13  | ChineseSteelDesign mode        |
                +-----+--------------------------------+

        Returns
        -------
        int
            True if successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> status = staad_obj.View.SetInterfaceMode(1)
        """
        retVal = self._view.SetInterfaceMode(interfaceMode)
        if not retVal:
            raise_os_error_if_error_code(-1)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()
        return retVal

    def SetModeSectionPage(
        self, interfaceMode: int, sectionNumber: int, pageNumber: int
    ):
        """
        This function sets the current page mode in the STAAD.Pro environment.

        Parameters
        ----------
        interfaceMode : int
            Variable to set the current visual mode in STAAD.Pro environment. Followings are the valid values for mode:
                +----+-----------------------------------------------------------+
                | ID | Interface Mode                                            |
                +====+===========================================================+
                | 0  | Pre-processor or modeling mode                            |
                +----+-----------------------------------------------------------+
                | 1  | Post-processing mode                                      |
                +----+-----------------------------------------------------------+
                | 2  | Interactive design mode for STAAD.etc interoperability    |
                +----+-----------------------------------------------------------+
                | 4  | Piping mode                                               |
                +----+-----------------------------------------------------------+
                | 5  | BEAVA (i.e., Bridge Deck) mode                            |
                +----+-----------------------------------------------------------+
        sectionNumber : int
            Variable to set the current main page (the tabs on the left-hand side of the screen) in the STAAD.Pro environment.  The following are valid values for section:
                +----+----------------------+
                | ID | Main Page            |
                +====+======================+
                | 1  | Setup page           |
                +----+----------------------+
                | 2  | Geometry page        |
                +----+----------------------+
                | 3  | General page         |
                +----+----------------------+
                | 5  | Node Results page    |
                +----+----------------------+
                | 6  | Beam Result page     |
                +----+----------------------+
                | 7  | Plate Results page   |
                +----+----------------------+
                | 8  | Solid Results page   |
                +----+----------------------+
        pageNumber : int
            Variable  to set the current sub page (within a particular main page - the tabs on the left-hand side of the screen) in the STAAD.Pro environment. The following are valid values for modeSubPage:
                +----+------------------------------+
                | ID | Page Number                  |
                +====+==============================+
                | 0  | Job Info page                |
                +----+------------------------------+
                | 1  | Beam page                    |
                +----+------------------------------+
                | 4  | Plate page                   |
                +----+------------------------------+
                | 5  | Solid page                   |
                +----+------------------------------+
                | 6  | Property page                |
                +----+------------------------------+
                | 7  | Constant page                |
                +----+------------------------------+
                | 8  | Material page                |
                +----+------------------------------+
                | 9  | Support page                 |
                +----+------------------------------+
                | 10 | Member Specifications page   |
                +----+------------------------------+
                | 11 | Load page                    |
                +----+------------------------------+
                | 17 | Reaction page                |
                +----+------------------------------+
                | 18 | Displacement page            |
                +----+------------------------------+
                | 19 | Failure page                 |
                +----+------------------------------+
                | 20 | Forces page                  |
                +----+------------------------------+
                | 21 | Beam Stress page             |
                +----+------------------------------+
                | 22 | Plate Stress page            |
                +----+------------------------------+
                | 23 | Solid Stress page            |
                +----+------------------------------+


        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SetModeSectionPage(1,6,20)
        """
        self._view.SetModeSectionPage(interfaceMode, sectionNumber, pageNumber)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SetBeamAnnotationMode(self, Type: int, DWFlags: int, RefreshFlag: bool):
        """
        This function sets the current page mode in the STAAD.Pro environment.

        Parameters
        ----------
        Type : int
            Variable controlling the annotation type.  It may be one of the following values:
                +----+----------------------+
                | ID | Annotation Type      |
                +====+======================+
                | 0  | Axial Diagram        |
                +----+----------------------+
                | 1  | Torsion Diagram      |
                +----+----------------------+
                | 2  | Moment Diagram       |
                +----+----------------------+
                | 3  | Shear Diagram        |
                +----+----------------------+
                | 4  | Stress Diagram       |
                +----+----------------------+
                | 5  | Displacement Diagram |
                +----+----------------------+
        DWFlags : int
            Variable controlling what values are to be shown for the annotationType.  It may be one of the following values:
                +----+-----------------------+
                | ID | Values                |
                +====+=======================+
                | 1  | End Values            |
                +----+-----------------------+
                | 2  | Max Absolute Values   |
                +----+-----------------------+
                | 3  | Mid-span Values       |
                +----+-----------------------+
        RefreshFlag : int
            Boolean variable (True or False). If True, STAAD.Pro viewing windows refresh with the current annotation mode.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SetBeamAnnotationMode(2, 1, True)
        """
        self._view.SetBeamAnnotationMode(Type, DWFlags, RefreshFlag)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def ShowMember(self, nMember: int):
        """
        Show the specified member.

        Parameters
        ----------
        nMember : int
            Variable that holds member number to be shown.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.ShowMember(5)
        """
        self._view.ShowMember(nMember)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SetUnits(self, uType: int, strUnit: str):
        """
        Set viewing unit for the active view.

        Parameters
        ----------
        uType : int
            Variable that holds unit type. Values are as follows:
                +-----+-----------------------------+
                | ID  | Unit Type                   |
                +=====+=============================+
                | 0   |Dimension                    |
                +-----+-----------------------------+
                | 1   |Displacement                 |
                +-----+-----------------------------+
                | 2   |SectionDimension             |
                +-----+-----------------------------+
                | 3   |SectionArea                  |
                +-----+-----------------------------+
                | 4   |Inertia                      |
                +-----+-----------------------------+
                | 5   |Force                        |
                +-----+-----------------------------+
                | 6   |Moment                       |
                +-----+-----------------------------+
                | 7   |DistributedForce             |
                +-----+-----------------------------+
                | 8   |DistributedMoment            |
                +-----+-----------------------------+
                | 9   |Density                      |
                +-----+-----------------------------+
                | 10  |Acceleration                 |
                +-----+-----------------------------+
                | 11  |Spring                       |
                +-----+-----------------------------+
                | 12  |RotSpring                    |
                +-----+-----------------------------+
                | 13  |MaterialModulus              |
                +-----+-----------------------------+
                | 14  |Stress                       |
                +-----+-----------------------------+
                | 15  |Alpha                        |
                +-----+-----------------------------+
                | 16  |Temperature                  |
                +-----+-----------------------------+
                | 17  |Mass                         |
                +-----+-----------------------------+
                | 18  |SectionModulus               |
                +-----+-----------------------------+
                | 19  |RotationalDisplacement       |
                +-----+-----------------------------+
                | 20  |SubgradeModulus              |
                +-----+-----------------------------+
                | -1  |NoUnit                       |
                +-----+-----------------------------+
        strUnit : str
            Variable array that holds the unit for the specified type. Like "cm", "kns", "feet", "kn/cm" etc.


        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SetUnits(0, "cm")
        """
        self._view.SetUnits(uType, strUnit)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def HidePlate(self, nPlate: int):
        """
        Hide the specified plate.

        Parameters
        ----------
        nPlate : int
            Variable that holds plate number to be hidden.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.HidePlate(5)
        """
        self._view.HidePlate(nPlate)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def HideSolid(self, nSolid: int):
        """
        Hide the specified solid.

        Parameters
        ----------
        nSolid : int
            Variable that holds solid number to be hidden.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.HideSolid(5)
        """
        self._view.HideSolid(nSolid)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def HideSurface(self, nSurface: int):
        """
        Hide the specified surface.

        Parameters
        ----------
        nSurface : int
            Variable that holds surface number to be hidden.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.HideSurface(5)
        """
        self._view.HideSurface(nSurface)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def HideEntity(self, nEntity: int):
        """
        Hides the specified entity, which may be a Beam, Plate, Solid, or Surface.

        Parameters
        ----------
        nEntity : int
            Variable that holds an entity (i.e., Member, Plates etc.) number to be hidden.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.HideEntity(5)
        """
        self._view.HideEntity(nEntity)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SelectMembersParallelTo(self, bstrAxis: str):
        """
        Select members parallel to the specified axis.

        Parameters
        ----------
        bstrAxis : str
            Variable that holds the Axis ID. It may have three values:
                +----+----------+
                | ID | Axis     |
                +====+==========+
                | X  | X-Axis   |
                +----+----------+
                | Y  | Y-Axis   |
                +----+----------+
                | Z  | Z-Axis   |
                +----+----------+

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SelectMembersParallelTo("X")
        """
        self._view.SelectMembersParallelTo(bstrAxis)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SelectGroup(self, bstrGroup: str):
        """
        Select the relevant entities of the specified group.

        Parameters
        ----------
        bstrGroup : str
            A string variable that holds the group name.

        Returns
        -------
        int
            Returns True if successful

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SelectGroup("grp_name")
        """
        retVal = self._view.SelectGroup(bstrGroup)
        if not retVal:
            raise_os_error_if_error_code(-1)
        return retVal

    def SelectInverse(self, entityType: int):
        """
        Inverse geometry selection for the specified entity.

        Parameters
        ----------
        entityType : int
            Variable that holds entity type. Values may be as follows:
                +-----+----------------+
                | ID  | Entity Type    |
                +=====+================+
                | 1   | Node           |
                +-----+----------------+
                | 2   | Beam           |
                +-----+----------------+
                | 3   | Plate          |
                +-----+----------------+
                | 4   | Solid          |
                +-----+----------------+
                | 5   | Surface        |
                +-----+----------------+

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SelectInverse(1)
        """
        self._view.SelectInverse(entityType)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SelectByItemList(self, entityType: int, itemList: list):
        """
        Select entities as specified.

        Parameters
        ----------
        entityType : int
            Variable that holds entity type. Values may be as follows:
                +-----+----------------+
                | ID  | Entity Type    |
                +=====+================+
                | 1   | Node           |
                +-----+----------------+
                | 2   | Beam           |
                +-----+----------------+
                | 3   | Plate          |
                +-----+----------------+
                | 4   | Solid          |
                +-----+----------------+
                | 5   | Surface        |
                +-----+----------------+
        itemList : list of int
            List holds the entity nos, which need to be selected.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SelectByItemList(1, 2, [1, 2])
        """
        entityList = make_safe_array_long_input(itemList)
        entity_list_vt = make_variant_vt_ref(
            entityList, automation.VT_ARRAY | automation.VT_I4
        )
        self._view.SelectByItemList(entityType, len(itemList), entity_list_vt)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SelectByMissingAttribute(self, attributeCode: int):
        """
        Select entity list for which specified entity is missing.

        Parameters
        ----------
        attributeCode : int
            Variable that holds attribute type. Values may be as follows:
                +----+--------------------------------------------------------+
                | ID | Entity Type                                            |
                +====+========================================================+
                | 1  | Missing Property                                       |
                +----+--------------------------------------------------------+
                | 2  | Missing Modulus of Elasticity                          |
                +----+--------------------------------------------------------+
                | 3  | Missing Density of Material                            |
                +----+--------------------------------------------------------+
                | 4  | Missing Alpha (Coefficient of Thermal Expansion)       |
                +----+--------------------------------------------------------+
                | 5  | Missing Poisson Ratio                                  |
                +----+--------------------------------------------------------+

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SelectByMissingAttribute(5)
        """
        self._view.SelectByMissingAttribute(attributeCode)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SelectEntitiesConnectedToNode(self, entityType: int, nodeNo: int):
        """
        Select entities as specified in type and connected with the specified node.

        Parameters
        ----------
        entityType : int
            Variable that holds entity type. Values may be as follows:
                +-----+----------------+
                | ID  | Entity Type    |
                +=====+================+
                | 0   | Geometry       |
                +-----+----------------+
                | 1   | Beam           |
                +-----+----------------+
                | 2   | Plate          |
                +-----+----------------+
                | 3   | Solid          |
                +-----+----------------+
        nodeNo : int
            Variable that holds node numbers with which connected entities needs to be selected.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SelectEntitiesConnectedToNode(0, 1)
        """
        self._view.SelectEntitiesConnectedToNode(entityType, nodeNo)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SelectEntitiesConnectedToMember(self, entityType: int, memberNo: int):
        """
        Select entities as specified in type and connected with the specified Member.

        Parameters
        ----------
        entityType : int
            Variable that holds entity type. Values may be as follows:
                +-----+----------------+
                | ID  | Entity Type    |
                +=====+================+
                | 0   | Geometry       |
                +-----+----------------+
                | 1   | Beam           |
                +-----+----------------+
                | 2   | Plate          |
                +-----+----------------+
                | 3   | Solid          |
                +-----+----------------+
        memberNo : int
            Variable that holds Member numbers with which connected entities needs to be selected.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SelectEntitiesConnectedToMember(1, 2)
        """
        self._view.SelectEntitiesConnectedToMember(entityType, memberNo)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SelectEntitiesConnectedToPlate(self, entityType: int, plateNo: int):
        """
        Select entities as specified in type and connected with the specified Plate.

        Parameters
        ----------
        entityType : int
            Variable that holds entity type. Values may be as follows:
                +-----+----------------+
                | ID  | Entity Type    |
                +=====+================+
                | 0   | Geometry       |
                +-----+----------------+
                | 1   | Beam           |
                +-----+----------------+
                | 2   | Plate          |
                +-----+----------------+
                | 3   | Solid          |
                +-----+----------------+
        plateNo : int
            Variable that holds Plate numbers with which connected entities needs to be selected.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SelectEntitiesConnectedToPlate(2, 3)
        """
        self._view.SelectEntitiesConnectedToPlate(entityType, plateNo)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def SelectEntitiesConnectedToSolid(self, entityType: int, solidNo: int):
        """
        Select entities as specified in type and connected with the specified Solid.

        Parameters
        ----------
        entityType : int
            Variable that holds entity type. Values may be as follows:
                +-----+----------------+
                | ID  | Entity Type    |
                +=====+================+
                | 0   | Geometry       |
                +-----+----------------+
                | 1   | Beam           |
                +-----+----------------+
                | 2   | Plate          |
                +-----+----------------+
                | 3   | Solid          |
                +-----+----------------+
        solidNo : int
            Variable that holds Solid numbers with which connected entities needs to be selected.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.SelectEntitiesConnectedToSolid(3, 4)
        """
        self._view.SelectEntitiesConnectedToSolid(entityType, solidNo)
        self._view.RefreshView()
        self._view.ZoomExtentsMainView()

    def GetNoOfBeamsInView(self):
        """
        Get No Of Beams In View

        Returns
        -------
        int
            Returns number of beams present in view.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> count = staad_obj.View.GetNoOfBeamsInView()
        """
        return self._view.GetNoOfBeamsInView()

    def GetBeamsInView(self, nBeamList: list):
        """
        Get Beams In View

        Parameters
        ----------
        nBeamList : nBeamList
            Collection of beam

        Returns
        -------
        int
            Returns number of beams present in view.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.View.GetBeamsInView([1, 2, 4])
        """
        nBeamList_ref = make_safe_array_long_input(nBeamList)
        nBeamList_vt = make_variant_vt_ref(
            nBeamList_ref, automation.VT_ARRAY | automation.VT_I4
        )
        return self._view.GetBeamsInView(nBeamList_vt)

    def CreateNewViewForSelectionsEx(self, windowOptions: int):
        """
        Creates a new view for the selected objects displayed in the active window based on specified options.

        Parameters
        ----------
        windowOptions : int
            0 = Creates a new window for the view, 1 = Display the view in the active window (type - Long).

        Returns
        -------
        bool
            True Creation of new view is successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.View.CreateNewViewForSelectionsEx(1)
        """
        retVal = self._view.CreateNewViewForSelectionsEx(windowOptions)
        if not retVal:
            raise_os_error_if_error_code(-1)
        return retVal

    def ExportView(
        self, FileLocation: str, FileName: str, FileFormat: int, Overwrite: bool
    ):
        """
        Used for exporting the information displayed in the active view window into a standard, graphical image format (e.g., Bitmap, JPEG, TIFF et.).

        Parameters
        ----------
        FileLocation : str
            Location of the saved view file (Folder need to be present otherwise it will return -100) .
        FileName : str
            Name of the saved view file.
        FileFormat : int
            0 = bmp, 1 = jpg, 2 = tga, 3 = tif - Create the view in the specific format.
        Overwrite : bool
            Boolean for provide option to  overwrite an existing file.
                - True - Allow Overwrite
                - False - No overwrite

        Returns
        -------
        int
            1 if Export view is successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.View.ExportView(r"<folderPath>", "<fileName>", 1, True)
        """
        retVal = self._view.ExportView(FileLocation, FileName, FileFormat, Overwrite)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def CopyPicture(self):
        """
        Copy active view to clipbord and gives size of image in it's referance variables. It does not copy table views but works with Structure View, Graphs, 3d Rendered view.

        Returns
        -------
        Tuple
            Returns a tuple containing xDim size of image in x direction (Length in Pixel) and yDim size of image in y direction (width in Pixel) respectively.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.View.CopyPicture()
        """
        safe_xDim = make_safe_array_long(0)
        xDim = make_variant_vt_ref(safe_xDim, automation.VT_I4)
        safe_yDim = make_safe_array_long(0)
        yDim = make_variant_vt_ref(safe_yDim, automation.VT_I4)

        retVal = self._view.CopyPicture(xDim, yDim)
        if not retVal:
            raise_os_error_if_error_code(-1)
        return (xDim[0], yDim[0])

    def GetScaleValues(self):
        """
        Obtain the current set of scales used for displaying loads and results shown in the Diagrams > Scales dialog.

        Returns
        -------
        list of float
            Returns list of float type and size same as number of scale types. API gets all the values in this list. Scale type and ID in array will be as following:
                +----+---------------+-----------------+---------------------+
                | ID | Type          | Scale Items     | Unit per length     |
                +====+===============+=================+=====================+
                | 0  | Loads         | Point Force     | Force               |
                +----+---------------+-----------------+---------------------+
                | 1  | Loads         | Dist. Force     | Force/length        |
                +----+---------------+-----------------+---------------------+
                | 2  | Loads         | Point Moment    | Force*length        |
                +----+---------------+-----------------+---------------------+
                | 3  | Loads         | Dist. Moment    | Force*length/length |
                +----+---------------+-----------------+---------------------+
                | 4  | Loads         | Pressure        | Force/length^2      |
                +----+---------------+-----------------+---------------------+
                | 5  | Results       | Bending Y       | Force*length        |
                +----+---------------+-----------------+---------------------+
                | 6  | Results       | Bending Z       | Force*length        |
                +----+---------------+-----------------+---------------------+
                | 7  | Results       | Shear Y         | Force               |
                +----+---------------+-----------------+---------------------+
                | 8  | Results       | Shear Z         | Force               |
                +----+---------------+-----------------+---------------------+
                | 9  | Results       | Axial           | Force               |
                +----+---------------+-----------------+---------------------+
                | 10 | Results       | Torsion         | Force*length        |
                +----+---------------+-----------------+---------------------+
                | 11 | Results       | Displacement    | Length              |
                +----+---------------+-----------------+---------------------+
                | 12 | Results       | Beam Stress     | Force/length^2      |
                +----+---------------+-----------------+---------------------+
                | 13 | Results       | Mode Shape      | (none)              |
                +----+---------------+-----------------+---------------------+

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> scales = staad_obj.View.GetScaleValues()
        >>> print(list)
        """
        scaleCount = self._view.GetScaleCount()
        scale_safe_list = make_safe_array_double(scaleCount)
        scaleList = make_variant_vt_ref(
            scale_safe_list, automation.VT_ARRAY | automation.VT_R8
        )
        retVal = self._view.GetScaleValues(scaleList)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return scaleList[0]

    def SetScaleValues(self, ScalesList: list):
        """
        Set the scales used for displaying loads and results as shown in the Diagrams>Scales dialog.

        Parameters
        ----------
        ScalesList : list of float
            List of float type and size same as number of scale types. API sets the values in this list to scale types

        Returns
        -------
        int
            Returns 1 if Values were successfully updated.
            Returns 0 if Values could not be updated.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> result = staad_obj.View.SetScaleValues([5, 8])
        """
        ScalesList_safe = make_safe_array_double_input(ScalesList)
        scalesList_vt = make_variant_vt_ref(
            ScalesList_safe, automation.VT_ARRAY | automation.VT_R8
        )
        retVal = self._view.SetScaleValues(scalesList_vt)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        return retVal

    def GetScaleValueByType(self, scaleTypeId: int):
        """
        Obtain the value of the scale that is used to display a specified load or result diagram as defined in the Diagrams>Scales dialog depending upon the scale ID passed.

        Parameters
        ----------
        scaleTypeId : int
            The index of the required load or result type

        Returns
        -------
        float
            Returns value of scale type listed, in Base Units

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> value = staad_obj.View.GetScaleValueByType(1)
        >>> print(value)
        """
        safe_value = make_safe_array_double(0)
        value = make_variant_vt_ref(safe_value, automation.VT_R8)

        retVal = self._view.GetScaleValueByType(scaleTypeId, value)
        if not retVal:
            raise_os_error_if_error_code(-1)
        return value[0]

    def SetScaleValueByType(self, scaleTypeId: int, value: float):
        """
        Set the scale used for displaying a chosen load or result diagram as shown in the Diagrams>Scales dialog.

        Parameters
        ----------
        scaleTypeId : int
            The index of the required load or result type to be set.
        value : float
            Value of scale type to be used.

        Returns
        -------
        bool
            True Value was successfully updated.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> retValue = staad_obj.View.SetScaleValueByType(1, 1.2)
        """
        retVal = self._view.SetScaleValueByType(scaleTypeId, value)
        if not retVal:
            raise_os_error_if_error_code(-1)
        return bool(retVal)

    def GetScaleCount(self):
        """
        Returns the count of scales that are used in STAAD.Pro which can be read and set using the functions GetScaleValues and SetScaleValues.

        Returns
        -------
        int
            Returns Positive_Value Count of scales.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> scaleCount = staad_obj.View.GetScaleCount()
        """
        return self._view.GetScaleCount()

    def DetachView(self):
        """
        Remove a view from the collection of saved views. The view to be removed must be open and the active window. Once the view is detached, the active window becomes <untitled>. To set the active window see the function OSViewUI::SetActiveWindow. Note that the active window view cannot be Whole Structure or untitled.

        Returns
        -------
        int
            1 if View successfully detached.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> status = staad_obj.View.DetachView()
        """
        retVal = self._view.DetachView()
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        elif retVal == 0:
            raise OsErrorBase("Unable to Detach Active View.", -1)
        return retVal

    def RenameView(self, viewName: str):
        """
        Renames a saved view. The view should be open and be the active window.

        Parameters
        ----------
        viewName : str
            New name of the saved view

        Returns
        -------
        int
            Returns 1 if Rename view is successful.
            Returns 0 if Unsuccessful
            Returns 2 if View name already used.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> status = staad_obj.View.RenameView("view1")
        """
        retVal = self._view.RenameView(viewName)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        elif retVal == 0:
            raise OsErrorBase("Unable to Rename Active View.", -1)
        return retVal

    def OpenView(self, viewName: str, windowOptions: bool):
        """
        Open a previously saved view in either the active window or create a new window which becomes the active window.

        Parameters
        ----------
        viewName : str
            New name of the saved view
        windowOptions : bool
            False = Creates a new window for the view which becomes the active window, True / 1 = Display the view in the current active window

        Returns
        -------
        int
            Returns 1 if View Successfully opened.
            Returns 0 if Unsuccessful
            Returns 2 if View name does not exist.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> status = staad_obj.View.OpenView("view1", True)
        >>> print(status)
        """
        retVal = self._view.OpenView(viewName, windowOptions)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        elif retVal == 0:
            raise OsErrorBase("Unable to open view.", -1)
        return retVal

    def SaveView(self, viewName: str, overWrite: bool):
        """
        Save the active graphic view to the collection of saved views which can be opened using the function OpenView.

        Parameters
        ----------
        viewName : str
            New name for the view
        overWrite : bool
            Option to overwrite if the given viewName already exists. False = Do not overwrite., TRUE = Overwrite viewName if it exists.

        Returns
        -------
        int
            Returns 1 if Save view is successful.
            Returns 0 if Unsuccessful
            Returns 2 if View name already exist and overWrite is false.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> status = staad_obj.View.SaveView("view1", True)
        >>> print(status)
        """
        retVal = self._view.SaveView(viewName, overWrite)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        elif retVal == 0:
            raise OsErrorBase("Unable to Save Active View.", -1)
        return retVal

    def GetWindowTitle(self, id: int):
        """
        Returns the Title of the Window.

        Parameters
        ----------
        id : int
            The index of the required Window (Type: Long). Note that IDs start from 1. Window IDs depend on the creation time; that is the last created window will have the last ID. Windows which are listed under in View tab > Windows Dropdown are supported in this API.

        Returns
        -------
        str
            Returns the Window string title.
            Returns Empty_String Window id not found.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> title = staad_obj.View.GetWindowTitle(1)
        """
        return self._view.GetWindowTitle(id)

    def GetWindowCount(self):
        """
        Get the number of windows currently open. This includes both graphic windows and tables.

        Returns
        -------
        int
            Returns Positive_Number The count of open Window.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> count = staad_obj.View.GetWindowCount()
        """
        return self._view.GetWindowCount()

    def CloseActiveWindow(self):
        """
        Closes the active graphic or table window, however there must be at least one graphic window remaining open. Note that window at position 1 can not be closed.

        Returns
        -------
        bool
            True if Window closed.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> output = staad_obj.View.CloseActiveWindow()
        """
        retVal = self._view.CloseActiveWindow()
        if not retVal:
            raise_os_error_if_error_code(-1)
        return retVal

    def SetActiveWindow(self, id: int):
        """
        Set a given window (active graphic or table window) with the provided id as the active window. The indexing starts from 1.

        Parameters
        ----------
        id : int
            The id of the window to be made the active window.

        Returns
        -------
        bool
            True if successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> status = staad_obj.View.SetActiveWindow(2)
        """
        retVal = self._view.SetActiveWindow(id)
        if not retVal:
            raise_os_error_if_error_code(-1)
        return retVal

    def SetDesignResults(self, utilization: int, color: bool, showValues: bool):
        """
        Sets Design Results to active view, this function replicates the setting of Design Results in the Diagrams>Design Results dialog.

        Parameters
        ----------
        ld : int
            Value of type Long. (0 = None, 1 = Actual Ratio, 2 = Normalised Ratio.)
        color : bool
            Value of type Boolean. (False/0 = Basic Colored, True/1 = Detailed Colored.)
        showValues : bool
            Value of type Boolean. (False/0 = Do Not Show Values, True/1 = Show Values.)

        Returns
        -------
        int
            Returns 1 if Set Design Results is successful.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> status = staad_obj.View.SetDesignResults(1, True, True)
        """
        retVal = self._view.SetDesignResults(utilization, color, showValues)
        if retVal < 0:
            raise_os_error_if_error_code(retVal)
        elif retVal == 0:
            raise OsErrorBase("Unable to set design results to active view.", -1)
        return retVal
