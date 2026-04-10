# ---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
# ---------------------------------------------------------------------------------------------
from .openStaadHelper import (
    create_bstr,
    make_byref,
)
from comtypes import CoInitialize
from .oserrors import raise_os_error_if_error_code


class OSTable:
    CoInitialize()

    def __init__(self, staadObj):
        self._staad = staadObj
        self._table = self._staad.Table

        self._functions = [
            "CreateReport",
            "SaveReport",
            "SaveReportAll",
            "GetReportCount",
            "AddTable",
            "RenameTable",
            "DeleteTable",
            "ResizeTable",
            "SaveTable",
            "GetTableCount",
            "SetCellValue",
            "GetCellValue",
            "SetColumnHeader",
            "SetColumnUnitString",
            "SetRowHeader",
            "SetCellTextColor",
            "SetCellTextBold",
            "SetCellTextItalic",
            "SetCellTextUnderline",
            "SetCellTextHorzAlignment",
            "SetCellTextVertAlignment",
            "SetCellTextSize",
            "SetCellTextSizeAll",
            "DeleteReport",
        ]

        for function_name in self._functions:
            self._table._FlagAsMethod(function_name)

    # Table Functions

    def CreateReport(self, report_title: str):
        """
        Creates a report with the specified title.

        Parameters
        ----------
        report_title : str
            A string containing the title of the report.

        Returns
        -------
        int
            Return report number\n
            Return 0 if Create Report error.

        Example
        -------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        """
        result = self._table.CreateReport(report_title)
        if result <= 0:
            raise_os_error_if_error_code(-1)
        return result

    def SaveReport(self, report_no: int):
        """
        Saves the specified report along with all its tables.

        Parameters
        ----------
        report_no : int
            Providing the report number identifying the STAAD report which is to be saved

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> staad_obj.Table.SaveReport(report_no)
        """
        self._table.SaveReport(report_no)

    def SaveReportAll(self):
        """
        Saves all the reports created.

        Returns
        -------
        None

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> staad_obj.Table.SaveReportAll()
        """
        self._table.SaveReportAll()

    def GetReportCount(self):
        """
        Returns the number of reports created.

        Returns
        -------
        int
            Return the number of reports created.

        Returns
        -------
        None

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_count = staad_obj.Table.GetReportCount()
        """
        return self._table.GetReportCount()

    def AddTable(self, report_no: int, table_name: str, row_count: int, col_count: int):
        """
        Add or create a report data table and returns the table no for a specified table.

        Parameters
        ----------
        report_no : int
            Report number.
        table_name  : str
            Report Table Name.
        row_count : int
            row count of the specified table.
        col_count : int
            column count of the specified table.

        Returns
        -------
        int
            Return table number\n
            Return 0 if Create table error.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        """
        result = self._table.AddTable(report_no, table_name, row_count, col_count)
        if result <= 0:
            raise_os_error_if_error_code(-1)
        return result

    def RenameTable(self, report_no: int, table_no: int, table_name: str):
        """
        Rename an existing report table.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no : int
            Table number in specified report.
        table_name : str
            Report Table Name.

        Returns
        -------
        None

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.RenameTable(report_no, table_number, "Table2")
        """
        self._table.RenameTable(report_no, table_no, table_name)

    def DeleteTable(self, report_no: int, table_no: int):
        """
        Delete an existing table from report data.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no : int
            Table number in specified report.

        Returns
        -------
        None

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.DeleteTable(report_no, table_number)
        """
        self._table.DeleteTable(report_no, table_no)

    def ResizeTable(self, report_no: int, table_no: int, row_nos: int, col_nos: int):
        """
        Resize existing table by increasing numbers of rows and columns from report data.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no : int
            Table number in specified report.
        row_nos : int
            row count for specified table.
        col_nos : int
            column count for specified table.

        Returns
        -------
        None

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.ResizeTable(report_no, table_number, 12, 6)
        """
        self._table.ResizeTable(report_no, table_no, row_nos, col_nos)

    def SaveTable(self, report_no: int, table_no: int):
        """
        Save a report table.

        Parameters
        ----------
        report_no : int
            Report number.
        TableNo  : int
            Table number in specified report.

        Returns
        -------
        None

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SaveTable(report_no, table_number)
        """
        self._table.SaveTable(report_no, table_no)

    def GetTableCount(self, report_no: int):
        """
        Get table count in a specified report

        Parameters
        ----------
        report_no : int
            Report number.

        Returns
        -------
        int
            Returns table count in a specified report

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> table_count = staad_obj.Table.GetTableCount(report_no)
        """
        return self._table.GetTableCount(report_no)

    def SetCellValue(
        self, report_no: int, table_no: int, row_no: int, col_no: int, value: str
    ):
        """
        Set cell values for a specified table.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no : int
            Table number in specified report.
        row_no : int
            Row number for specified table, start from 1.
        col_no : int
            Column number for specified table, start from 1.
        value : str
            Set cell value for specified cell in the table.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetCellValue(report_no, table_number, 1, 5, "abc")
        """
        self._table.SetCellValue(report_no, table_no, row_no, col_no, value)

    def GetCellValue(self, report_no: int, table_no: int, row_no: int, col_no: int):
        """
        Get cell values for a specified table.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no : int
            Table number in specified report.
        row_no : int
            Row number for specified table, start from 1.
        col_no : int
            Column number for specified table, start from 1.

        Returns
        -------
        string
            Returns cell value for specified cell in the table

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> table_count = staad_obj.Table.SetCellValue(report_no, table_number, 1, 5, "abc")
        >>> cell_value = staad_obj.Table.GetCellValue(report_no, table_number, 1, 5)
        """
        cell_value = create_bstr()
        ref_cell_value = make_byref(cell_value)
        self._table.GetCellValue(report_no, table_no, row_no, col_no, ref_cell_value)
        return cell_value.value

    def SetColumnHeader(self, report_no: int, table_no: int, col_no: int, header: str):
        """
        Sets column header for a specified table.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no : int
            Table number in specified report.
        col_no : int
            Column number for specified table, start from 1.
        header : str
            Column header for a specified column in table.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetColumnHeader(report_no, table_number, 1, "header1")
        """
        self._table.SetColumnHeader(report_no, table_no, col_no, header)

    def SetColumnUnitString(
        self, report_no: int, table_no: int, col_no: int, unit_string: str
    ):
        """
        Sets unit for the specific column of a specified table.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no : int
            Table number in specified report.
        col_no : int
            Column number for specified table, start from 1.
        unit_string : str
            Set unit for the specific column of a specified table.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetColumnUnitString(report_no, table_number, 1, "mm")
        """
        self._table.SetColumnUnitString(report_no, table_no, col_no, unit_string)

    def SetRowHeader(self, report_no: int, table_no: int, row_no: int, header: str):
        """
        Sets row header for the specific row of a specified table.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no  : int
            Table number in specified report.
        row_no : int
            Row number for specified table, start from 1
        header : str
            Set header for the specific row of a specified table

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetRowHeader(report_no, table_number, 1, "row1")
        """
        self._table.SetRowHeader(report_no, table_no, row_no, header)

    def SetCellTextColor(
        self,
        report_no: int,
        table_no: int,
        row_no: int,
        col_no: int,
        red: int,
        green: int,
        blue: int,
    ):
        """
        Sets the color of the text to be displayed in a particular cell. By default, the color is always set to black.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no  : int
            Table number in specified report.
        row_no : int
            Row number for specified table, start from 1.
        col_no : int
            Column number for specified table, start from 1.
        red : int
            An integer between 0 and 255 that represents the intensity of red in the color for the text.
        green : int
            An integer between 0 and 255 that represents the intensity of green in the color for the text.
        blue : int
            An integer between 0 and 255 that represents the intensity of blue in the color for the text.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetCellValue(report_no, table_number, 1, 5, "abc")
        >>> staad_obj.Table.SetCellTextColor(report_no, table_number, 1, 5, 240, 0, 0)
        """
        self._table.SetCellTextColor(
            report_no, table_no, row_no, col_no, red, green, blue
        )

    def SetCellTextBold(self, report_no: int, table_no: int, row_no: int, col_no: int):
        """
        Sets the text in a given row and column to bold.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no  : int
            Table number in specified report.
        row_no : int
            Row number for specified table, start from 1.
        col_no : int
            Column number for specified table, start from 1.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetCellValue(report_no, table_number, 1, 5, "abc")
        >>> staad_obj.Table.SetCellTextBold(report_no, table_number, 1, 5)
        """
        self._table.SetCellTextBold(report_no, table_no, row_no, col_no)

    def SetCellTextItalic(
        self, report_no: int, table_no: int, row_no: int, col_no: int
    ):
        """
        Italicizes the text in a given row and column.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no  : int
            Table number in specified report.
        row_no : int
            Row number for specified table, start from 1.
        col_no : int
            Column number for specified table, start from 1.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetCellValue(report_no, table_number, 1, 5, "abc")
        >>> staad_obj.Table.SetCellTextItalic(report_no, table_number, 1, 5)
        """
        self._table.SetCellTextItalic(report_no, table_no, row_no, col_no)

    def SetCellTextUnderline(
        self, report_no: int, table_no: int, row_no: int, col_no: int
    ):
        """
        Underlines the text in a given row and column.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no  : int
            Table number in specified report.
        row_no : int
            Row number for specified table, start from 1.
        col_no : int
            Column number for specified table, start from 1.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetCellValue(report_no, table_number, 1, 5, "abc")
        >>> staad_obj.Table.SetCellTextUnderline(report_no, table_number, 1, 5)
        """
        self._table.SetCellTextUnderline(report_no, table_no, row_no, col_no)

    def SetCellTextHorzAlignment(
        self, report_no: int, table_no: int, row_no: int, col_no: int, align: int
    ):
        """
        Sets the text in a particular row and column to a specified horizontal alignment. By default, all the text is right justified.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no  : int
            Table number in specified report.
        row_no : int
            Row number for specified table, start from 1.
        col_no : int
            Column number for specified table, start from 1.
        align : int
           align Sets the text in a particular row and column to a specified horizontal alignment. The possible values are:0 = left; 1 = center; 2 = right

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetCellValue(report_no, table_number, 1, 5, "abc")
        >>> staad_obj.Table.SetCellTextHorzAlignment(report_no, table_number, 1, 5, 2)
        """
        self._table.SetCellTextHorzAlignment(report_no, table_no, row_no, col_no, align)

    def SetCellTextVertAlignment(
        self, report_no: int, table_no: int, row_no: int, col_no: int, align: int
    ):
        """
        Sets the text in a particular row and column to a specified vertical alignment. By default, all the text is top justified.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no  : int
            Table number in specified report.
        row_no : int
            Row number for specified table, start from 1.
        col_no : int
            Column number for specified table, start from 1.
        align : int
           align Sets the text in a particular row and column to a specified vertical alignment. The possible values are: 0 = top; 4 = center; 8 = bottom

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetCellValue(report_no, table_number, 1, 5, "abc")
        >>> staad_obj.Table.SetCellTextVertAlignment(report_no, table_number, 1, 5, 2)
        """
        self._table.SetCellTextVertAlignment(report_no, table_no, row_no, col_no, align)

    def SetCellTextSize(
        self, report_no: int, table_no: int, row_no: int, col_no: int, size: float
    ):
        """
        Sets the text in a particular row and column to a certain font size. The font sizes are equivalent to the ones used in Microsoft Word.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no  : int
            Table number in specified report.
        row_no : int
            Row number for specified table, start from 1.
        col_no : int
            Column number for specified table, start from 1.
        size : float
           Containing the size of the font to set the text to.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.SetCellValue(report_no, table_number, 1, 5, "abc")
        >>> staad_obj.Table.SetCellTextSize(report_no, table_number, 1, 5, 12.25)
        """
        self._table.SetCellTextSize(report_no, table_no, row_no, col_no, size)

    def SetCellTextSizeAll(self, report_no: int, table_no: int, size: float):
        """
        Sets the text in the entire table to FontSize. The font sizes are equivalent to the ones used in Microsoft Word.

        Parameters
        ----------
        report_no : int
            Report number.
        table_no  : int
            Table number in specified report.
        size : float
           Containing the size of the font to set the text to.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 11, 5)
        >>> staad_obj.Table.SetCellTextSizeAll(report_no, table_number, 5)
        """
        self._table.SetCellTextSizeAll(report_no, table_no, size)

    def DeleteReport(self, report_no: int):
        """
        Delete an existing report.

        Parameters
        ----------
        report_no : int
            Report number.

        Examples
        --------
        >>> from openstaadpy import os_analytical
        >>> staad_obj = os_analytical.connect()
        >>> report_no = staad_obj.Table.CreateReport("testreport")
        >>> table_number = staad_obj.Table.AddTable(report_no, "Table1", 10, 5)
        >>> staad_obj.Table.DeleteReport(report_no)
        """
        self._table.DeleteReport(report_no)
