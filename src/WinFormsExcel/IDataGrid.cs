/* WinForms Excel library 
 * Copyright (c) 2021,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
/*Windows Forms Excel Library - IDataGrid interface
 http://smallsoft2.blogspot.ru/
 */

namespace ExtraControls
{
    /// <summary>
    /// Provides a set of properties and methods shared by DataGrid controls able to display 
    /// multiple data tables
    /// </summary>
    public interface IDataGrid
    {
        /// <summary>
        /// Gets or sets the content of currently active grid sheet via DataTable object
        /// </summary>
        /// <remarks>
        /// NOTE: Controls implementing IDataGrid are not guaranteed to support data binding. 
        /// The DataSource Property is only a convenient way to manipulate active sheet’s contents.
        /// </remarks>
        object DataSource {get;set;}

        /// <summary>
        /// Gets or sets current active sheet.
        /// </summary>
        int ActiveSheet { get; set; }

        /// <summary>
        /// Returns the amount of sheets in current workbook (read-only).
        /// </summary>
        int SheetsCount { get; }

        /// <summary>
        /// Sets contents of the cell specified by sheet, row and column numbers into an object 
        /// of any type
        /// </summary>
        /// <param name="sheet">Sheet number (1-based)</param>
        /// <param name="row">Row number</param>
        /// <param name="col">Column number</param>
        /// <param name="val">New cell value</param>
        void SetCellContent(int sheet, int row, int col, object val);

        /// <summary>
        /// Gets the content of the cell specified by sheet, row and column numbers.
        /// </summary>
        /// <param name="sheet">Sheet number (1-based)</param>
        /// <param name="row">Row number</param>
        /// <param name="col">Column number</param>
        /// <returns>
        /// Value of the specified cell, or null if arguments are incorrect
        /// </returns>
        object GetCellContent(int sheet, int row, int col);

        /// <summary>
        /// Fills the specified sheet with a content of given DataTable object
        /// </summary>
        /// <param name="sheet">Sheet number (1-based)</param>
        /// <param name="t">DataTable to fill sheet's content</param>
        void SetSheetContent(int sheet, DataTable t);

        /// <summary>
        /// Loads content of specified sheet as DataTable object.
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="FirstRowHasHeaders">Specifies that first row contains column headers</param>
        /// <param name="n_col">Maximum number of columns to load (0 - automatic)</param>
        /// <param name="n_row">Maximum number of rows to load (0 - automatic)</param>
        /// <returns>DataTable object filled with sheet content</returns>
        DataTable GetSheetContent(int sheet, bool FirstRowHasHeaders, int n_col = 0, int n_row = 0);

        /// <summary>
        /// Gets the number of currently active sheet. 
        /// </summary>
        /// <returns>Sheet number (1-based)</returns>
        int GetActiveSheet();

        /// <summary>
        /// Activates specified sheet in this control instance
        /// </summary>
        /// <param name="index">Sheet number (1-based)</param>
        void SetActiveSheet(int index);

        /// <summary>
        /// Removes the specified sheet
        /// </summary>
        /// <param name="index">Sheet number (1-based)</param>
        void DeleteSheet(int index);

        /// <summary>
        /// Adds new sheet into the workbook of this control instance
        /// </summary>
        /// <param name="name">Worksheet name (optional)</param>
        void AddSheet(string name = "");

        /// <summary>
        /// Gets names of all sheets in this control instance as a list of strings.
        /// </summary>
        /// <returns>List of worksheet names</returns>
        List<XlSheet> GetSheets();

        /// <summary>
        /// Changes the name of specified sheet.
        /// </summary>
        /// <param name="sheet">Sheet number (1-based)</param>
        /// <param name="name">Sheet name</param>
        void SetSheetName(int sheet, string name);

        /// <summary>
        /// Gets the name for the specified sheet
        /// </summary>
        /// <param name="sheet">Sheet number (1-based)</param>
        string GetSheetName(int sheet);

        /// <summary>
        /// Gets the index of sheet with specified name.
        /// </summary>
        /// <param name="name">Sheet name</param>
        /// <returns>Sheet index, or -1 if the sheet is not found</returns>
        int FindSheet(string name);

        /// <summary>
        /// Closes the current workbook, and loads a new empty workbook into this control
        /// </summary>
        void NewEmptyWorkbook();
    }
}
