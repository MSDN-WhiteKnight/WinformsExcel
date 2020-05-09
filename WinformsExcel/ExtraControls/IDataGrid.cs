/* WinForms Excel library 
 * Copyright (c) 2020,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
/*Windows Forms Excel Library - IDataGrid interface
 http://smallsoft2.blogspot.ru/
 */

namespace ExtraControls
{
    /// <summary>
    /// Provides a set of properties and methods shared by DataGrid controls able to display multiple data tables
    /// </summary>
    public interface IDataGrid
    {

        object DataSource {get;set;}
        int ActiveSheet { get; set; }
        int SheetsCount { get; }        
        
       void SetCellContent(int sheet, int row, int col, object val);
        object GetCellContent(int sheet, int row, int col);        
        void SetSheetContent(int sheet, DataTable t);
        DataTable GetSheetContent(int sheet, bool FirstRowHasHeaders, int n_col = 0, int n_row = 0);
        int GetActiveSheet();
        void SetActiveSheet(int index);
        void DeleteSheet(int index);
        void AddSheet(string name = "");
        List<XlSheet> GetSheets();
        void SetSheetName(int sheet, string name);
        string GetSheetName(int sheet);
        int FindSheet(string name);        
        void NewEmptyWorkbook();


    }
}
