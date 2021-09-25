/* WinForms Excel library 
 * Copyright (c) 2021,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Text;
/*Windows Forms Excel Library - XlSheet object
 http://smallsoft2.blogspot.ru/
 */

namespace ExtraControls
{
    /// <summary>
    /// Represents data sheet in Excel Workbook
    /// </summary>
    public class XlSheet
    {
        /// <summary>
        /// Index of the sheet in the workbook (starts from 1)
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Name of the sheet (displayed in tab)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Specifies that this sheet is a chart. 
        /// Unlike with worksheets, you can’t manipulate the contents of chart sheets.
        /// </summary>
        public bool IsChart { get; set; }
    }
}
