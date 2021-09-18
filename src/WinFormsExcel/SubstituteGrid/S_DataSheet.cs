/* WinForms Excel library 
 * Copyright (c) 2020,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;
/*Windows Forms Excel Library - S_DataSheet class
 http://smallsoft2.blogspot.ru/
 */

namespace ExtraControls.SubstituteGrid
{
    /// <summary>
    /// Represents S_AdvancedDataGrid's data sheet. Infrastructure.
    /// </summary>
    public class S_DataSheet
    {
        /// <summary>
        /// Sheet name, displayed in associated button
        /// </summary>
        public string name;

        /// <summary>
        /// Button used to activated this sheet
        /// </summary>
        public Button title_button;

        /// <summary>
        /// actual data object this sheet is bound to
        /// </summary>
        public DataTable data;
    }
}
