/* WinForms Excel library 
 * Copyright (c) 2021,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.Text;

namespace ExtraControls
{
    /// <summary>
    /// Represents mode in which C_AdvancedDataGrid control operates
    /// </summary>
    public enum GridMode
    {
        /// <summary>
        /// Undefined (control is not initialized yet)
        /// </summary>
        Undefined = 0,
        
        /// <summary>
        /// Choose mode automatically: Excel mode if possible, otherwise Substitute
        /// </summary>
        Auto = 1,
        
        /// <summary>
        /// Uses Excel implementation (AdvancedDataGrid)
        /// </summary>        
        Excel = 2,
        
        /// <summary>
        /// Uses substituted implementation (S_AdvancedDataGrid)
        /// </summary>
        Substitute =3
    }
}
