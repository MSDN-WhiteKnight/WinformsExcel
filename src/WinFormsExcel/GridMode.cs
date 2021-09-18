/* WinForms Excel library 
 * Copyright (c) 2020,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExtraControls
{
    public enum GridMode
    {
        Undefined = 0, //not initialized yet
        Auto = 1, //Choose mode automatically: Excel mode if possible, otherwise Substitute
        Excel = 2, //Uses AdvancedDataGrid
        Substitute =3 //Uses S_AdvancedDataGrid
    }
}
