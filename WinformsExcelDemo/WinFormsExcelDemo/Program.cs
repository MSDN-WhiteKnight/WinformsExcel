/* WinForms Excel library 
 * Copyright (c) 2020,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

/* WinForms Excel library demo application: Main entry point */

namespace WinFormsExcelDemo
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
