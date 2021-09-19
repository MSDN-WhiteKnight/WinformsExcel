/* WinForms Excel library 
 * Copyright (c) 2021,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace WinFormsExcel.Internal
{
    internal static class NativeMethods
    {
        /*Объявления неуправляемых WINAPI функций*/

        /// <summary>
        /// Changes hWnd's owner to NewParent window 
        /// </summary>
        /// <param name="hWnd">handle of the window to change owner</param>
        /// <param name="NewParent">handle of the new owner window</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern int SetParent(IntPtr hWnd, IntPtr NewParent);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        /// <summary>
        /// Sets window property defined by dword's offset in window structure
        /// </summary>
        /// <param name="hWnd">handle of the window to change property</param>
        /// <param name="nIndex">property dword's offset in window structure</param>
        /// <param name="dwNewLong">new dword value of the property</param>        
        [DllImport("user32.dll")]
        public static extern uint SetWindowLong(IntPtr hWnd, int nIndex, uint dwNewLong);

        /// <summary>
        /// Gets the value of the window property defined by dword's offset in window structure
        /// </summary>
        /// <param name="hWnd">handle of the window to get property from</param>
        /// <param name="nIndex">property dword's offset in window structure</param>        
        [DllImport("user32.dll")]
        public static extern uint GetWindowLong(IntPtr hWnd, int nIndex);

        /// <summary>
        /// Adjusts window position and size based on coordinates, width and height values
        /// </summary>
        /// <param name="hWnd">handle of the window to adjust values</param>
        /// <param name="x">X coordinate of window on the screen</param>
        /// <param name="y">Y coordinate of window on the screen</param>
        /// <param name="w">window's width</param>
        /// <param name="h">window's height</param>
        /// <param name="repaint">repaint window after adjusting</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern int MoveWindow(IntPtr hWnd, int x, int y, int w, int h, int repaint);

        /*объявления констант для WINAPI функций*/
        /// <summary>
        /// Window style: child window, has no titlebar or sizebox
        /// </summary>
        public const uint WS_CHILD = 0x40000000;

        /// <summary>
        /// Window style: popup window
        /// </summary>
        public const uint WS_POPUP = 0x80000000;

        /// <summary>
        /// Window style: has border
        /// </summary>
        public const uint WS_BORDER = 0x00800000;

        /// <summary>
        /// Window style:  WS_BORDER | WS_DLGFRAME
        /// </summary>
        public const uint WS_CAPTION = 0x00C00000;

        /// <summary>
        /// Window style:  frame allows to resize this window
        /// </summary>
        public const uint WS_THICKFRAME = 0x00040000;

        /// <summary>
        /// Window style:  frame allows to resize this window
        /// </summary>
        public const uint WS_SIZEBOX = WS_THICKFRAME;

        public const uint WS_SYSMENU = 0x00080000;

        /// <summary>
        /// The offset of style dword in window structure (passed to GetWindowLong or SetWindowLong)
        /// </summary>
        public const int GWL_STYLE = (-16);//смещение стиля в структуре окна

        public const int SW_HIDE = 0;
        public const int SW_SHOW = 5;
    }
}
