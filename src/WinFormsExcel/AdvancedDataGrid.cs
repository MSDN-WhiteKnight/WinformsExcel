/* WinForms Excel library 
 * Copyright (c) 2021,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using WinFormsExcel.Internal;

/*Windows Forms user control providing the functionality of displaying and editing data in MS Excel window.
 http://smallsoft2.blogspot.ru/
*/

namespace ExtraControls
{
    /// <summary>
    /// Provides the functionality of hosting MS Excel window in Windows Forms application in order to display and edit
    /// one or several tables of data
    /// </summary>
    public partial class AdvancedDataGrid : UserControl, IDisposable, IDataGrid
    {
        #region PROTECTED VARIABLES

        /// <summary>
        /// Underlying Excel appplication of this control instance
        /// </summary>
        protected Excel.Application _Xl=null;//приложение Excel
        int _ver;
        bool _editmode = false;

        /// <summary>
        /// Indicates that excel was loaded for this control instance
        /// </summary>
        protected bool _Initialized = false;//указывает, что Excel загружен

        /// <summary>
        /// Indicates that formula bar should be displayed
        /// </summary>
        protected bool display_formula_bar = false;

        /// <summary>
        /// Indicates that status bar should be displayed
        /// </summary>
        protected bool display_status_bar = false;

        /// <summary>
        /// Indicates that window title should be displayed
        /// </summary>
        protected bool display_window_title = false;

        /// <summary>
        /// Indicates that the user input is disabled for this control
        /// </summary>
        protected bool disabled=false;

        List<string> tmp_file_names = new List<string>(10);

        #endregion

        #region PUBLIC PROPERTIES

        /// <summary>
        /// Specifies that the use input is disabled for this control
        /// </summary>
        [Category("Behavior"), Browsable(true), EditorBrowsable(EditorBrowsableState.Always)]
        [Description("Specifies if control intaracts with user input"),DefaultValue(false)]
        public bool Inactive
        {
            get { return this.disabled; }
            set
            {
                if (this._Initialized)
                {
                    _Xl.Interactive = !value;
                }
                this.disabled = value;                
            }
        }

        /// <summary>
        /// Gets or sets the content of currently active Excel sheet via DataTable object        
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public object DataSource
        {
            set
            {
                if (!_Initialized) return;

                if (value is DataTable)
                {
                    try
                    {
                        this.SetSheetContent(this.GetActiveSheet(), value as DataTable);
                    }
                    catch (Exception) { }
                }
            }

            get
            {
                if (!_Initialized) return null;
                try
                {
                    return this.GetSheetContent(this.GetActiveSheet(), true);
                }
                catch (Exception) { return null; }
            }

        }

        /// <summary>
        /// Enables the standard Excel status bar below worksheet area
        /// </summary>
        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always),
        Category("Appearance"),
        Description("Enables standard excel status bar below worksheet area."), DefaultValue(false)]
        public bool DisplayStatusBar
        {
            get { return this.display_status_bar; }
            set
            {
                if (this._Initialized)
                {
                    if (this._ver >= VersionExcel2016)
                    {
                        throw new NotSupportedException(
                            "Modifying this property after the grid is initialized is not supported in Excel 2016+"
                            );
                    }

                    _Xl.DisplayStatusBar = value;
                }
                this.display_status_bar = value;
            }

        }

        /// <summary>
        /// Enables the standard Excel formula bar above worksheet area
        /// </summary>
        [Browsable(true),EditorBrowsable(EditorBrowsableState.Always),
        Category("Appearance"),
        Description("Enables standard excel formula bar above worksheet area."),DefaultValue(false)]
        public bool DisplayFormulaBar
        {
            get { return this.display_formula_bar; }
            set
            {
                if (this._Initialized)
                {
                    if (this._ver >= VersionExcel2016)
                    {
                        throw new NotSupportedException(
                            "Modifying this property after the grid is initialized is not supported in Excel 2016+"
                            );
                    }

                    _Xl.DisplayFormulaBar = value;
                }
                this.display_formula_bar = value;
            }

        }

        /// <summary>
        /// Enables Excel window title, including the ribbon menu in new versions
        /// </summary>
        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always),
        Category("Appearance"),
        Description("Enables Excel window title, including ribbon menu in new versions"), DefaultValue(false)]
        public bool DisplayWindowTitle
        {
            get { return this.display_window_title; }
            set
            {
                if (this._Initialized)
                {
                    if (this._ver >= VersionExcel2016)
                    {
                        throw new NotSupportedException(
                            "Modifying this property after the grid is initialized is not supported in Excel 2016+"
                            );
                    }
                }

                this.display_window_title = value;
            }
        }

        /// <summary>
        /// Gets or sets current active sheet.                
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public int ActiveSheet
        {
            get
            {
                if (!_Initialized) return -1;
                try
                {
                    return this.GetActiveSheet();
                }
                catch (Exception) { return -1; }
            }
            set
            {
                if (!_Initialized) return;
                try
                {
                    this.SetActiveSheet(value);
                }
                catch (Exception) { return; }
            }
        }

        /// <summary>
        /// Gets the underlying Excel Application object of this control instance (read-only)
        /// </summary>
        [Browsable(false),EditorBrowsable(EditorBrowsableState.Never)]
        public Excel.Application XlApplication
        {
            get
            {
                return this._Xl;
            }

        }

        /// <summary>
        /// Gets the count of sheets in an active workbook
        /// </summary>
        public int SheetsCount
        {
            get
            {

                if (!_Initialized) return -1;

                Excel.Workbook wb = null;
                Excel.Sheets sh = null;


                try
                {
                    wb = this.GetWorkbook();
                    sh = wb.Sheets;

                    int c = sh.Count;
                    return c;

                }
                catch (Exception)
                {
                    return -1;
                }
                finally
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                    if (sh != null) Marshal.ReleaseComObject(sh);
                }
                
            }
        }

        /// <summary>
        /// Specifies the initial count of sheets for this control
        /// </summary>
        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always),        
        Description("Specifies the initial count of sheets in the workbook"), DefaultValue(1)]
        public int InitialSheetsCount { get; set; }

        /// <summary>
        /// Specifies the workbook file to open when initializing this control
        /// </summary>
        [Browsable(true), EditorBrowsable(EditorBrowsableState.Always),
        Description("Specifies workbook file to open when initializing this control "), 
        DefaultValue("")]
        public string SourceFile { get; set; }

        #endregion

        /// <summary>
        /// Creates AdvacedDataGrid control in uninitialized state
        /// </summary>
        public AdvancedDataGrid()
        {
            InitializeComponent();

            this._Initialized = false;
            this.InitialSheetsCount = 1;
            this.SourceFile = "";
        }

        void EmbedExcel()
        {
            IntPtr wnd = (IntPtr)_Xl.Hwnd;//дескриптор окна Excel
            NativeMethods.SetParent(wnd, this.Handle);//изменение владельца окна Excel на этот элемент управления

            uint style = NativeMethods.GetWindowLong(wnd, NativeMethods.GWL_STYLE);//получим стиль окна
            style &= ~NativeMethods.WS_SIZEBOX;//уберем возможность изменения размера

            if (!this.display_window_title)
            {
                style &= ~NativeMethods.WS_CAPTION;//уберем заголовок
            }

            NativeMethods.SetWindowLong(wnd, NativeMethods.GWL_STYLE, style);//установка нового стиля

            //установка размеров окна
            NativeMethods.MoveWindow(wnd, 0, 0, this.ClientRectangle.Width, this.ClientRectangle.Height, 1);
        }

        void UnembedExcel()
        {
            IntPtr wnd = (IntPtr)_Xl.Hwnd;//дескриптор окна Excel
            NativeMethods.SetParent(wnd, IntPtr.Zero);            
            _Xl.WindowState = Excel.XlWindowState.xlMaximized;
            _Xl.DisplayExcel4Menus = false;
        }

        void EnterEditMode()
        {
            if (this._editmode) return;

            if(this._ver>=VersionExcel2016)this.UnembedExcel();

            this._editmode = true;
        }

        void LeaveEditMode()
        {
            if (this._editmode == false) return;

            if (this._ver >= VersionExcel2016) this.EmbedExcel();

            this._editmode = false;
        }

        /// <summary>
        /// Initializes Excel application and creates empty workbook for this control instance
        /// </summary>
        public void InitializeExcel()//загрузка Excel
        {
            Excel.Workbooks wbs = null;
            Excel.Workbook book = null;
            Excel.Sheets sheets = null;

            try
            {
                _Xl = new Excel.Application();//запуск приложения
                this._ver = this.GetExcelVersion();

                if (!this.display_window_title)
                {
                    _Xl.WindowState = Excel.XlWindowState.xlMinimized;
                }

                _Xl.DisplayExcel4Menus = false;//выключить меню
                _Xl.DisplayFormulaBar = display_formula_bar;//выключить строку формул
                _Xl.ShowWindowsInTaskbar = false;//не показывать в панели задач
                _Xl.DisplayAlerts = false;//не показывать сообщения
                _Xl.DisplayStatusBar = display_status_bar;//не показывать строку состояния   
                _Xl.Interactive = !disabled;
                _Xl.Visible = true;//окно видимо
                
                wbs = _Xl.Workbooks;

                if (this.SourceFile == String.Empty)
                {
                    book = wbs.Add(Type.Missing);//создание новой пустой книги
                    sheets = book.Sheets;

                    while (true)
                    {
                        if (sheets.Count >= this.InitialSheetsCount) break;
                        sheets.Add();
                    }
                }
                else book = wbs.Open(Filename: this.SourceFile);   

                EmbedExcel();

                this._Initialized = true;//Excel загружен!
            }
            catch (Exception)
            {
                if (_Xl != null)
                {
                    _Xl.Quit();
                    Marshal.ReleaseComObject(_Xl);
                    this._Initialized = false;
                    _Xl = null;
                }
            }
            finally
            {
                if(book!=null)Marshal.ReleaseComObject(book);
                if (wbs != null) Marshal.ReleaseComObject(wbs);
                if (sheets != null) Marshal.ReleaseComObject(sheets);
            }
        }

        /// <summary>
        /// Cleans up resources of current AdvancedDataGrid control, breaking cross-process window relationships
        /// and closing Excel application. The control will be in uninitialized state.
        /// </summary>
        public void Destroy()//освобождение ресурсов
        {
            if (_Xl != null)
            {
                try
                {
                    _Xl.Visible = false;
                    NativeMethods.SetParent((IntPtr)_Xl.Hwnd, (IntPtr)0);//убираем владельца окна                    
                }
                catch (Exception) { }

                try
                {
                    /*Restore original Excel UI state*/
                    
                    _Xl.DisplayExcel4Menus = true;
                    _Xl.DisplayFormulaBar = true;
                    _Xl.ShowWindowsInTaskbar = true;
                    _Xl.DisplayAlerts = true;
                    _Xl.DisplayStatusBar = true;
                    _Xl.WindowState = Excel.XlWindowState.xlMaximized;
                    _Xl.Visible = false;

                    Excel.Workbook wb = this.GetWorkbook();

                    if (wb != null)
                    {
                        wb.Close(false, Type.Missing, Type.Missing);//закрытие книги
                        Marshal.ReleaseComObject(wb);
                    }
                }
                catch (Exception) { }

                try
                {
                    _Xl.Quit();//выход из приложения
                    Marshal.ReleaseComObject(_Xl);                    
                }
                catch (Exception) { }

                _Xl = null;
            }

            if (tmp_file_names != null)
            {
                foreach (string s in tmp_file_names)
                {
                    try
                    {
                        System.IO.File.Delete(s);
                    }
                    catch (System.IO.IOException) { }
                }
                tmp_file_names.Clear();
            }

            _Initialized = false;
        }


        /// <summary>
        /// Destroys this control instance
        /// </summary>
        ~AdvancedDataGrid()
        {
            this.Destroy();
        }

        /// <summary>
        /// Adjusts Excel window size to fit AdvancedDataGrid control's new size
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AdvancedDataGrid_Resize(object sender, EventArgs e)//подгонка размеров окна при изменении размера элемента
        {
            if (!_Initialized) return;
            if (_Xl == null) return;
            try
            {
                NativeMethods.MoveWindow(
                    (IntPtr)(_Xl.Hwnd), 0, 0, this.ClientRectangle.Width, this.ClientRectangle.Height, 1
                    );
            }
            catch (Exception) { }
        }

        Excel.Workbook GetWorkbook()
        {
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;

            try
            {
                wb = _Xl.ActiveWorkbook;

                if (wb == null)
                {
                    wbs = _Xl.Workbooks;

                    if (wbs.Count == 0) return null;

                    wb = wbs[1];
                }
            }
            finally
            {
                if (wbs != null) Marshal.ReleaseComObject(wbs);
            }

            return wb;
        }

        const int VersionExcel2013 = 15;
        const int VersionExcel2016 = 16;

        int GetExcelVersion()
        {
            int res;
            string s = _Xl.Version.Trim();
            string[] arr = s.Split(new char[]{'.'},StringSplitOptions.RemoveEmptyEntries);

            if (arr.Length == 0) return -1;
                        
            if (Int32.TryParse(arr[0], out res)) return res;
            else return -1;
        }

        #region Data Access
        /// <summary>
        /// Sets contents of the cell specified by sheet, row and column numbers into an object of any type
        /// 
        /// Throws:
        /// InvalidOperationException (Excel is not initialized)
        /// ArgumentException (sheet number is incorrect);
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="row">Row number</param>
        /// <param name="col">Column number</param>
        /// <param name="val">New cell value</param>
        public void SetCellContent(int sheet, int row, int col, object val)//установка значения ячейки
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (sheet <= 0) throw new ArgumentException("sheet number must be positive"); 
            if (row < 0) return;
            if (col < 0) return;

            bool pr = false;
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;
            Excel.Worksheet sh = null;
            object obj = null;

            try
            {
                wb = this.GetWorkbook();

                obj=wb.Sheets[sheet];//получение листа
                if (obj is Excel.Worksheet)
                {
                    sh = (Excel.Worksheet)obj;

                    if (this._ver < 16)
                    {
                        if (sh.ProtectContents)
                        {
                            sh.Protect(Contents: false);
                            pr = true;
                        }
                    }

                    sh.Cells[row, col] = val;//установка значения

                    if (pr) sh.Protect(Contents: true);
                }
                else
                {
                    throw new InvalidOperationException("sheet #"+sheet.ToString()+" is not a worksheet");
                }
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if(sh!=null) Marshal.ReleaseComObject(sh);
                if (obj != null) Marshal.ReleaseComObject(obj);
                if (wbs != null) Marshal.ReleaseComObject(wbs);
            }
            
        }

        /// <summary>
        /// Gets the content of the cell specified by sheet, row and column numbers.
        /// Returns null in case of incorrect arguments.
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="row">Row number</param>
        /// <param name="col">Column number</param>
        /// <returns>Value of the specified cell</returns>
        public object GetCellContent(int sheet, int row, int col)//значения ячейки
        {
            if (!_Initialized) return null;
            if (sheet < 0) return null;
            if (row < 0) return null;
            if (col < 0) return null;

            Excel.Workbook wb = null;
            Excel.Worksheet sh = null;
            Excel.Range r = null;
            object obj = null;
            object val;

            try
            {
                wb = this.GetWorkbook();

                obj=wb.Sheets[sheet];//получение листа
                if (obj is Excel.Worksheet)
                {
                    sh = (Excel.Worksheet)obj;

                    r = sh.Cells[row, col] as Excel.Range;//значения 
                    val = r.Value2;
                }
                else
                {
                    val = null;
                }
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (r != null) Marshal.ReleaseComObject(r);
                if (obj != null) Marshal.ReleaseComObject(obj);
            }
            
            return val;
        }

        /// <summary>
        /// Gets the address of the specified cell
        /// </summary>
        /// <param name="sheet">Sheet number (1-based)</param>
        /// <param name="row">Cell row number (1-based)</param>
        /// <param name="col">Cell column number (1-based)</param>        
        /// <remarks>
        /// The address consists of column represented by letter, and row represented by number 
        /// (such as "B12" for 2nd column 12th row).
        /// </remarks>
        public string GetCellAddress(int sheet, int row, int col)
        {
            if (!_Initialized) return null;
            if (sheet < 0) return null;
            if (row < 0) return null;
            if (col < 0) return null;

            Excel.Workbook wb = null;
            Excel.Worksheet sh = null;
            Excel.Range r = null;
            object obj = null;
            string val = null;

            try
            {

                wb = this.GetWorkbook();

                obj = wb.Sheets[sheet];//получение листа
                if (obj is Excel.Worksheet)
                {
                    sh = (Excel.Worksheet)obj;

                    r = sh.Cells[row, col] as Excel.Range;//значения 
                    val = r.Address;
                }
                else
                {
                    val = null;
                }
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (r != null) Marshal.ReleaseComObject(r);
                if (obj != null) Marshal.ReleaseComObject(obj);
            }

            return val;
        }


        /// <summary>
        /// Fills the specified sheet with a content of given DataTable object
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="t">DataTable to fill sheet's content</param>
        /// <exception cref="InvalidOperationException">Excel is not initialized</exception>
        /// <exception cref="ArgumentException">Sheet number is incorrect</exception>
        public void SetSheetContent(int sheet, DataTable t)
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (sheet <= 0) throw new ArgumentException("sheet number must be positive");            
            bool pr = false;

            Excel.Workbook wb = null;
            Excel.Worksheet sh = null;
            object obj = null;

            try
            {
                wb = this.GetWorkbook();

                obj = wb.Sheets[sheet];//получение листа

                if (obj is Excel.Worksheet)
                {
                    sh = (Excel.Worksheet)obj;
                    
                    if (this._ver < 16)
                    {
                        if (sh.ProtectContents)
                        {
                            sh.Protect(Contents: false);
                            pr = true;
                        }
                    }
                    
                    int i, j;
                    
                    //attempting to set sheet name
                    if (this._ver < 16)
                    {
                        if (t.TableName.Trim() != "")
                        {
                            try { sh.Name = t.TableName; }
                            catch (Exception) { }
                        }
                    }

                    //filling column names
                    for (i = 0; i < t.Columns.Count; i++)
                    {
                        sh.Cells[1, i + 1] = t.Columns[i].ColumnName;
                    }

                    //filling data
                    for (i = 0; i < t.Rows.Count; i++)
                    {
                        for (j = 0; j < t.Columns.Count; j++)
                        {
                            sh.Cells[i + 2, j + 1] = t.Rows[i][j];
                        }
                    }

                    if (pr) sh.Protect(Contents: true);
                }
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (obj != null) Marshal.ReleaseComObject(obj);
            }
        }

        /// <summary>
        /// Loads content of specified sheet as DataTable object. 
        /// Tries to guess column types based on first cells (uses string as default type).
        /// 
        /// Throws:
        /// InvalidOperationException (Excel is not initialized)
        /// ArgumentException (passed arguments are incorrect);  
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="FirstRowHasHeaders">Specifies that first row contains column headers</param>
        /// <param name="n_col">Maximum number of columns to load (0 - automatic)</param>
        /// <param name="n_row">Maximum number of rows to load (0 - automatic)</param>
        /// <returns>DataTable object filled with sheet content</returns>
        public DataTable GetSheetContent(int sheet, bool FirstRowHasHeaders,int n_col=0,int n_row=0)
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (sheet <= 0) throw new ArgumentException("sheet number must be positive");
            if (n_col < 0) throw new ArgumentException("n_col must be non-negative");
            if (n_row < 0) throw new ArgumentException("n_row must be non-negative");
            
            DataTable t = new System.Data.DataTable();
            DataColumn col;
            DataRow row;

            Excel.Workbook wb = null;
            Excel.Worksheet sh = null;
            Excel.Range r = null;
            object obj = null;

            try
            {
                wb = this.GetWorkbook();
                obj=wb.Sheets[sheet];//получение листа
                if (obj is Excel.Worksheet)
                {
                    sh = (Excel.Worksheet)obj;

                    int i, j;
                    object val;

                    if (n_row == 0)//guessing amount of rows
                    {
                        i = 1;
                        while (true)
                        {
                            r = (sh.Cells[i, 1]);
                            if (r == null) break;
                            val = r.Value2;
                            Marshal.ReleaseComObject(r);
                            r = null;
                            if (val == null) break;
                            if (val.ToString() == "") break;
                            if (i > 20000) break;
                            i++;
                        }
                        n_row = i - 1;
                    }

                    if (n_col == 0)//guessing amount of columns
                    {
                        i = 1;
                        while (true)
                        {
                            r = (sh.Cells[1, i]);
                            if (r == null) break;
                            val = r.Value2;
                            Marshal.ReleaseComObject(r);
                            r = null;
                            if (val == null) break;
                            if (val.ToString() == "") break;
                            if (i > 200) break;
                            ;
                            i++;
                        }
                        n_col = i - 1;
                    }

                    //determining the first row containing the actual data
                    int start_row;
                    if (FirstRowHasHeaders) start_row = 2;
                    else start_row = 1;

                    //setting table name
                    t.TableName = sh.Name;

                    //adding columns
                    for (i = 1; i <= n_col; i++)
                    {
                        col = new System.Data.DataColumn();
                        //determining column name
                        if (FirstRowHasHeaders)
                        {
                            r = sh.Cells[1, i];
                            if (r != null)
                            {
                                val = (r).Value2;
                                Marshal.ReleaseComObject(r);
                                r = null;
                                if (val == null) val = "C" + i.ToString();
                                if (val.ToString().Trim() == "") val = "C" + i.ToString();
                            }
                            else
                            {
                                val = "C" + i.ToString();
                            }
                            col.ColumnName = val.ToString();
                        }
                        else
                        {
                            col.ColumnName = "C" + i.ToString();
                        }

                        //guessing column data type
                        r = sh.Cells[start_row, i];
                        if (r != null)
                        {
                            val = (r).Value2;
                            Marshal.ReleaseComObject(r);
                            r = null;
                            if (val == null) col.DataType = typeof(string);
                            else col.DataType = val.GetType();
                        }
                        else
                        {
                            col.DataType = typeof(string);
                        }

                        t.Columns.Add(col);
                    }


                    //filling data
                    for (i = start_row; i <= n_row; i++)
                    {
                        row = t.NewRow();
                        for (j = 1; j <= n_col; j++)
                        {
                            r = sh.Cells[i, j];
                            if (r != null)
                            {
                                row[j - 1] = (r).Value2;
                                Marshal.ReleaseComObject(r);
                                r = null;
                            }
                            else
                            {
                                row[j - 1] = DBNull.Value;
                            }
                        }
                        t.Rows.Add(row);
                    }
                }
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (r != null) Marshal.ReleaseComObject(r);
                if (obj != null) Marshal.ReleaseComObject(obj);
            }
            
            return t;
        }
        
        #endregion



        #region Sheet functions
        /// <summary>
        /// Gets the number of currently active sheet.
        /// Returns -1 if excel is not initialized.
        /// </summary>
        /// <returns>Sheet number</returns>
        public int GetActiveSheet()
        {
            if (!_Initialized) return -1;

            Excel.Workbook wb = null;
            Excel.Worksheet sh = null;
            Excel.Chart ch = null;
            object obj = null;
            int val = 0;

            try
            {
                wb = this.GetWorkbook();
                obj = wb.ActiveSheet;
                if (obj is Excel.Worksheet)
                {
                    sh = (Excel.Worksheet)obj;
                    val = sh.Index;
                }
                else if (obj is Excel.Chart)
                {
                    ch = (Excel.Chart)obj;
                    val = ch.Index;
                }
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (obj != null) Marshal.ReleaseComObject(obj);
                if (ch != null) Marshal.ReleaseComObject(ch);
            }
            
            return val;
        }

        /// <summary>
        /// Activates specified sheet in this control instance
        /// 
        /// Throws:
        /// InvalidOperationException (Excel is not initialized);
        /// ArgumentException (index is incorrect);
        /// </summary>
        /// <param name="index">Sheet number</param>
        public void SetActiveSheet(int index)
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (index <= 0) throw new ArgumentException("index must be positive");

            EnterEditMode();
            Excel.Workbook wb = null;
            Excel.Sheets sh = null;
            Excel.Worksheet wsh = null;
            object obj = null;
            Excel.Chart ch = null;

            try
            {
                wb = this.GetWorkbook();
                
                sh = wb.Sheets;
                obj = sh[index];
                if (obj is Excel.Worksheet)
                {
                    wsh = (Excel.Worksheet)obj;
                    wsh.Activate();
                }
                else if (obj is Excel.Chart)
                {
                    ch = (Excel.Chart)obj;
                    ch.Activate();
                }
                
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (wsh != null) Marshal.ReleaseComObject(wsh);
                if (obj != null) Marshal.ReleaseComObject(obj);
                if (ch != null) Marshal.ReleaseComObject(ch);
                LeaveEditMode();
            }
            
        }

        /// <summary>
        /// Removes specified sheet from workbook. 
        /// Note: You can't remove all sheets. At least one sheet must be present in workbook all the time.
        /// 
        /// Throws:
        /// InvalidOperationException (Excel is not initialized);
        /// ArgumentException (index is incorrect);
        /// </summary>
        /// <param name="index">Sheet number</param>
        public void DeleteSheet(int index)
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (this._ver >= 16) throw new NotSupportedException("This API is not supported in Excel 2016+");
            if (index <= 0) throw new ArgumentException("index must be positive");

            Excel.Workbook wb = null;
            Excel.Sheets sh = null;
            Excel.Worksheet wsh = null;
            object obj = null;
            Excel.Chart ch = null;

            try
            {
                wb = this.GetWorkbook();
                sh = wb.Sheets;

                obj = sh[index];
                if (obj is Excel.Worksheet)
                {
                    wsh = (Excel.Worksheet)obj;
                    wsh.Delete();
                }
                else if (obj is Excel.Chart)
                {
                    ch = (Excel.Chart)obj;
                    ch.Delete();
                }
                
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (wsh != null) Marshal.ReleaseComObject(wsh);
                if (obj != null) Marshal.ReleaseComObject(obj);
                if (ch != null) Marshal.ReleaseComObject(ch);
            }
            
        }

        /// <summary>
        /// Adds new sheet into the workbook of this control instance
        /// 
        /// Throws:
        /// InvalidOperationException (Excel is not initialized);
        /// </summary>
        /// <param name="name">Worksheet name (optional)</param>
        public void AddSheet(string name="")
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (this._ver >= VersionExcel2016) throw new NotSupportedException("This API is not supported in Excel 2016+");
            
            Excel.Workbook wb = null;
            Excel.Sheets sh = null;
            Excel.Worksheet wsh = null;

            try
            {
                wb = this.GetWorkbook();
                sh = wb.Worksheets;

                int c = sh.Count;

                if (name.Trim() == "") name = "Sheet " + (c + 1).ToString();
                wsh = sh.Add();
                wsh.Name = name;
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (wsh != null) Marshal.ReleaseComObject(wsh);
            }
        }

        /// <summary>
        /// Gets names of all sheets in this control instance as a list of string objects.
        /// 
        /// Throws:
        /// InvalidOperationException (Excel is not initialized);
        /// </summary>
        /// <returns>List of worksheet names</returns>
        public List<XlSheet> GetSheets()
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");

            Excel.Workbook wb = null;
            Excel.Sheets sh = null;
            List<XlSheet> sheets = null;
            XlSheet val;

            try
            {
                wb = this.GetWorkbook();
                sh = wb.Sheets;
                
                int c = sh.Count;
                sheets = new List<XlSheet>(c);

                foreach (object wsh in sh)
                {
                    try
                    {                        
                        val = new XlSheet();
                        if (wsh is Excel.Worksheet)
                        {
                            val.Name = (wsh as Excel.Worksheet).Name;
                            val.Index = (wsh as Excel.Worksheet).Index;
                            val.IsChart = false;
                            
                        }
                        else if (wsh is Excel.Chart)
                        {
                            val.Name = (wsh as Excel.Chart).Name;
                            val.Index = (wsh as Excel.Chart).Index;
                            val.IsChart = true;
                        }
                        sheets.Add(val);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(wsh);
                    }
                }
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);                
            }
            
            return sheets;
        }

        


        /// <summary>
        /// Changes the name of specified sheet.
        /// 
        /// Throws:
        /// InvalidOperationException (Excel is not initialized);
        /// ArgumentException (index is incorrect or name is not specified);
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="name">Sheet ne name</param>
        public void SetSheetName(int sheet,string name)
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (this._ver >= 16) throw new NotSupportedException("This API is not supported in Excel 2016+");
            if (sheet <= 0) throw new ArgumentException("index must be positive");
            if (name== null) throw new ArgumentException("name can't be null");
            if (name.Trim().Length <= 0) throw new ArgumentException("name can't be omitted");
            
            Excel.Workbook wb = null;
            Excel.Sheets sh = null;
            Excel.Worksheet wsh = null;
            object obj = null;
            Excel.Chart ch = null;

            try
            {
                wb = this.GetWorkbook();
                sh = wb.Sheets;

                obj = sh[sheet];
                if (obj is Excel.Worksheet)
                {
                    wsh = (Excel.Worksheet)obj;
                    wsh.Name = name;
                }
                else if (obj is Excel.Chart)
                {
                    ch = (Excel.Chart)obj;
                    ch.Name = name;
                }
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (wsh != null) Marshal.ReleaseComObject(wsh);
                if (obj != null) Marshal.ReleaseComObject(obj);
                if (ch != null) Marshal.ReleaseComObject(ch);
            }
            
        }

        /// <summary>
        /// Gets the name for the specified sheet
        /// </summary>
        /// <param name="sheet">Sheet number (1-based)</param>
        public string GetSheetName(int sheet)
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (sheet <= 0) throw new ArgumentException("index must be positive");            

            Excel.Workbook wb = null;
            Excel.Sheets sh = null;
            Excel.Worksheet wsh = null;
            object obj = null;
            Excel.Chart ch = null;
            string name=null;

            try
            {
                wb = this.GetWorkbook();
                sh = wb.Sheets;

                obj = sh[sheet];
                if (obj is Excel.Worksheet)
                {
                    wsh = (Excel.Worksheet)obj;
                    name = wsh.Name;
                }
                else if (obj is Excel.Chart)
                {
                    ch = (Excel.Chart)obj;
                    name = ch.Name;
                }
                return name;
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (wsh != null) Marshal.ReleaseComObject(wsh);
                if (obj != null) Marshal.ReleaseComObject(obj);
                if (ch != null) Marshal.ReleaseComObject(ch);
            }

        }

        /// <summary>
        /// Gets the index of sheet with specified name. Returns -1 if the sheet is not found.
        /// 
        /// Throws:
        /// InvalidOperationException (Excel is not initialized);
        /// ArgumentException (name is not specified); 
        /// </summary>
        /// <param name="name">Sheet name</param>
        /// <returns>Sheet index or -1</returns>
        public int FindSheet(string name)
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");            
            if (name == null) throw new ArgumentException("name can't be null");
            if (name.Trim().Length <= 0) throw new ArgumentException("name can't be omitted");

            int index = -1;
            Excel.Workbook wb = null;
            Excel.Sheets sh = null;
            Excel.Worksheet wsh = null;
            Excel.Chart ch = null;
            int cindex=0;
            string cname="";

            try
            {
                wb = this.GetWorkbook();
                sh = wb.Sheets;

                foreach (object o in sh)
                {
                    try
                    {
                        if (o is Excel.Worksheet)
                        {
                            wsh = (o as Excel.Worksheet);
                            cindex = wsh.Index;
                            cname = wsh.Name;
                        }
                        else if (o is Excel.Chart)
                        {
                            ch = (o as Excel.Chart);
                            cindex = ch.Index;
                            cname = ch.Name;
                        }
                        if (cname == name)
                        {
                            index = cindex;//found!
                            break;
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(o);                        
                    }
                }
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (wsh != null) Marshal.ReleaseComObject(wsh);                
                if (ch != null) Marshal.ReleaseComObject(ch);
            }
            
            return index;           
        }

        /// <summary>
        /// Inserts the worksheet before or after the target worksheet in the workbook
        /// 
        /// Throws:
        /// InvalidOperationException (Excel is not initialized);
        /// ArgumentException (incorrect index); 
        /// </summary>
        /// <param name="curr_index">Number of sheet to move</param>
        /// <param name="new_index">Target sheet index</param>
        /// <param name="before">Specifies that sheet must be placed before the target sheet</param>
        public void MoveSheet(int curr_index, int new_index,bool before=true)
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (this._ver >= 16) throw new NotSupportedException("This API is not supported in Excel 2016+");
            if (curr_index <= 0) throw new ArgumentException("index must be positive");
            if (new_index <= 0) throw new ArgumentException("index must be positive");

            Excel.Worksheet wsh = null;
            Excel.Chart ch = null;
            Excel.Workbook wb = null;
            Excel.Sheets sh = null;
            object target = null;
            object obj = null;

            try
            {
                wb = this.GetWorkbook();
                sh = wb.Sheets;

                obj = sh[curr_index];
                target = sh[new_index];

                if (obj is Excel.Worksheet)
                {
                    wsh = sh[curr_index];
                    if (before)
                        wsh.Move(Before: target);
                    else
                        wsh.Move(After: target);
                }
                else if (obj is Excel.Chart)
                {
                    ch = sh[curr_index];
                    if (before)
                        ch.Move(Before: target);
                    else
                        ch.Move(After: target);
                }
                
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (wsh != null) Marshal.ReleaseComObject(wsh);
                if (obj != null) Marshal.ReleaseComObject(obj);
                if (target != null) Marshal.ReleaseComObject(target);
                if (ch != null) Marshal.ReleaseComObject(ch);
            }
            
        }
        #endregion

        public void SaveIntoFile(string file)
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            
            Excel.Workbook wb = null;

            string tmpfile = System.IO.Path.GetTempFileName();
            
            try { System.IO.File.Delete(tmpfile); }
            catch (Exception) { }

            try
            {
                EnterEditMode();
                wb = this.GetWorkbook();
                wb.SaveAs(tmpfile);
                this.tmp_file_names.Add(tmpfile);

                if (System.IO.File.Exists(file))
                {
                    try { System.IO.File.Delete(file); }
                    catch (Exception) { }
                }

                System.IO.File.Copy(tmpfile, file);
            }
            finally
            {
                if(wb!=null)Marshal.ReleaseComObject(wb);
                LeaveEditMode();
            }

        }

        public void NewEmptyWorkbook()
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (this._ver >= 16) throw new NotSupportedException("This API is not supported in Excel 2016+");
            
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;

            try
            {
                wbs = _Xl.Workbooks;
                wb = this.GetWorkbook();

                if (wb != null)
                {
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                }

                wb = null;
                wb = wbs.Add(Type.Missing);
            }
            finally
            {
                if (wbs != null) Marshal.ReleaseComObject(wbs);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }

        }

        public void OpenFile(string file)
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (this._ver >= 16) throw new NotSupportedException("This API is not supported in Excel 2016+");
            if (System.IO.File.Exists(file) == false) throw new System.IO.FileNotFoundException("File "+file+" not found", file);
            
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;

            try
            {
                wbs = _Xl.Workbooks;
                wb = this.GetWorkbook();
                
                wb.Close(false);
                Marshal.ReleaseComObject(wb);
                wb = null;
                wb = wbs.Open(Filename: file);
            }
            finally
            {
                if (wbs != null) Marshal.ReleaseComObject(wbs);
                if (wb != null) Marshal.ReleaseComObject(wb);
            }
        }

        void AddChartImpl(bool embedded,
            int sheet, string cell1, string cell2,
            double x, double y, double w, double h,
            ChartType ct = ChartType.xlXYScatterLines,
            string title = "")
        {
            Excel.Workbook wb = null;
            Excel.Sheets sh = null;
            Excel.Worksheet wsh = null;
            Excel.Chart ch = null;
            Excel.Range r = null;
            Excel.ChartTitle t = null;
            Excel.ChartObject co = null;
            Excel.ChartObjects cobjs = null;

            try
            {
                wb = this.GetWorkbook();
                sh = wb.Sheets;
                wsh = sh[sheet];

                if (embedded)
                {
                    cobjs = wsh.ChartObjects();
                    co = cobjs.Add(x, y, w, h);
                    ch = co.Chart;
                }
                else
                {
                    ch = wb.Charts.Add(After: wsh);
                }

                r = wsh.Range[cell1, cell2];

                ch.ChartType = (Excel.XlChartType)ct;
                ch.SetSourceData(r, Excel.XlRowCol.xlColumns);

                ch.HasLegend = false;
                ch.SizeWithWindow = true;

                if (title != "")
                {
                    ch.HasTitle = true;
                    t = ch.ChartTitle;
                    t.Caption = title;
                }
            }
            finally
            {
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (sh != null) Marshal.ReleaseComObject(sh);
                if (wsh != null) Marshal.ReleaseComObject(wsh);
                if (ch != null) Marshal.ReleaseComObject(ch);
                if (r != null) Marshal.ReleaseComObject(r);
                if (t != null) Marshal.ReleaseComObject(t);
                if (co != null) Marshal.ReleaseComObject(co);
                if (cobjs != null) Marshal.ReleaseComObject(cobjs);
            }
        }

        public void AddChart(int sheet, string cell1, string cell2, ChartType ct = ChartType.xlXYScatterLines,string title = "" )
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");
            if (this._ver >= 16) throw new NotSupportedException("This API is not supported in Excel 2016+");

            this.AddChartImpl(false, sheet, cell1, cell2, 0, 0, 0, 0, ct, title);
        }

        public void AddEmbeddedChart(
            int sheet, string cell1, string cell2,
            double x, double y, double w, double h,
            ChartType ct = ChartType.xlXYScatterLines,
            string title = ""
            )
        {
            if (!_Initialized) throw new InvalidOperationException("Excel is not initialized");

            this.AddChartImpl(true, sheet, cell1, cell2, x, y, w, h, ct, title);
        }


        /// <summary>
        /// Frees resources accociated with this control instance
        /// </summary>
        void IDisposable.Dispose()
        {
            try
            {
                this.Destroy();
            }
            catch (Exception) { }
        }
    }
}
