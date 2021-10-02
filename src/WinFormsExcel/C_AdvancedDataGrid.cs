/* WinForms Excel library 
 * Copyright (c) 2021,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;
/*WinForms Excel Library - C_AdvancedDataGrid user control
 http://smallsoft2.blogspot.ru/
 */

namespace ExtraControls
{
    /// <summary>
    /// Wrapper control for AdvancedDataGrid/S_AdvancedDataGrid controls. Switches between them at runtime, allowing 
    /// to use required control depending on Excel presence on target machine
    /// </summary>
    public partial class C_AdvancedDataGrid : UserControl, IDisposable, IDataGrid
    {
        /// <summary>
        /// Specifies the underlying grid control used to actually display data
        /// </summary>
        protected IDataGrid basegrid;

        /// <summary>
        /// Indicates that this control uses Excel mode (backed by AdvancedDataGrid)
        /// </summary>
        protected bool excel_on;

        /// <summary>
        /// Indicates that excel was loaded for this control instance
        /// </summary>
        protected bool initialized;//underlying grid is initialized        

        /// <summary>
        /// Displays error message within this control
        /// </summary>
        /// <param name="s">Error description</param>
        protected void ErrorMessage(string s)
        {
            TextBox l = new TextBox();
            l.Text = s;
            l.Select(0, 0); l.AutoSize = true;
            l.ReadOnly = true; l.Multiline = true;
            l.Dock = DockStyle.Fill;
            this.Controls.Add(l);
        }

        /// <summary>
        /// Creates C_AdvancedDataGrid control in substitute mode
        /// </summary>
        public C_AdvancedDataGrid()
        {
            InitializeComponent();
            basegrid = null;
            excel_on = false;
            
            initialized = false;
            try
            {
                this.Initialize(GridMode.Substitute);
            }
            catch (Exception ex)
            {
                this.ErrorMessage("Failed to initialize the control. " + Environment.NewLine + ex.ToString());                
            }
        }

        /// <summary>
        /// Creates C_AdvancedDataGrid control in specified mode
        /// </summary>
        public C_AdvancedDataGrid(GridMode mode)
        {
            InitializeComponent();
            basegrid = null;
            excel_on = false;
            initialized = false;
            
            try
            {
                this.Initialize(mode);
            }
            catch (Exception ex)
            {
                this.ErrorMessage("Failed to initialize the control. " + Environment.NewLine + ex.ToString()); 
            }
        }

        /// <summary>
        /// Gets or sets the current grid mode. Setting this property will cause control to be initialized again.
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public GridMode Mode
        {
            set
            {
                try
                {
                    Initialize(value);
                }
                catch (Exception) { ;}
            }

            get
            {
                if (!initialized) return GridMode.Undefined;

                if (excel_on) return GridMode.Excel;
                else return GridMode.Substitute;
            }

        }

        /// <summary>
        /// Returns the underlying grid, which type is determined by Mide property 
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public IDataGrid BaseGrid
        {  
            get
            {
                return this.basegrid;
            }
        }

        /// <summary>
        /// Initializes this control in specified mode
        /// </summary>        
        public void Initialize(GridMode mode)
        {
            try
            {
                //clean up resources of current control
                this.Controls.Clear();                

                if (initialized && excel_on)
                {
                    (basegrid as AdvancedDataGrid).Destroy();
                    
                    excel_on = false;
                }
                basegrid = null;
                initialized = false;

                //initialize control based on chosen mode
                switch (mode)
                {
                    case GridMode.Substitute:
                        basegrid = new SubstituteGrid.S_AdvancedDataGrid();
                        excel_on = false;
                        this.Controls.Add((Control)basegrid);
                        (basegrid as Control).Dock = DockStyle.Fill;
                        initialized = true;
                        break;
                    case GridMode.Excel:
                        basegrid = new AdvancedDataGrid();
                        (basegrid as AdvancedDataGrid).InitializeExcel();
                        excel_on = true;
                        this.Controls.Add((Control)basegrid);
                        (basegrid as Control).Dock = DockStyle.Fill;
                        initialized = true;
                        break;
                    case GridMode.Auto:
                        try
                        {
                            //tries to initialize excel mode
                            basegrid = new AdvancedDataGrid();
                            (basegrid as AdvancedDataGrid).InitializeExcel();
                            excel_on = true;
                        }
                        catch (Exception)
                        {
                            //on fail, tries to use substitute mode
                            basegrid = new SubstituteGrid.S_AdvancedDataGrid();
                            excel_on = false;
                        }
                        this.Controls.Add((Control)basegrid);
                        (basegrid as Control).Dock = DockStyle.Fill;
                        initialized = true;
                        break;                    
                }

                
            }
            catch (Exception)
            {
                this.excel_on = false;
                this.initialized = false;
                this.basegrid = null;
                throw;
            }

        }

        #region RESOURCE CLEANUP METHODS

        /// <summary>
        /// Cleans up resources used by this control. The control will be in uninitialized state.
        /// </summary>
        public void Destroy()
        {
            if (initialized && excel_on)
            {                
                ((AdvancedDataGrid)basegrid).Destroy();
                initialized = false;
                excel_on = false;
            }
        }

        void IDisposable.Dispose()
        {
            try { this.Destroy(); }
            catch (Exception) { ;}
        }

        /// <summary>
        /// Destroys this control instance
        /// </summary>
        ~C_AdvancedDataGrid()
        {
            try { this.Destroy(); }
            catch (Exception) { ;}
        }
        #endregion

        /*Underlying grid interface wrappers*/
        ///<inheritdoc/>
        public object DataSource
        {
            get
            {
                return basegrid.DataSource;
            }
            set
            {
                basegrid.DataSource = value;
            }
        }

        ///<inheritdoc/>
        public int ActiveSheet
        {
            get
            {
                return basegrid.ActiveSheet;
            }
            set
            {
                basegrid.ActiveSheet = value;
            }
        }

        ///<inheritdoc/>
        public int SheetsCount
        {
            get { return basegrid.SheetsCount; }
        }

        ///<inheritdoc/>
        public void SetCellContent(int sheet, int row, int col, object val)
        {
            basegrid.SetCellContent(sheet, row, col, val);
        }

        ///<inheritdoc/>
        public object GetCellContent(int sheet, int row, int col)
        {
            return basegrid.GetCellContent(sheet, row, col);
        }

        ///<inheritdoc/>
        public void SetSheetContent(int sheet, DataTable t)
        {
            basegrid.SetSheetContent(sheet, t);
        }

        ///<inheritdoc/>
        public DataTable GetSheetContent(int sheet, bool FirstRowHasHeaders, int n_col = 0, int n_row = 0)
        {
            return basegrid.GetSheetContent(sheet, FirstRowHasHeaders, n_col, n_row);
        }

        ///<inheritdoc/>
        public int GetActiveSheet()
        {
            return basegrid.GetActiveSheet();
        }

        ///<inheritdoc/>
        public void SetActiveSheet(int index)
        {
            basegrid.SetActiveSheet(index);
        }

        ///<inheritdoc/>
        public void DeleteSheet(int index)
        {
            basegrid.DeleteSheet(index);
        }

        ///<inheritdoc/>
        public void AddSheet(string name = "")
        {
            basegrid.AddSheet(name);
        }

        ///<inheritdoc/>
        public List<XlSheet> GetSheets()
        {
            return basegrid.GetSheets();
        }

        ///<inheritdoc/>
        public void SetSheetName(int sheet, string name)
        {
            basegrid.SetSheetName(sheet, name);
        }

        ///<inheritdoc/>
        public string GetSheetName(int sheet)
        {
            return basegrid.GetSheetName(sheet);
        }

        ///<inheritdoc/>
        public int FindSheet(string name)
        {
            return basegrid.FindSheet(name);
        }

        ///<inheritdoc/>
        public void NewEmptyWorkbook()
        {
            basegrid.NewEmptyWorkbook();
        }
    }
}
