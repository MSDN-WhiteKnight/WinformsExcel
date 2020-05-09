/* WinForms Excel library 
 * Copyright (c) 2020,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
/*Windows Forms Excel Library - S_AdvancedDataGrid user control
 http://smallsoft2.blogspot.ru/
 */

namespace ExtraControls.SubstituteGrid
{
    /// <summary>
    /// Provides an implementation of IDataGrid interface via standard Windows Forms control, allowing to display and edit one or
    /// more tables of data. Serves as a Excel-free substitute for AdvancedDataGrid control.
    /// </summary>
    public partial class S_AdvancedDataGrid : UserControl, IDataGrid
    {
        #region PROTECTED VARIABLES
        /// <summary>
        /// The collection of sheets displayed in this control
        /// </summary>
        protected List<S_DataSheet> sheets = new List<S_DataSheet>(6);

        /// <summary>
        /// Reference to the sheet currently displayed in the grid
        /// </summary>
        protected S_DataSheet active_sheet = null;
        #endregion

        #region INFRASTRUCTURE FUNCTIONS
        /// <summary>
        /// Event fired when user clicks a button associated with certain sheet
        /// </summary>        
        private void button_Click(object sender, EventArgs e)
        {
            Button butt = (Button)sender;
            if (butt == null) return;
            S_DataSheet sh = (S_DataSheet)butt.Tag;
            if (sh == null) return;
            ActivateSheet(sh);//activate sheet associated with this button
        }

        /// <summary>
        /// Creates the new sheet and adds it to this data grid control
        /// </summary>
        /// <param name="name">Optional sheet name</param>
        protected void NewSheet(string name="")
        {
            if (sheets == null) sheets = new List<S_DataSheet>(6);
            if (name == null) name = "";

            S_DataSheet sh = new S_DataSheet();
            int c = sheets.Count+1;
            DataColumn col;
            DataRow row;

            if (name.Trim().Length == 0)
                sh.name = "Sheet " + c.ToString();
            else 
                sh.name = name;

            //create table with empty rows an columns
            sh.data = new DataTable(sh.name);
            for (int i = 1; i <= 50; i++)
            {
                col = new DataColumn(i.ToString());
                sh.data.Columns.Add(col);
            }

            for (int i = 1; i <= 200; i++)
            {
                row = sh.data.NewRow();
                sh.data.Rows.Add(row);
            }

            //create title button associated with this sheet and add it into the grid
            sh.title_button = new Button();
            sh.title_button.Text = sh.name;
            sh.title_button.Height = flowLayoutPanel1.Height - 4;
            sh.title_button.Width = 70;            
            sh.title_button.Click += button_Click;
            sh.title_button.Tag = sh;
            this.sheets.Add(sh);
            flowLayoutPanel1.Controls.Add(sh.title_button);
        }

        /// <summary>
        /// Displays specified sheet in the grid
        /// </summary>
        /// <param name="sh">Reference to sheet object</param>
        protected void ActivateSheet(S_DataSheet sh)
        {
            if (sh == null) return;
            active_sheet = sh;

            foreach (Control c in flowLayoutPanel1.Controls)//make all sheets' apperence standart
            {
                if (c is Button)
                {
                    (c as Button).FlatStyle = FlatStyle.Standard;
                    (c as Button).ForeColor = Color.Black;
                }
            }
            
            //highlight active sheet
            sh.title_button.FlatStyle = FlatStyle.Flat;
            sh.title_button.ForeColor = Color.Blue;

            //adjust data binding of the underlying DataGridView control
            dataGridView1.AutoGenerateColumns = true;
            if(sh.data!=null)dataGridView1.DataSource = sh.data;
        }

        /// <summary>
        /// Chenges the name of specified sheet, adjusting UI
        /// </summary>
        /// <param name="i">internal sheet index</param>
        /// <param name="name">new name</param>
        protected void UpdateSheetName(int i, string name)
        {
            sheets[i].name = name;
            sheets[i].title_button.Name = name;
        }
        #endregion

        /// <summary>
        /// Creates new S_AdvancedDataGrid control and fills it with default empty content
        /// </summary>
        public S_AdvancedDataGrid()
        {
            InitializeComponent();
            NewEmptyWorkbook();
        }

        /// <summary>
        /// Sets contents of the cell specified by sheet, row and column numbers into an object of any type
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="row">Row number</param>
        /// <param name="col">Column number</param>
        /// <param name="val">New cell value</param>
        public void SetCellContent(int sheet, int row, int col, object val)
        {
            if (sheet <= 0) throw new ArgumentException("sheet number must be positive");
            if (row <= 0) throw new ArgumentException("row number must be positive");
            if (col <= 0) throw new ArgumentException("column number must be positive");

            if (sheet > sheets.Count) throw new ArgumentException("sheet number exceeds collection size");

            //if cell exceeds the boundaries of existing data, add more rows or columns
            if (row - 1 >= sheets[sheet - 1].data.Rows.Count)
            {
                for (int i = 0; i <= row - sheets[sheet - 1].data.Rows.Count; i++)
                {
                    sheets[sheet - 1].data.Rows.Add(sheets[sheet - 1].data.NewRow());
                }
            }

            if (col - 1 >= sheets[sheet - 1].data.Columns.Count)
            {
                for (int i = sheets[sheet - 1].data.Columns.Count+1; i <= col; i++)
                {
                    sheets[sheet - 1].data.Columns.Add((i).ToString());
                }
            }

            sheets[sheet - 1].data.Rows[row - 1][col - 1] = val;
            if (this.active_sheet == sheets[sheet - 1]) ActivateSheet(this.active_sheet);
        }

        /// <summary>
        /// Gets the content of the cell specified by sheet, row and column numbers.
        /// 
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="row">Row number</param>
        /// <param name="col">Column number</param>
        /// <returns>Value of the specified cell</returns>
        public object GetCellContent(int sheet, int row, int col)
        {
            if (sheet <= 0) throw new ArgumentException("sheet number must be positive");
            if (row <= 0) return null;
            if (col <= 0) return null;

            if (sheet > sheets.Count) throw new ArgumentException("sheet number exceeds collection size");
            if (row - 1 >= sheets[sheet - 1].data.Rows.Count)
            {
                return null;
            }

            if (col - 1 >= sheets[sheet - 1].data.Columns.Count)
            {
                return null;
            }

            return sheets[sheet - 1].data.Rows[row - 1][col - 1];
        }

        /// <summary>
        /// Fills the specified sheet with a content of given DataTable object
        ///          
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="t">DataTable to fill sheet's content</param>
        public void SetSheetContent(int sheet, DataTable t)
        {
            if (sheet <= 0) throw new ArgumentException("sheet number must be positive");
            
            sheets[sheet - 1].data = t;
            if (this.active_sheet == sheets[sheet - 1]) ActivateSheet(this.active_sheet);
        }

        /// <summary>
        /// Loads content of specified sheet as DataTable object.        
        ///        
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="FirstRowHasHeaders">Specifies that first row contains column headers (unused)</param>
        /// <param name="n_col">Maximum number of columns to load (unused)</param>
        /// <param name="n_row">Maximum number of rows to load (unused)</param>
        /// <returns>DataTable object filled with sheet content</returns>
        public DataTable GetSheetContent(int sheet, bool FirstRowHasHeaders, int n_col = 0, int n_row = 0)
        {
            if (sheet <= 0) throw new ArgumentException("sheet number must be positive");
            if (n_col < 0) throw new ArgumentException("n_col must be non-negative");
            if (n_row < 0) throw new ArgumentException("n_row must be non-negative");
            
            return sheets[sheet - 1].data;
        }

        /// <summary>
        /// Gets the number of currently active sheet.
        /// Returns -1 if excel is not initialized.
        /// </summary>
        /// <returns>Sheet number</returns>
        public int GetActiveSheet()
        {
            return sheets.IndexOf(active_sheet) + 1;
        }

        /// <summary>
        /// Activates specified sheet in this control instance
        ///        
        /// </summary>
        /// <param name="index">Sheet number</param>
        public void SetActiveSheet(int index)
        {
            if (index <= 0) throw new ArgumentException("sheet number must be positive");
            
            this.active_sheet = sheets[index - 1];
        }

        /// <summary>
        /// Removes specified sheet from workbook. 
        /// Note: You can't remove all sheets. At least one sheet must be present in workbook all the time.
        ///         
        /// </summary>
        /// <param name="index">Sheet number</param>
        public void DeleteSheet(int index)
        {
            if (index <= 0) throw new ArgumentException("sheet number must be positive");
            
            if(sheets.Count<=1)return;

            if (active_sheet == sheets[index - 1])
            {
                int i=index-2;
                if (i >= 0) ActivateSheet(sheets[i]);
            }
            sheets.RemoveAt(index - 1);
        }

        /// <summary>
        /// Adds new sheet into the workbook of this control instance
        ///         
        /// </summary>
        /// <param name="name">Worksheet name (optional)</param>
        public void AddSheet(string name = "")
        {
            NewSheet(name);
        }

        /// <summary>
        /// Gets names of all sheets in this control instance as a list of string objects.
        ///        
        /// </summary>
        /// <returns>List of worksheet names</returns>
        public List<XlSheet> GetSheets()
        {
            List<XlSheet> sh = new List<XlSheet>(sheets.Count);
            XlSheet x;

            foreach (S_DataSheet s in sheets)
            {
                x = new XlSheet();
                x.Name = s.name;
                x.Index = sheets.IndexOf(s);
                x.IsChart = false;
                sh.Add(x);
            }
            return sh;
        }

        /// <summary>
        /// Changes the name of specified sheet.
        ///         
        /// </summary>
        /// <param name="sheet">Sheet number</param>
        /// <param name="name">Sheet ne name</param>
        public void SetSheetName(int sheet, string name)
        {
            if (sheet <= 0) throw new ArgumentException("sheet number must be positive");
            if (name==null) throw new ArgumentNullException("name must be specified");
            

            UpdateSheetName(sheet - 1, name);
        }

        /// <summary>
        /// Gets the index of sheet with specified name. Returns -1 if the sheet is not found.
        ///         
        /// </summary>
        /// <param name="name">Sheet name</param>
        /// <returns>Sheet index or -1</returns>
        public string GetSheetName(int sheet)
        {
            if (sheet <= 0) throw new ArgumentException("sheet number must be positive");
            
            return sheets[sheet - 1].name;
        }

        /// <summary>
        /// Inserts the worksheet before or after the target worksheet in the workbook
        ///          
        /// </summary>
        /// <param name="curr_index">Number of sheet to move</param>
        /// <param name="new_index">Target sheet index</param>
        /// <param name="before">Specifies that sheet must be placed before the target sheet</param>
        public int FindSheet(string name)
        {
            foreach (S_DataSheet s in sheets)
            {
                if (s.name == name) return sheets.IndexOf(s);
            }
            return -1;
        }
        

        public void NewEmptyWorkbook()
        {
            for (int i = 1; i <= 3; i++)
            {
                NewSheet();
            }
            ActivateSheet(sheets[0]);
        }

        public object DataSource
        {
            set
            {                

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
                try
                {
                    return this.GetSheetContent(this.GetActiveSheet(), true);
                }
                catch (Exception) { return null; }
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
                
                try
                {
                    return this.GetActiveSheet();
                }
                catch (Exception) { return -1; }
            }
            set
            {
                
                try
                {
                    this.SetActiveSheet(value);
                }
                catch (Exception) { return; }
            }
        }

        public int SheetsCount
        {
            get
            { 
                try
                {                   
                    int c = sheets.Count;
                    return c;

                }
                catch (Exception)
                {
                    return -1;
                }                

            }
        }

        [Category("Behavior"),Browsable(true),EditorBrowsable(EditorBrowsableState.Always)]
        [Description("Specifies if user is allowed to edit data in this control"),DefaultValue(false)]
        public bool ReadOnly
        {
            get { return dataGridView1.ReadOnly; }
            set { dataGridView1.ReadOnly = value; }
        }

    }
}
