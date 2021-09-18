/* WinForms Excel library 
 * Copyright (c) 2020,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using ExtraControls;
using ExtraControls.SubstituteGrid;

/*AdvancedDataGrid control demonstration window*/

namespace MyExcel
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
            advancedDataGrid1.InitializeExcel();

            //advancedDataGrid1.OpenFile("e:\\coal_x10.xls");

            advancedDataGrid1.SetCellContent(1, 1, 3, "s");//set cell value directly

            //create table of Sin math function
            DataTable t = new DataTable("1");
            DataRow row;
            t.Columns.Add("x", typeof(System.Single));
            t.Columns.Add("sin(x)",typeof(System.Single));

            for (float x = 0.0f; x <= 2.0f; x += 0.1f)
            {
                row = t.NewRow();
                row[0] = x;
                row[1] = Math.Sin(x);
                t.Rows.Add(row);
            }

            //change active sheet
            advancedDataGrid1.SetActiveSheet(1);

            //display table in AdvancedDataGridControl
            advancedDataGrid1.DataSource = t;

            //display diagram in the workbook

            advancedDataGrid1.AddChart(1, 
                advancedDataGrid1.GetCellAddress(1, 1, 1),
                advancedDataGrid1.GetCellAddress(1, 20, 2),  
                ChartType.xlXYScatterLines, "sin(x)");
            
            

            //add one more sheet to the workbook            
            advancedDataGrid1.AddSheet();
            advancedDataGrid1.SetActiveSheet(2);

            //change sheet names
            for (int i = 1; i <= advancedDataGrid1.SheetsCount; i++)
            {
                advancedDataGrid1.SetSheetName(i, "Sheet #"+i.ToString());
            }

            //change the order of sheets in workbook
            //advancedDataGrid1.MoveSheet(2,3,false);                        

            //load sheet's content as DataTable object
            DataTable tt = advancedDataGrid1.GetSheetContent(1, true);            
            advancedDataGrid1.SetSheetContent(4, tt);

            //place formula in specified cell so its value will be calculated
            advancedDataGrid1.SetCellContent(4,3, 3, "=A3+B3");

            //display calculation result in message box
            MessageBox.Show(advancedDataGrid1.GetCellContent(4,3,3).ToString());            

            //save workbook into file
            advancedDataGrid1.SaveIntoFile("c:\\Test\\file.xls");
            
            List<XlSheet> list = advancedDataGrid1.GetSheets();
            string s = "";
            foreach (XlSheet sh in list)
            {
                s += sh.Index.ToString() + ": " + sh.Name + Environment.NewLine;
            }
            //MessageBox.Show(s);
            advancedDataGrid1.SetActiveSheet(3);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            advancedDataGrid1.Destroy();//no longer need the Excel            
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            
        }

       
    }

    


}
