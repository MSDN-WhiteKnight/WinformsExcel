/* WinForms Excel library 
 * Copyright (c) 2020,  MSDN.WhiteKnight (https://github.com/MSDN-WhiteKnight) 
 * License: BSD 3-Clause */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Linq;
using System.Windows.Forms;

/* WinForms Excel library demo application: Main Window */

namespace WinFormsExcelDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void bOpenFile_Click(object sender, EventArgs e)
        {
            

            OpenFileDialog ofn = new OpenFileDialog();
            if (ofn.ShowDialog(this) != DialogResult.Cancel)
            {
                try
                {
                    advancedDataGrid1.SourceFile = ofn.FileName;//load excel file
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.GetType() + ": " + ex.Message, 
                        "Error while opening file");
                }

                advancedDataGrid1.DisplayFormulaBar = cbFormulaBar.Checked;
                advancedDataGrid1.DisplayStatusBar = cbStatusBar.Checked;
                advancedDataGrid1.InitializeExcel();//load resources accosiated with this control

                this.bGenerate.Enabled = false;
                this.bOpenFile.Enabled = false;
                this.cbFormulaBar.Enabled = false;
                this.cbStatusBar.Enabled = false;
            }
        }

        private void bSaveFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog sf = new SaveFileDialog();
            if (sf.ShowDialog(this) != DialogResult.Cancel)
            {
                try
                {
                    advancedDataGrid1.SaveIntoFile(sf.FileName);//save excel file
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.GetType() + ": " + ex.Message,
                        "Error while saving file");
                }

            }
        }

        private void bGenerate_Click(object sender, EventArgs e)
        {
            advancedDataGrid1.DisplayFormulaBar = cbFormulaBar.Checked;
            advancedDataGrid1.DisplayStatusBar = cbStatusBar.Checked;
            advancedDataGrid1.InitializeExcel();//load resources accosiated with this control

            this.bGenerate.Enabled = false;
            this.bOpenFile.Enabled = false;
            this.cbFormulaBar.Enabled = false;
            this.cbStatusBar.Enabled = false;

            /*Generates table and graph for the function specified by user*/                       
            float dx = 0.1f;//argument interval
            float x_min = 0.0f, x_max = 0.0f;//argument bounds
            float x;//current argument value

            //validate input data...
            if (Single.TryParse(tbXMin.Text, out x_min) == false)
            {
                MessageBox.Show(this, "X Min is not number",
                        "Error");
                return;
            }

            if (Single.TryParse(tbXMax.Text, out x_max) == false)
            {
                MessageBox.Show(this, "X Max is not number",
                        "Error");
                return;
            }

            if (tbFunc.Text.Trim().Length == 0)
            {
                MessageBox.Show(this, "Function is not specified",
                        "Error");
                return;
            }

            //create data table and define columns
            DataTable dt = new DataTable();
            dt.TableName = tbFunc.Text;
            DataRow row;
            int nrow=2;//row number in excel
            dt.Columns.Add(new DataColumn("x", typeof(System.Single)));//argument (numeric value)
            dt.Columns.Add(new DataColumn("y", typeof(System.String)));//function (formula)

            //fill table with data
            for (x = x_min; x <= x_max; x += dx)
            {
                row = dt.NewRow();
                row[0] = x;
                //replace "x" in formula with actual Excel address
                row[1] = "="+tbFunc.Text.Replace("x", "A" + nrow.ToString());

                dt.Rows.Add(row);//add row into table
                nrow++;//increment Excel row
            }

            advancedDataGrid1.SetSheetContent(1, dt);//display table in Excel

            advancedDataGrid1.AddEmbeddedChart(1, "A1", "B" + nrow.ToString(),100,10,350,350);//add diagram
            advancedDataGrid1.Focus();

        }

        private void cbFormulaBar_CheckedChanged(object sender, EventArgs e)
        {
            //show/hide standart Excel formula bar
            advancedDataGrid1.DisplayFormulaBar = cbFormulaBar.Checked;
        }

        private void cbStatusBar_CheckedChanged(object sender, EventArgs e)
        {
            //show/hide standart Excel status bar
            advancedDataGrid1.DisplayStatusBar = cbStatusBar.Checked;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            advancedDataGrid1.Destroy();//clean up resources associated with this control
        }
    }
}
