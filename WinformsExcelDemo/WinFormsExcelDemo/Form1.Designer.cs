namespace WinFormsExcelDemo
{
    partial class Form1
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.advancedDataGrid1 = new ExtraControls.AdvancedDataGrid();
            this.bOpenFile = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.bGenerate = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.tbXMax = new System.Windows.Forms.TextBox();
            this.tbXMin = new System.Windows.Forms.TextBox();
            this.tbFunc = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.bSaveFile = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cbFormulaBar = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // advancedDataGrid1
            // 
            this.advancedDataGrid1.ActiveSheet = -1;
            this.advancedDataGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.advancedDataGrid1.BackColor = System.Drawing.SystemColors.ControlDark;
            this.advancedDataGrid1.DataSource = null;
            this.advancedDataGrid1.Location = new System.Drawing.Point(215, 22);
            this.advancedDataGrid1.Name = "advancedDataGrid1";
            this.advancedDataGrid1.Size = new System.Drawing.Size(286, 336);
            this.advancedDataGrid1.TabIndex = 0;
            // 
            // bOpenFile
            // 
            this.bOpenFile.Location = new System.Drawing.Point(23, 182);
            this.bOpenFile.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.bOpenFile.Name = "bOpenFile";
            this.bOpenFile.Size = new System.Drawing.Size(68, 26);
            this.bOpenFile.TabIndex = 3;
            this.bOpenFile.Text = "Load file";
            this.bOpenFile.UseVisualStyleBackColor = true;
            this.bOpenFile.Click += new System.EventHandler(this.bOpenFile_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.bGenerate);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.tbXMax);
            this.groupBox1.Controls.Add(this.tbXMin);
            this.groupBox1.Controls.Add(this.tbFunc);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(16, 22);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(182, 146);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Function table and graph";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 79);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(34, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "x max";
            // 
            // bGenerate
            // 
            this.bGenerate.Location = new System.Drawing.Point(55, 106);
            this.bGenerate.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.bGenerate.Name = "bGenerate";
            this.bGenerate.Size = new System.Drawing.Size(76, 26);
            this.bGenerate.TabIndex = 3;
            this.bGenerate.Text = "Generate";
            this.bGenerate.UseVisualStyleBackColor = true;
            this.bGenerate.Click += new System.EventHandler(this.bGenerate_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 52);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(31, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "x min";
            // 
            // tbXMax
            // 
            this.tbXMax.Location = new System.Drawing.Point(55, 76);
            this.tbXMax.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tbXMax.Name = "tbXMax";
            this.tbXMax.Size = new System.Drawing.Size(109, 20);
            this.tbXMax.TabIndex = 1;
            this.tbXMax.Text = "1";
            // 
            // tbXMin
            // 
            this.tbXMin.Location = new System.Drawing.Point(55, 50);
            this.tbXMin.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tbXMin.Name = "tbXMin";
            this.tbXMin.Size = new System.Drawing.Size(109, 20);
            this.tbXMin.TabIndex = 1;
            this.tbXMin.Text = "0";
            // 
            // tbFunc
            // 
            this.tbFunc.Location = new System.Drawing.Point(55, 20);
            this.tbFunc.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tbFunc.Name = "tbFunc";
            this.tbFunc.Size = new System.Drawing.Size(109, 20);
            this.tbFunc.TabIndex = 1;
            this.tbFunc.Text = "sin(x)";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 23);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "y(x)=";
            // 
            // bSaveFile
            // 
            this.bSaveFile.Location = new System.Drawing.Point(112, 182);
            this.bSaveFile.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.bSaveFile.Name = "bSaveFile";
            this.bSaveFile.Size = new System.Drawing.Size(68, 26);
            this.bSaveFile.TabIndex = 3;
            this.bSaveFile.Text = "Save file";
            this.bSaveFile.UseVisualStyleBackColor = true;
            this.bSaveFile.Click += new System.EventHandler(this.bSaveFile_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cbFormulaBar);
            this.groupBox2.Location = new System.Drawing.Point(20, 224);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Size = new System.Drawing.Size(178, 119);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Appearance";
            // 
            // cbFormulaBar
            // 
            this.cbFormulaBar.AutoSize = true;
            this.cbFormulaBar.Location = new System.Drawing.Point(14, 26);
            this.cbFormulaBar.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cbFormulaBar.Name = "cbFormulaBar";
            this.cbFormulaBar.Size = new System.Drawing.Size(106, 17);
            this.cbFormulaBar.TabIndex = 0;
            this.cbFormulaBar.Text = "show formula bar";
            this.cbFormulaBar.UseVisualStyleBackColor = true;
            this.cbFormulaBar.CheckedChanged += new System.EventHandler(this.cbFormulaBar_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(532, 384);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.bSaveFile);
            this.Controls.Add(this.bOpenFile);
            this.Controls.Add(this.advancedDataGrid1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private ExtraControls.AdvancedDataGrid advancedDataGrid1;
        private System.Windows.Forms.Button bOpenFile;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button bGenerate;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbXMax;
        private System.Windows.Forms.TextBox tbXMin;
        private System.Windows.Forms.TextBox tbFunc;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button bSaveFile;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox cbFormulaBar;
    }
}

