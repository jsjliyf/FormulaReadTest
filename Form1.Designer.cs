namespace FormulaReadTest
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.button_Import = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.button_ImportAudit = new System.Windows.Forms.Button();
            this.button_ImportSum = new System.Windows.Forms.Button();
            this.button_Check = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button_ExportToExcel = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button_DeleteSum = new System.Windows.Forms.Button();
            this.comboBox_ServerIP = new System.Windows.Forms.ComboBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(12, 18);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(105, 23);
            this.comboBox1.TabIndex = 0;
            // 
            // button_Import
            // 
            this.button_Import.Location = new System.Drawing.Point(250, 12);
            this.button_Import.Name = "button_Import";
            this.button_Import.Size = new System.Drawing.Size(117, 32);
            this.button_Import.TabIndex = 1;
            this.button_Import.Text = "导入原始表";
            this.button_Import.UseVisualStyleBackColor = true;
            this.button_Import.Click += new System.EventHandler(this.Button_Import_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.tabControl1);
            this.panel1.Location = new System.Drawing.Point(12, 102);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1125, 450);
            this.panel1.TabIndex = 2;
            // 
            // tabControl1
            // 
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1125, 450);
            this.tabControl1.TabIndex = 0;
            // 
            // button_ImportAudit
            // 
            this.button_ImportAudit.Location = new System.Drawing.Point(127, 12);
            this.button_ImportAudit.Name = "button_ImportAudit";
            this.button_ImportAudit.Size = new System.Drawing.Size(117, 32);
            this.button_ImportAudit.TabIndex = 3;
            this.button_ImportAudit.Text = "更新审核表";
            this.button_ImportAudit.UseVisualStyleBackColor = true;
            this.button_ImportAudit.Click += new System.EventHandler(this.Button_ImportAudit_Click);
            // 
            // button_ImportSum
            // 
            this.button_ImportSum.Location = new System.Drawing.Point(373, 12);
            this.button_ImportSum.Name = "button_ImportSum";
            this.button_ImportSum.Size = new System.Drawing.Size(117, 32);
            this.button_ImportSum.TabIndex = 4;
            this.button_ImportSum.Text = "导入汇总表";
            this.button_ImportSum.UseVisualStyleBackColor = true;
            this.button_ImportSum.Click += new System.EventHandler(this.Button_ImportSum_Click);
            // 
            // button_Check
            // 
            this.button_Check.Location = new System.Drawing.Point(752, 9);
            this.button_Check.Name = "button_Check";
            this.button_Check.Size = new System.Drawing.Size(117, 32);
            this.button_Check.TabIndex = 5;
            this.button_Check.Text = "校验当前表";
            this.button_Check.UseVisualStyleBackColor = true;
            this.button_Check.Click += new System.EventHandler(this.Button_Check_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(875, 9);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(138, 32);
            this.button1.TabIndex = 6;
            this.button1.Text = "校验当月/季表";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button_ExportToExcel
            // 
            this.button_ExportToExcel.Location = new System.Drawing.Point(1019, 9);
            this.button_ExportToExcel.Name = "button_ExportToExcel";
            this.button_ExportToExcel.Size = new System.Drawing.Size(117, 32);
            this.button_ExportToExcel.TabIndex = 7;
            this.button_ExportToExcel.Text = "导出至Excel";
            this.button_ExportToExcel.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 47);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(930, 49);
            this.textBox1.TabIndex = 8;
            this.textBox1.Text = "公式说明";
            // 
            // button_DeleteSum
            // 
            this.button_DeleteSum.Location = new System.Drawing.Point(496, 12);
            this.button_DeleteSum.Name = "button_DeleteSum";
            this.button_DeleteSum.Size = new System.Drawing.Size(135, 32);
            this.button_DeleteSum.TabIndex = 9;
            this.button_DeleteSum.Text = "删除所有汇总表";
            this.button_DeleteSum.UseVisualStyleBackColor = true;
            this.button_DeleteSum.Click += new System.EventHandler(this.Button_DeleteSum_Click);
            // 
            // comboBox_ServerIP
            // 
            this.comboBox_ServerIP.FormattingEnabled = true;
            this.comboBox_ServerIP.Items.AddRange(new object[] {
            "10.20.66.6",
            "111.11.100.6",
            "192.168.101.108"});
            this.comboBox_ServerIP.Location = new System.Drawing.Point(979, 58);
            this.comboBox_ServerIP.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.comboBox_ServerIP.Name = "comboBox_ServerIP";
            this.comboBox_ServerIP.Size = new System.Drawing.Size(157, 23);
            this.comboBox_ServerIP.TabIndex = 10;
            this.comboBox_ServerIP.Text = "111.11.100.6";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1149, 564);
            this.Controls.Add(this.comboBox_ServerIP);
            this.Controls.Add(this.button_DeleteSum);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button_ExportToExcel);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button_Check);
            this.Controls.Add(this.button_ImportSum);
            this.Controls.Add(this.button_ImportAudit);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.button_Import);
            this.Controls.Add(this.comboBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button button_Import;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.Button button_ImportAudit;
        private System.Windows.Forms.Button button_ImportSum;
        private System.Windows.Forms.Button button_Check;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button_ExportToExcel;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button_DeleteSum;
        private System.Windows.Forms.ComboBox comboBox_ServerIP;
    }
}

