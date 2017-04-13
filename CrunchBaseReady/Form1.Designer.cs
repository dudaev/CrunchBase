namespace CrunchBaseReady
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.uploadExcelBtn = new System.Windows.Forms.Button();
            this.excelColumnBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.clmnNames = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.xprtToXls = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // uploadExcelBtn
            // 
            this.uploadExcelBtn.Location = new System.Drawing.Point(97, 45);
            this.uploadExcelBtn.Name = "uploadExcelBtn";
            this.uploadExcelBtn.Size = new System.Drawing.Size(121, 23);
            this.uploadExcelBtn.TabIndex = 0;
            this.uploadExcelBtn.Text = "Upload";
            this.uploadExcelBtn.UseVisualStyleBackColor = true;
            this.uploadExcelBtn.Click += new System.EventHandler(this.uploadExcelBtn_Click);
            // 
            // excelColumnBtn
            // 
            this.excelColumnBtn.Location = new System.Drawing.Point(97, 133);
            this.excelColumnBtn.Name = "excelColumnBtn";
            this.excelColumnBtn.Size = new System.Drawing.Size(121, 23);
            this.excelColumnBtn.TabIndex = 1;
            this.excelColumnBtn.Text = "Bring All Data";
            this.excelColumnBtn.UseVisualStyleBackColor = true;
            this.excelColumnBtn.Click += new System.EventHandler(this.excelColumnBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(94, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "STEP 1 - Upload Excel";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(94, 90);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(200, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "STEP 2 - Which field is the company url?";
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(400, 29);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(687, 381);
            this.listBox1.TabIndex = 7;
            // 
            // clmnNames
            // 
            this.clmnNames.FormattingEnabled = true;
            this.clmnNames.Location = new System.Drawing.Point(97, 106);
            this.clmnNames.Name = "clmnNames";
            this.clmnNames.Size = new System.Drawing.Size(121, 21);
            this.clmnNames.TabIndex = 8;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(97, 194);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(121, 23);
            this.button1.TabIndex = 9;
            this.button1.Text = "Export";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // xprtToXls
            // 
            this.xprtToXls.AutoSize = true;
            this.xprtToXls.Location = new System.Drawing.Point(94, 178);
            this.xprtToXls.Name = "xprtToXls";
            this.xprtToXls.Size = new System.Drawing.Size(130, 13);
            this.xprtToXls.TabIndex = 10;
            this.xprtToXls.Text = "STEP 3 - Export your data";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1110, 444);
            this.Controls.Add(this.xprtToXls);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.clmnNames);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.excelColumnBtn);
            this.Controls.Add(this.uploadExcelBtn);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button uploadExcelBtn;
        private System.Windows.Forms.Button excelColumnBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.ComboBox clmnNames;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label xprtToXls;
    }
}

