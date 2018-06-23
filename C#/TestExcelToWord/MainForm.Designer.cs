namespace TestExcelToWord
{
    partial class MainForm
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
            this.textBoxSourceFile = new System.Windows.Forms.TextBox();
            this.textBoxOutputFile = new System.Windows.Forms.TextBox();
            this.labelSourceFile = new System.Windows.Forms.Label();
            this.labelOutputFile = new System.Windows.Forms.Label();
            this.btnSelectSourceFile = new System.Windows.Forms.Button();
            this.btnSelectOutputFile = new System.Windows.Forms.Button();
            this.openSourceFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveOutputFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.saveLogFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.btnViewLog = new System.Windows.Forms.Button();
            this.btnSaveLog = new System.Windows.Forms.Button();
            this.btnTransfer = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBoxSourceFile
            // 
            this.textBoxSourceFile.Location = new System.Drawing.Point(142, 14);
            this.textBoxSourceFile.Name = "textBoxSourceFile";
            this.textBoxSourceFile.Size = new System.Drawing.Size(370, 20);
            this.textBoxSourceFile.TabIndex = 0;
            // 
            // textBoxOutputFile
            // 
            this.textBoxOutputFile.Location = new System.Drawing.Point(142, 43);
            this.textBoxOutputFile.Name = "textBoxOutputFile";
            this.textBoxOutputFile.Size = new System.Drawing.Size(370, 20);
            this.textBoxOutputFile.TabIndex = 1;
            // 
            // labelSourceFile
            // 
            this.labelSourceFile.AutoSize = true;
            this.labelSourceFile.Location = new System.Drawing.Point(12, 17);
            this.labelSourceFile.Name = "labelSourceFile";
            this.labelSourceFile.Size = new System.Drawing.Size(124, 13);
            this.labelSourceFile.TabIndex = 2;
            this.labelSourceFile.Text = "Path to source Excel file:";
            // 
            // labelOutputFile
            // 
            this.labelOutputFile.AutoSize = true;
            this.labelOutputFile.Location = new System.Drawing.Point(12, 46);
            this.labelOutputFile.Name = "labelOutputFile";
            this.labelOutputFile.Size = new System.Drawing.Size(122, 13);
            this.labelOutputFile.TabIndex = 3;
            this.labelOutputFile.Text = "Path to output Word file:";
            // 
            // btnSelectSourceFile
            // 
            this.btnSelectSourceFile.Location = new System.Drawing.Point(518, 12);
            this.btnSelectSourceFile.Name = "btnSelectSourceFile";
            this.btnSelectSourceFile.Size = new System.Drawing.Size(30, 23);
            this.btnSelectSourceFile.TabIndex = 4;
            this.btnSelectSourceFile.Text = "...";
            this.btnSelectSourceFile.UseVisualStyleBackColor = true;
            this.btnSelectSourceFile.Click += new System.EventHandler(this.btnSelectSourceFile_Click);
            // 
            // btnSelectOutputFile
            // 
            this.btnSelectOutputFile.Location = new System.Drawing.Point(518, 41);
            this.btnSelectOutputFile.Name = "btnSelectOutputFile";
            this.btnSelectOutputFile.Size = new System.Drawing.Size(30, 23);
            this.btnSelectOutputFile.TabIndex = 5;
            this.btnSelectOutputFile.Text = "...";
            this.btnSelectOutputFile.UseVisualStyleBackColor = true;
            this.btnSelectOutputFile.Click += new System.EventHandler(this.btnSelectOutputFile_Click);
            // 
            // openSourceFileDialog
            // 
            this.openSourceFileDialog.DefaultExt = "*.xlsx";
            this.openSourceFileDialog.Filter = "Excel files|*.xlsx";
            this.openSourceFileDialog.Title = "Select source Excel file";
            // 
            // saveOutputFileDialog
            // 
            this.saveOutputFileDialog.DefaultExt = "*.docx";
            this.saveOutputFileDialog.Filter = "Word files|*.docx";
            this.saveOutputFileDialog.Title = "Select output Word file";
            // 
            // saveLogFileDialog
            // 
            this.saveLogFileDialog.DefaultExt = "*.txt";
            this.saveLogFileDialog.Filter = "Text files|*.txt";
            this.saveLogFileDialog.Title = "Save log file";
            // 
            // btnViewLog
            // 
            this.btnViewLog.Location = new System.Drawing.Point(12, 97);
            this.btnViewLog.Name = "btnViewLog";
            this.btnViewLog.Size = new System.Drawing.Size(119, 23);
            this.btnViewLog.TabIndex = 6;
            this.btnViewLog.Text = "Show Log";
            this.btnViewLog.UseVisualStyleBackColor = true;
            // 
            // btnSaveLog
            // 
            this.btnSaveLog.Location = new System.Drawing.Point(137, 97);
            this.btnSaveLog.Name = "btnSaveLog";
            this.btnSaveLog.Size = new System.Drawing.Size(119, 23);
            this.btnSaveLog.TabIndex = 7;
            this.btnSaveLog.Text = "Save Log into file";
            this.btnSaveLog.UseVisualStyleBackColor = true;
            this.btnSaveLog.Click += new System.EventHandler(this.btnSaveLog_Click);
            // 
            // btnTransfer
            // 
            this.btnTransfer.Location = new System.Drawing.Point(403, 80);
            this.btnTransfer.Name = "btnTransfer";
            this.btnTransfer.Size = new System.Drawing.Size(147, 40);
            this.btnTransfer.TabIndex = 8;
            this.btnTransfer.Text = "Transfer data";
            this.btnTransfer.UseVisualStyleBackColor = true;
            this.btnTransfer.Click += new System.EventHandler(this.btnTransfer_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(562, 132);
            this.Controls.Add(this.btnTransfer);
            this.Controls.Add(this.btnSaveLog);
            this.Controls.Add(this.btnViewLog);
            this.Controls.Add(this.btnSelectOutputFile);
            this.Controls.Add(this.btnSelectSourceFile);
            this.Controls.Add(this.labelOutputFile);
            this.Controls.Add(this.labelSourceFile);
            this.Controls.Add(this.textBoxOutputFile);
            this.Controls.Add(this.textBoxSourceFile);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "TestExcelToWord";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxSourceFile;
        private System.Windows.Forms.TextBox textBoxOutputFile;
        private System.Windows.Forms.Label labelSourceFile;
        private System.Windows.Forms.Label labelOutputFile;
        private System.Windows.Forms.Button btnSelectSourceFile;
        private System.Windows.Forms.Button btnSelectOutputFile;
        private System.Windows.Forms.OpenFileDialog openSourceFileDialog;
        private System.Windows.Forms.SaveFileDialog saveOutputFileDialog;
        private System.Windows.Forms.SaveFileDialog saveLogFileDialog;
        private System.Windows.Forms.Button btnViewLog;
        private System.Windows.Forms.Button btnSaveLog;
        private System.Windows.Forms.Button btnTransfer;
    }
}

