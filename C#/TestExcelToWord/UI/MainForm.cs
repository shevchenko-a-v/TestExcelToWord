using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestExcelToWord
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnSelectSourceFile_Click(object sender, EventArgs e)
        {
            try
            {
                openSourceFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(textBoxSourceFile.Text);
            }
            catch { }// hide all exceptions here since they are not too important

            if (openSourceFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                textBoxSourceFile.Text = openSourceFileDialog.FileName;
            }
        }

        private void btnSelectOutputFile_Click(object sender, EventArgs e)
        {
            try
            {
                saveOutputFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(textBoxOutputFile.Text);
            }
            catch { }// hide all exceptions here since they are not too important

            if (saveOutputFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                textBoxOutputFile.Text = saveOutputFileDialog.FileName;
            }
        }

        private void btnSaveLog_Click(object sender, EventArgs e)
        {
            if (saveLogFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                try
                {
                    using (var fs = new FileStream(saveLogFileDialog.FileName, FileMode.Create, FileAccess.ReadWrite))
                    using (var writer = new StreamWriter(fs))
                    {
                        writer.Write(_converter.Log);
                    }
                }
                catch
                {
                    MessageBox.Show("Could not save log file to the specified location.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnTransfer_Click(object sender, EventArgs e)
        {
            var sourceFilePath = textBoxSourceFile.Text;
            var outputFilePath = textBoxOutputFile.Text;

            try
            {
                _converter.TransferExcelToWord(sourceFilePath, outputFilePath);
                MessageBox.Show("Operation completed successfully.", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Operation has failed. See log for details.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private ExcelToWordConverter _converter { get; } = new ExcelToWordConverter();

        private void btnViewLog_Click(object sender, EventArgs e)
        {
            var logDialog = new LogForm(_converter.Log);
            logDialog.ShowDialog(this);
        }
    }
}
