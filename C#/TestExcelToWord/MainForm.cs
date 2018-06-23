using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
                // TODO: Save log to file
            }
        }

        private void btnTransfer_Click(object sender, EventArgs e)
        {
            var sourceFilePath = textBoxSourceFile.Text;
            var outputFilePath = textBoxOutputFile.Text;

            _converter.TransferExcelToWord(sourceFilePath, outputFilePath);
        }

        private ExcelToWordConverter _converter { get; } = new ExcelToWordConverter();
    }
}
