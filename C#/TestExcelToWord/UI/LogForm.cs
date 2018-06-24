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
    public partial class LogForm : Form
    {
        public LogForm(string log)
        {
            InitializeComponent();
            textBoxLog.Text = log;
        }
    }
}
