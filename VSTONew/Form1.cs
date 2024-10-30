using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;


namespace VSTONew
{
    public partial class Form1 : Form
    {
        public Excel.Application excelApp;
        public Form1()
        {
            InitializeComponent();
            excelApp = Globals.ThisAddIn.Application;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
          

        }

        private void button1_Click(object sender, EventArgs e)
        {

            excelApp.ActiveCell.Value = this.textBox1.Text;

        }
    }
}
