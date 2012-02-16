using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace testReport
{
    public partial class Form2 : Form
    {
        string xmlStr = "";
        public Form2(string str)
        {
            InitializeComponent();
            xmlStr = str;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            ReportDLL.ReportGenerator rg = new ReportDLL.ReportGenerator(xmlStr, "Test", Properties.Settings.Default.TemplatePath, Properties.Settings.Default.ReportPath);
            tvcReportViewer1.ReportSource = rg.ExportExcel(Properties.Settings.Default.ReportPath);

        }
    }
}
