using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace testReport
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2(richTextBox1.Text);
            frm.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ReportDLL.ReportGenerator rg = new ReportDLL.ReportGenerator(richTextBox1.Text, "Test", Properties.Settings.Default.TemplatePath, Properties.Settings.Default.ReportPath);
            rg.DesignReport();
        }
    }
}
