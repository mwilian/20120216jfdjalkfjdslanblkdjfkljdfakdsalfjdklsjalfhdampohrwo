using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace testReport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage2)
            {
                ReportDLL.ReportGenerator rg = new ReportDLL.ReportGenerator(GetXML(), txtCode.Text, txtTmp.Text, txtReport.Text);
                tvcReportViewer1.ReportSource = rg.ExportExcel(txtReport.Text);
                lbErr.Text = rg._sErr;
            }
        }

        private string GetXML()
        {
            string result="<DataSet>";
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    result += "<Record>";
                    foreach (DataGridViewColumn col in dataGridView1.Columns)
                    {
                        result += "<" + col.Name + ">" + row.Cells[col.Name].Value + "</" + col.Name + ">";
                    }
                    result += "</Record>";
                }
            }
            return result + "</DataSet>";
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnTmp_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtTmp.Text = folderBrowserDialog1.SelectedPath + "\\";
            }
        }

        private void btnRep_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtReport.Text = folderBrowserDialog1.SelectedPath + "\\";
            }
        }

        private void btnDesign_Click(object sender, EventArgs e)
        {
            ReportDLL.ReportGenerator rg = new ReportDLL.ReportGenerator(GetXML(), txtCode.Text, txtTmp.Text, txtReport.Text);
            rg.DesignReport();
            lbErr.Text = rg._sErr; 
        }

        private void lbErr_Click(object sender, EventArgs e)
        {
            MessageBox.Show(lbErr.Text);
        }
    }
}
