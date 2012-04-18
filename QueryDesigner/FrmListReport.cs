using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


namespace dCube
{
    public partial class FrmListReport :Form
    {
        public FrmListReport()
        {
            InitializeComponent();
        }
        public FrmListReport(DataTable dt)
        {
            InitializeComponent();
            _dataTable = dt;
        }
        DataTable _dataTable;

        public DataTable DataTable
        {
            get { return _dataTable; }
            set { _dataTable = value; }
        }
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void FrmListReport_Load(object sender, EventArgs e)
        {
            dgvResult.DataSource = _dataTable;
            dgvResult.RetrieveStructure();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {

        }

        private void dgvResult_FormattingRow(object sender, Janus.Windows.GridEX.RowLoadEventArgs e)
        {

        }
    }
}
