using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;

namespace dCube
{
    public partial class frmChartPro : Form
    {
        clsChartProperty _property = new clsChartProperty();

        public clsChartProperty ReturnProperty
        {
            get { return _property; }
            set { _property = value; }
        }
        public frmChartPro(clsChartProperty property)
        {
            InitializeComponent();
            _property = property;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void frmChartPro_Load(object sender, EventArgs e)
        {
            pgChart.SelectedObject = _property;
        }
    }
}
