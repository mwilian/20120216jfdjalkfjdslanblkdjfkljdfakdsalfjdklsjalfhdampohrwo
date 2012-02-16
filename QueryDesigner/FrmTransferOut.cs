using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;

namespace QueryDesigner
{
    public partial class FrmTransferOut : Form
    {
        public FrmTransferOut(string type)
        {
            InitializeComponent();
            _type = type;
            if (type == "QD")
                Text = "Query Transfer Out";
            else if (type == "QDADD")
                Text = "Query Address Transfer Out";
        }
        DataTable dt = new DataTable();
        string _type = "QD";

        public string Type
        {
            get { return _type; }
            set { _type = value; }
        }
        public string DTB = "";
        public string QD_CODE = "";
        string sErr = "";
        private void FrmTransferOut_Load(object sender, EventArgs e)
        {
            if (_type == "QD")
            {
                BUS.LIST_QDControl control = new BUS.LIST_QDControl();
                dt = control.GetTransferOut_LIST_QD(DTB, QD_CODE, ref sErr);
            }
            else if (_type == "QDADD")
            {
                BUS.LIST_QD_SCHEMAControl control = new BUS.LIST_QD_SCHEMAControl();
                dt = control.GetAll(DTB, ref sErr);
            }
            radGridView1.DataSource = dt;
            radGridView1.RetrieveStructure();
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.DefaultExt = "xml";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                dt.WriteXml(sfd.FileName);
            }
            Close();
        }
    }
}
