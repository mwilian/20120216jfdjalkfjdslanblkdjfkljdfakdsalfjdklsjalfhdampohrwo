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
        public FrmTransferOut(string db, string type)
        {
            InitializeComponent();
            _type = type;
            _db = db;
            if (type == "QD")
                Text = "Query Transfer Out";
            else if (type == "QDADD")
                Text = "Query Address Transfer Out";
            else if (type == "TASK")
                Text = "Task Transfer Out";
        }
        DataTable dt = new DataTable();
        DataTable dtEnd = new DataTable();
        string _type = "QD";
        string _db = "";
        public string Type
        {
            get { return _type; }
            set { _type = value; }
        }
        public string QD_CODE = "";
        string sErr = "";
        private void FrmTransferOut_Load(object sender, EventArgs e)
        {
            if (_type == "QD")
            {
                BUS.LIST_QDControl control = new BUS.LIST_QDControl();
                dt = control.GetTransferOut_LIST_QD(_db, QD_CODE, ref sErr);

            }
            else if (_type == "QDADD")
            {
                BUS.LIST_QD_SCHEMAControl control = new BUS.LIST_QD_SCHEMAControl();
                dt = control.GetAll(_db, ref sErr);
            }
            else if (_type == "TASK")
            {
                BUS.LIST_TASKControl control = new BUS.LIST_TASKControl();
                dt = control.GetAll(_db, ref sErr);
            }
            dtEnd = dt.Copy();
            radGridView1.DataSource = dtEnd;
            radGridView1.RetrieveStructure();
        }

        private void radButton1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.DefaultExt = "xml";
            if (sfd.ShowDialog() == DialogResult.OK)
            {

                dtEnd.WriteXml(sfd.FileName);
            }
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filterFrom = "";
            string filterTo = "";

            DataRow[] rows = null;
            if (To.Text == "")
            {
                if (_type == "QD")
                {
                    rows = dt.Select("QD_ID ='" + From.Text + "'");
                }
                else if (_type == "QDADD")
                {
                    rows = dt.Select("SCHEMA_ID ='" + From.Text + "'");
                }
                else if (_type == "TASK")
                {
                    rows = dt.Select("CODE ='" + From.Text + "'");
                }
            }
            else
            {
                if (_type == "QD")
                {
                    rows = dt.Select("QD_ID >='" + From.Text + "'" + "and  QD_ID <='" + To.Text + "'");
                }
                else if (_type == "QDADD")
                {
                    rows = dt.Select("SCHEMA_ID >='" + From.Text + "'" + "and  SCHEMA_ID <='" + To.Text + "'");
                }
                else if (_type == "TASK")
                {
                    rows = dt.Select("CODE >='" + From.Text + "'" + "and  CODE <='" + To.Text + "'");
                }
            }
            if (rows != null || rows.Length > 0)
            {
                dtEnd = dt.Clone();
                for (int i = 0; i < rows.Length; i++)
                {
                    dtEnd.ImportRow(rows[i]);
                }
            }
            radGridView1.DataSource = dtEnd;

        }
    }
}
