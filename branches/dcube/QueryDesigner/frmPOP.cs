using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace dCube
{
    public partial class frmPOP : Form
    {
        BUS.POPControl ctr = new BUS.POPControl();
        string _sErr = "";
        string _processStatus = "";
        public frmPOP()
        {
            InitializeComponent();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            RefreshForm("");
            EnabledForm(true);
            _processStatus = "C";
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            frmPOPView frm = new frmPOPView();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                SetDataToForm(ctr.Get(frm.Code, frm.Lookup, ref _sErr));
                EnabledForm(false);
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (ctr.IsExist(txtCode.Text, txtDB.Text))
            {
                EnabledForm(true);
                txtCode.Enabled = false;
                _processStatus = "A";
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //string sErr = "";

            DTO.POPInfo inf = new DTO.POPInfo();

            if (_processStatus == "C")
            {
                if (!ctr.IsExist(txtCode.Text, txtDB.Text))
                {
                    ctr.Add(GetDataFromForm(inf), ref _sErr);
                }
                else
                    _sErr = txtCode.Text.Trim() + " is exist!";
            }
            else if (_processStatus == "A")
            {
                _sErr = ctr.InsertUpdate(GetDataFromForm(inf));
            }
            if (_sErr == "")
            {
                _processStatus = "V";
                EnabledForm(false);
            }
            else lbErr.Text = _sErr;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (ctr.IsExist(txtCode.Text, txtDB.Text))
            {
                _sErr = ctr.Delete(txtCode.Text, txtDB.Text);
                RefreshForm("");
                EnabledForm(false);
                _processStatus = "";
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (ctr.IsExist(txtCode.Text, txtDB.Text))
            {
                EnabledForm(true);
                txtCode.Focus();
                _processStatus = "C";
            }
        }


        private void frmPOD_Load(object sender, EventArgs e)
        {
            RefreshForm("");
            EnabledForm(false);
            DataTable dt = new DataTable("POP");
            DataColumn[] columns = new DataColumn[] { 
                new DataColumn("ID",typeof(int)), 
                new DataColumn("Module"), 
                new DataColumn("Function"), 
                new DataColumn("Action"), 
                new DataColumn("Permission"),
                new DataColumn("Address")
            };
            dt.Columns.AddRange(columns);
            //string fileName = Application.StartupPath + "\\Configuration\\Security.xml";
            //XmlReader xmlRead = new StringReader(Properties.Resources.Security);
            StringReader reader = new StringReader(Properties.Resources.Security);
            dt.ReadXml(reader);
            dgvData.DataSource = dt;
        }
        private void lbErr_Click(object sender, EventArgs e)
        {
            if (lbErr.Text != "" && lbErr.Text != "...")
                MessageBox.Show(lbErr.Text);
        }
        private void SetDataToForm(DTO.POPInfo obj)
        {
            txtCode.Text = obj.ROLE_ID;
            txtDB.Text = obj.DB;
            txtDefault.Text = obj.DEFAULT_VALUE;
            DataTable dt = dgvData.DataSource as DataTable;
            dt.DefaultView.Sort = "ID ASC";

            if (obj.PERMISSION != "")
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["Permission"] = obj.PERMISSION[i + 1].ToString();

                }
            else
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["Permission"] = "";

                }
            }

        }
        private DTO.POPInfo GetDataFromForm(DTO.POPInfo obj)
        {
            //DTO.PODInfo obj = new DTO.PODInfo();
            obj.ROLE_ID = txtCode.Text;
            obj.DB = txtDB.Text;
            obj.DEFAULT_VALUE = txtDefault.Text;
            string permis = obj.DEFAULT_VALUE;
            DataTable dt = dgvData.DataSource as DataTable;
            dt.DefaultView.Sort = "ID ASC";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string tmp = dt.Rows[i]["Permission"].ToString().Trim();
                permis += tmp != "" ? tmp : " ";
            }
            obj.PERMISSION = permis;
            return obj;
        }
        private void EnabledForm(bool value)
        {
            pContain.Enabled = value;
            txtCode.Enabled = true;
        }
        private void RefreshForm(string value)
        {
            foreach (Control x in pContain.Controls)
            {
                if (x is TextBox)
                    x.Text = value;
            }
            DataTable dt = dgvData.DataSource as DataTable;
            if (dt != null)
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["Permission"] = "";
                }
        }

        private void btnGroup_Click(object sender, EventArgs e)
        {
            frmPOGView frm = new frmPOGView();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                txtCode.Text = frm.Code;
            }
        }




    }
}
