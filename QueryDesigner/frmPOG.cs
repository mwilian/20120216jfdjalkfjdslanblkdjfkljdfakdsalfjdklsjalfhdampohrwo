using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace QueryDesigner
{
    public partial class frmPOG : Form
    {
        BUS.POGControl ctr = new BUS.POGControl();
        string _sErr = "";
        string _processStatus = "";
        public frmPOG()
        {
            InitializeComponent();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            RefreshForm("");
            EnabledForm(true);
            txtCode.Enabled = true;
            _processStatus = "C";
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            frmPOGView frm = new frmPOGView();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                SetDataToForm(ctr.Get(frm.Code, ref _sErr));
                EnabledForm(false);
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (ctr.IsExist(txtCode.Text))
            {
                EnabledForm(true);
                txtCode.Enabled = false;
                _processStatus = "A";
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //string sErr = "";

            DTO.POGInfo inf = new DTO.POGInfo();

            if (_processStatus == "C")
            {
                if (!ctr.IsExist(txtCode.Text))
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
            if (ctr.IsExist(txtCode.Text))
            {
                _sErr = ctr.Delete(txtCode.Text);
                RefreshForm("");
                EnabledForm(false);
                _processStatus = "";
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (ctr.IsExist(txtCode.Text))
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
        }
        private void lbErr_Click(object sender, EventArgs e)
        {
            if (lbErr.Text != "" && lbErr.Text != "...")
                MessageBox.Show(lbErr.Text);
        }
        private void SetDataToForm(DTO.POGInfo obj)
        {
            txtCode.Text = obj.ROLE_ID;
            txtName.Text = obj.ROLE_NAME;
            txtLen.Text = obj.PASS_MIN_LEN;
            txtQD.Text = obj.RPT_CODE;
            txtValid.Text = obj.PASS_VALID;
        }
        private DTO.POGInfo GetDataFromForm(DTO.POGInfo obj)
        {
            //DTO.PODInfo obj = new DTO.PODInfo();
            obj.ROLE_ID1 = obj.ROLE_ID = txtCode.Text;
            obj.ROLE_NAME = txtName.Text;
            obj.PASS_MIN_LEN = txtLen.Text;
            obj.PASS_VALID = txtValid.Text;
            obj.RPT_CODE = txtQD.Text;
            obj.TB = "POG";
            return obj;
        }
        private void EnabledForm(bool value)
        {
            pContain.Enabled = value;
        }
        private void RefreshForm(string value)
        {
            foreach (Control x in pContain.Controls)
            {
                if (x is TextBox)
                    x.Text = value;
            }
        }


        private void btnRole_Click(object sender, EventArgs e)
        {

        }

        private void btnDB_Click(object sender, EventArgs e)
        {

        }




    }
}
