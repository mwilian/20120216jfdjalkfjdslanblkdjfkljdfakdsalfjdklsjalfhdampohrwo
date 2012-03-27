using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace QueryDesigner
{
    public partial class frmPOD : Form
    {
        BUS.PODControl ctr = new BUS.PODControl();
        string _sErr = "";
        string _processStatus = "";
        public frmPOD()
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
            frmPODView frm = new frmPODView();
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

            DTO.PODInfo inf = new DTO.PODInfo();

            if (_processStatus == "C")
            {
                if (!ctr.IsExist(txtCode.Text))
                {
                    inf.PASS = Convert.ToBase64String(new System.Security.Cryptography.SHA1CryptoServiceProvider().ComputeHash(Encoding.ASCII.GetBytes("")));
                    ctr.Add(GetDataFromForm(inf), ref _sErr);
                }
                else
                    _sErr = txtCode.Text.Trim() + " is exist!";
            }
            else if (_processStatus == "A")
            {
                inf = ctr.Get(txtCode.Text, ref _sErr);
                _sErr = ctr.Update(GetDataFromForm(inf));
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
        private void SetDataToForm(DTO.PODInfo obj)
        {
            txtCode.Text = obj.USER_ID;
            txtName.Text = obj.USER_NAME;
            txtGroup.Text = obj.ROLE_ID;
            txtDB.Text = obj.DB_DEFAULT;
            txtLanguage.Text = obj.LANGUAGE;
        }
        private DTO.PODInfo GetDataFromForm(DTO.PODInfo obj)
        {
            //DTO.PODInfo obj = new DTO.PODInfo();
            obj.USER_ID = obj.USER_ID1 = txtCode.Text;
            obj.USER_NAME = txtName.Text;
            obj.ROLE_ID = txtGroup.Text;
            obj.LANGUAGE = txtLanguage.Text;
            obj.DB_DEFAULT = txtDB.Text;
            obj.TB = "POD";
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
            frmPOGView frm = new frmPOGView();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                txtGroup.Text = frm.Code;
            }
        }

        private void txtDB_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtGroup_TextChanged(object sender, EventArgs e)
        {
            BUS.POGControl pogCtr = new BUS.POGControl();
            lbRole.Text = pogCtr.Get(txtGroup.Text, ref _sErr).ROLE_NAME;
        }




    }
}
