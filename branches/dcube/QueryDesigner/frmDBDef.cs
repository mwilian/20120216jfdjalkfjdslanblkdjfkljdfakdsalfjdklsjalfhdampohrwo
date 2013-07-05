using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using BUS;
using Janus.Windows.GridEX;

namespace dCube
{
    public partial class frmDBDef : Form
    {
        string _sErr = "";
        string _processStatus = "";
        string _code = "";
        public QDConfig _config = null;
        string _dtb = "";

        public string DTB
        {
            get { return _dtb; }
            set { _dtb = value; }
        }
        public frmDBDef(QDConfig config)
        {
            InitializeComponent();
            _config = config;
        }

        private void frmQDADD_Load(object sender, EventArgs e)
        {
            //InitConnection();
            EnableForm(false);
            //txtdatabase.Text = _dtb;
        }
        private void RefreshForm(string str)
        {
            txtCode.Text = str;
            _code = str;
            txtDescription.Text = str;

        }
        private void EnableForm(bool val)
        {
            txtCode.Enabled = val;
            panelControl.Enabled = val;
        }
        private void SetDataToForm(DTO.DBAInfo inf)
        {
            RefreshForm("");
            txtDescription.Text = inf.DESCRIPTION;
            txtCode.Text = inf.DB;
            txtTMP.Text = inf.REPORT_TEMPLATE_DRIVER;
        }


        private DTO.DBAInfo GetDataFromForm(DTO.DBAInfo inf)
        {
            inf.DB = inf.DB1 = inf.DB2 = txtCode.Text;
            inf.DESCRIPTION = txtDescription.Text;
            inf.REPORT_TEMPLATE_DRIVER = txtTMP.Text;

            return inf;
        }



        private string GetDocumentDirec()
        {
            return System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\" + Form_QD.DocumentFolder;
        }




        private void btnNew_Click(object sender, EventArgs e)
        {
            _processStatus = "C";
            RefreshForm("");
            EnableForm(true);
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            string sErr = "";
            _processStatus = "V";
            Form_DTBView frm = new Form_DTBView();
            //frm.Connect = _dtb;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (frm.Code_DTB != "")
                {
                    BUS.DBAControl ctr = new BUS.DBAControl();
                    DTO.DBAInfo inf = ctr.Get(frm.Code_DTB, ref sErr);
                    SetDataToForm(inf);
                }
            }
            if (sErr == "")
            {
                EnableForm(false);
                _processStatus = "V";
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            BUS.DBAControl ctr = new BUS.DBAControl();
            if (ctr.IsExist(txtCode.Text))
            {
                EnableForm(true);
                //ddlQD.Enabled = false;
                txtCode.Enabled = false;
                _processStatus = "A";
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string sErr = "";
            BUS.DBAControl ctr = new BUS.DBAControl();
            DTO.DBAInfo inf = new DTO.DBAInfo();

            if (_processStatus == "C")
            {
                if (!ctr.IsExist( txtCode.Text))
                    ctr.Add(GetDataFromForm(inf), ref sErr);
                else
                    sErr = txtCode.Text.Trim() + " is exist!";
            }
            else if (_processStatus == "A")
            {
                sErr = ctr.Update(GetDataFromForm(inf));
                _config.DIR[0].TMP = inf.REPORT_TEMPLATE_DRIVER;
            }
            if (sErr == "")
            {
                if (inf.REPORT_TEMPLATE_DRIVER != "")
                {
                    if (!File.Exists(inf.REPORT_TEMPLATE_DRIVER + "\\-.template.xls"))
                    {
                        File.Copy(Application.StartupPath + "\\-.template.xls", inf.REPORT_TEMPLATE_DRIVER + "\\-.template.xls");
                    }
                    if (!File.Exists(inf.REPORT_TEMPLATE_DRIVER + "\\-.template.xlsx"))
                    {
                        File.Copy(Application.StartupPath + "\\-.template.xlsx", inf.REPORT_TEMPLATE_DRIVER + "\\-.template.xlsx");
                    }
                    if (!File.Exists(inf.REPORT_TEMPLATE_DRIVER + "\\NODATA.xls"))
                    {
                        File.Copy(Application.StartupPath + "\\NODATA.xls", inf.REPORT_TEMPLATE_DRIVER + "\\NODATA.xls");
                    }
                }
                _processStatus = "V";
                EnableForm(false);
            }
            else MessageBox.Show(sErr);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            BUS.DBAControl ctr = new BUS.DBAControl();
            if (ctr.IsExist( txtCode.Text))
            {
                if (MessageBox.Show("Do you want to delete " + txtCode.Text + " schema?", "Message", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    string sErr = ctr.Delete( txtCode.Text);
                    RefreshForm("");
                    EnableForm(false);
                }
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            BUS.DBAControl ctr = new BUS.DBAControl();
            if (ctr.IsExist( txtCode.Text))
            {
                EnableForm(true);
                txtCode.Focus();
                txtCode.SelectAll();
                //txtCode.Text = "";
                //_code = "";
                _processStatus = "C";
            }
        }



        private void txtCode_TextChanged(object sender, EventArgs e)
        {


        }

        private void txtCode_Leave(object sender, EventArgs e)
        {

        }


        private void lbErr_Click(object sender, EventArgs e)
        {
            MessageBox.Show(lbErr.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                txtTMP.Text = folderBrowserDialog1.SelectedPath + "\\";
        }








    }
}
