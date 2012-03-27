using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BUS;

namespace QueryDesigner
{
    public partial class frmChangeDB : Form
    {
        string sErr = "";
        string _user;

        public string User
        {
            get { return _user; }
            set { _user = value; }
        }
        string _dtb = "";
        public frmChangeDB(string db)
        {
            InitializeComponent();
            _dtb = db;
        }

        private void txtdatabase_Validated(object sender, EventArgs e)
        {
            BUS.DBAControl dbaCtr = new DBAControl();
            DTO.DBAInfo dbaInf = dbaCtr.Get(txtdatabase.Text, ref sErr);
            txt_database.Text = dbaInf.DESCRIPTION;
            //ResetForm();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Form_QD.DB = txtdatabase.Text;
            if (_user != "TVC" && checkBox1.Checked)
            {
                BUS.PODControl podCtr = new PODControl();
                DTO.PODInfo podInf = podCtr.Get(_user, ref sErr);
                podInf.DB_DEFAULT = txtdatabase.Text;
                podCtr.Update(podInf);
            }
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void txtdatabase_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                bt_database_Click(null, null);
        }

        private void bt_database_Click(object sender, EventArgs e)
        {

            Form_DTBView a = new Form_DTBView();
            //a.themname = THEME;
            a.BringToFront();
            if (a.ShowDialog(this) == DialogResult.OK)
            {
                txtdatabase.Text = a.Code_DTB;
                txt_database.Text = a.Description_DTB;
            }
        }

        private void frmChangeDB_Load(object sender, EventArgs e)
        {
            txtdatabase.Text = _dtb;
            BUS.DBAControl dbaCtr = new DBAControl();
            DTO.DBAInfo dbaInf = dbaCtr.Get(txtdatabase.Text, ref sErr);
            txt_database.Text = dbaInf.DESCRIPTION;
        }
    }
}
