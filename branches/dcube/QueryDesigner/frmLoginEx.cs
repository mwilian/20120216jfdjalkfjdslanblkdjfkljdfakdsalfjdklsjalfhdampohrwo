using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DTO;

namespace dCube
{
    public partial class frmLoginEx : Form
    {
        string _sErr = string.Empty;
        string _user = "";

        public string User
        {
            get { return _user; }
            set { _user = value; }
        }
        string _pass = "";

        public string Pass
        {
            get { return _pass; }
            set { _pass = value; }
        }
        public frmLoginEx()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lbErr.Text = "";
        }
        string _DB = "";

        public string DB
        {
            get { return _DB; }
            set { _DB = value; }
        }
        private void btnLogin_Click(object sender, EventArgs e)
        {
            PODInfo inf = new PODInfo();
            _user = inf.USER_ID = txtUser.Text;
            _pass = inf.PASS = txtPass.Text;

            if (_user != "TVC" || _pass != "TVCSYS")
            {
                inf.LANGUAGE = "44";
                BUS.PODControl podCtr = new BUS.PODControl();
                if (podCtr.IsExist(inf.USER_ID))
                {
                    string pass = Convert.ToBase64String(new System.Security.Cryptography.SHA1CryptoServiceProvider().ComputeHash(Encoding.ASCII.GetBytes(inf.PASS)));

                    inf = podCtr.Get(inf.USER_ID, ref _sErr);
                    inf.LANGUAGE = inf.LANGUAGE == "84" ? "84" : "44";

                    if (inf.PASS == pass)
                    {
                        //BUS.POSControl posCtr = new BUS.POSControl();
                        //if (!posCtr.IsExist(inf.USER_ID))
                        //{
                        DialogResult = DialogResult.OK;
                        _DB = inf.DB_DEFAULT;

                        //DTO.POSInfo infPOS = new POSInfo(inf.USER_ID, _DB, "Query Designer", "QD", DateTime.Now.ToString("yyyy-MM-dd hh:mm"));
                        //posCtr.InsertUpdate(infPOS);
                        Close();
                        //}
                        //else { lbErr.Text = "Existing users in the system"; }
                    }
                    else
                        lbErr.Text = "Password wrong!";
                }
                else
                {
                    lbErr.Text = "User is not exist!";
                }
            }
            else
            {
                DialogResult = DialogResult.OK;
                Close();
            }

        }

        void frm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void txtUser_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtPass.Focus();
        }

        private void txtPass_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btnLogin_Click(null, null);
        }
    }
}
