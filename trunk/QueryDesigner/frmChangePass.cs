using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace dCube
{
    public partial class frmChangePass : Form
    {
        string _user = "";
        string _sErr = "";
        public frmChangePass(string user)
        {
            InitializeComponent();
            _user = user;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            BUS.PODControl podCtr = new BUS.PODControl();
            DTO.PODInfo podInf = podCtr.Get(_user, ref _sErr);
            string infpass = podInf.PASS;
            string oldpass = Convert.ToBase64String(new System.Security.Cryptography.SHA1CryptoServiceProvider().ComputeHash(Encoding.ASCII.GetBytes(txtOld.Text)));
            string newpass = Convert.ToBase64String(new System.Security.Cryptography.SHA1CryptoServiceProvider().ComputeHash(Encoding.ASCII.GetBytes(txtNew.Text)));
            string confirmpass = Convert.ToBase64String(new System.Security.Cryptography.SHA1CryptoServiceProvider().ComputeHash(Encoding.ASCII.GetBytes(txtConfirm.Text)));
            if (infpass != oldpass)
                MessageBox.Show("Wrong password");
            else if (newpass != confirmpass)
                MessageBox.Show("Confirm password is not correct");
            else
            {
                podInf.PASS = newpass;
                _sErr = podCtr.Update(podInf);
                Close();
            }
            
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
