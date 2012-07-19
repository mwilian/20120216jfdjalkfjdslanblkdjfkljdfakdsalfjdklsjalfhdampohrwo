using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace dCube
{
    public partial class frmTestMail : Form
    {
        DTO.LIST_TASKInfo _info = null;
        string _userID = "";

        public string UserID
        {
            get { return _userID; }
            set { _userID = value; }
        }
        public frmTestMail(DTO.LIST_TASKInfo info)
        {
            InitializeComponent();
            _info = info;
        }

        private void frmTestMail_Load(object sender, EventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string kq = "tavico://TASK?id=" + _info.Code;
            if (P1.Text != "")
                kq += "&P1=" + P1.Text;
            if (P2.Text != "")
                kq += "&P2=" + P2.Text;
            if (P3.Text != "")
                kq += "&P3=" + P3.Text;
            if (P4.Text != "")
                kq += "&P4=" + P4.Text;
            if (P5.Text != "")
                kq += "&P5=" + P5.Text;
            if (P6.Text != "")
                kq += "&P6=" + P6.Text;
            if (P7.Text != "")
                kq += "&P7=" + P7.Text;
            if (P8.Text != "")
                kq += "&P8=" + P8.Text;
            if (P9.Text != "")
                kq += "&P9=" + P9.Text;
            if (P10.Text != "")
                kq += "&P10=" + P10.Text;

            toolStripStatusLabel1.Text = kq;
            string[] cmd = kq.Split('?');
            CmdManager.UserID = _userID;
            string value = CmdManager.RunCmd(cmd[0], cmd[1]);
            if (value != "")
                toolStripStatusLabel1.Text = value;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(toolStripStatusLabel1.Text);
        }


    }
}
