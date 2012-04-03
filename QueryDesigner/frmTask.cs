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

namespace QueryDesigner
{
    public partial class frmTask : Form
    {
        DataSet _data = new DataSet("Schema");
        string _processStatus = "";
        string _code = "";
        string _dtb = "";
        string _user = "";
        public QDConfig _config = null;
        public frmTask(string dtb, string user)
        {
            InitializeComponent();
            _dtb = dtb;
            _user = user;
        }

        private void frmQDADD_Load(object sender, EventArgs e)
        {
            EnableForm(false);
        }
        private void RefreshForm(string str)
        {
            txtCode.Text = str;
            txtDescription.Text = str;
            ckbUse.Checked = false;
            AttQD_ID.Text = str;
            AttTemplate.Text = str;
            ValidRange.Text = str;
            CntQD_ID.Text = str;
            CntTemplate.Text = str;
            LoadListEmail();
            ClearEmailUser();
            Server.Text = "gmail.com";
            Port.Text = "587";
            Protocol.Text = "smtp";
            Email.Text = str;
            Password.Text = str;
            Lookup.Text = str;
            subRange.Text = str;
            Type.Text = str;
        }

        private void ClearEmailUser()
        {
            DataTable dt = dgvEmail.DataSource as DataTable;
            if (dt != null)
                dt.Rows.Clear();
            dgvEmail.DataSource = dt;
        }

        private void LoadListEmail()
        {

            ClearEmailUser();
            BUS.CommonControl ctr = new BUS.CommonControl();
            dgvList.DataSource = ctr.executeSelectQuery("Select * from LIST_EMAIL");

        }
        private void EnableForm(bool val)
        {
            txtCode.Enabled = val;
            txtDescription.Enabled = val;
            ckbUse.Enabled = val;
            AttQD_ID.Enabled = val;
            AttTemplate.Enabled = val;
            ValidRange.Enabled = val;
            CntQD_ID.Enabled = val;
            CntTemplate.Enabled = val;
            btnAdd.Enabled = btnFill.Enabled = btnRemove.Enabled = btnRemoveAll.Enabled = val;
            Server.Enabled = val;
            Port.Enabled = val;
            Protocol.Enabled = val;
            Email.Enabled = val;
            Password.Enabled = val;
            Lookup.Enabled = val;
            subRange.Enabled = val;
            Type.Enabled = val;
        }
        private void SetDataToForm(DTO.LIST_TASKInfo inf)
        {
            txtCode.Text = inf.Code;
            Lookup.Text = inf.Lookup;
            txtDescription.Text = inf.Description;
            AttQD_ID.Text = inf.AttQD_ID;
            AttTemplate.Text = inf.AttTmp;
            CntQD_ID.Text = inf.CntQD_ID;
            Type.Text = inf.Type;
            CntTemplate.Text = inf.CntTmp;
            Password.Text = inf.Password;
            LoadListEmail();
            SetListEmail(inf.Emails);
            Port.Text = inf.Port;
            Protocol.Text = inf.Protocol;
            Server.Text = inf.Server;
            Type.Text = inf.Type;
            Email.Text = inf.UserID;
            string[] arr = inf.ValidRange.Split(';');
            if (arr.Length >= 1)
                ValidRange.Text = arr[0];
            if (arr.Length >= 2)
                subRange.Text = arr[1];
            ckbUse.Checked = inf.IsUse == "Y";
        }

        private void SetListEmail(string emails)
        {
            if (emails != "")
            {
                DataTable dt = dgvList.DataSource as DataTable;
                DataTable dtU = dgvEmail.DataSource as DataTable;
                string[] arrMail = emails.Split(',');
                for (int i = 0; i < arrMail.Length; i++)
                {
                    Match mail = Regex.Match(arrMail[i], "<.+>");
                    for (int j = 0; j < dgvList.RowCount; j++)
                    {
                        if (dgvList.GetRow(j).Cells["Email"].Value.ToString() == mail.Value.Substring(1, mail.Value.Length - 2))
                        {
                            MoveRow(ref dt, ref dtU, ((DataRowView)dgvList.GetRow(j).DataRow).Row);
                        }
                    }
                }
                dgvEmail.DataSource = dtU;
                dgvList.DataSource = dt;
            }
        }
        private DTO.LIST_TASKInfo GetDataFromForm(DTO.LIST_TASKInfo inf)
        {
            inf.Code = txtCode.Text;
            inf.DTB = _dtb;
            inf.Lookup = Lookup.Text;
            inf.Description = txtDescription.Text;
            inf.AttQD_ID = AttQD_ID.Text;
            inf.AttTmp = AttTemplate.Text;
            inf.CntQD_ID = CntQD_ID.Text;
            inf.CntTmp = CntTemplate.Text;
            inf.Password = Password.Text;
            inf.Emails = GetListEmail();
            inf.Port = Port.Text;
            inf.Protocol = Protocol.Text;
            inf.Server = Server.Text;
            inf.UserID = Email.Text;
            inf.ValidRange = ValidRange.Text + ";" + subRange.Text;
            inf.IsUse = ckbUse.Checked ? "Y" : "N";
            inf.Type = Type.Text;
            return inf;
        }

        private string GetListEmail()
        {
            DataTable dtU = dgvEmail.DataSource as DataTable;
            string kq = "";
            if (dtU != null)
            {
                foreach (DataRow row in dtU.Rows)
                {
                    kq += ",\"" + row["Name"].ToString() + "\" <" + row["Mail"].ToString() + ">";
                }
            }
            if (kq.Length > 0)
                return kq.Substring(1);
            return "";
        }

        private string GetDocumentDirec()
        {
            return System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\TVC-QD";
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
            frmTaskView frm = new frmTaskView(_dtb, _user);
            //frm.Connect = Form_QD._dtb;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (frm.returnValue != null)
                {
                    BUS.LIST_TASKControl ctr = new BUS.LIST_TASKControl();
                    DTO.LIST_TASKInfo inf = ctr.Get(_dtb, ((object[])frm.returnValue)[0].ToString(), ref sErr);
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
            BUS.LIST_TASKControl ctr = new BUS.LIST_TASKControl();
            if (ctr.IsExist(_dtb, txtCode.Text))
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
            BUS.LIST_TASKControl ctr = new BUS.LIST_TASKControl();
            DTO.LIST_TASKInfo inf = new DTO.LIST_TASKInfo();

            if (_processStatus == "C")
            {
                if (!ctr.IsExist(_dtb, txtCode.Text))
                    ctr.Add(GetDataFromForm(inf), ref sErr);
                else
                    sErr = txtCode.Text.Trim() + " is exist!";
            }
            else if (_processStatus == "A")
            {
                sErr = ctr.InsertUpdate(GetDataFromForm(inf));
            }
            if (sErr == "")
            {
                _processStatus = "V";
                EnableForm(false);
            }
            else
            {
                lbErr.Text = sErr;
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            BUS.LIST_TASKControl ctr = new BUS.LIST_TASKControl();
            if (ctr.IsExist(_dtb, txtCode.Text))
            {
                if (MessageBox.Show("Do you want to delete " + txtCode.Text + " schema?", "Message", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    string sErr = ctr.Delete(_dtb, txtCode.Text);
                    RefreshForm("");
                    EnableForm(false);
                    _processStatus = "";
                }
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            BUS.LIST_TASKControl ctr = new BUS.LIST_TASKControl();
            if (ctr.IsExist(_dtb, txtCode.Text))
            {
                EnableForm(true);
                txtCode.Focus();
                txtCode.SelectAll();
                //txtCode.Text = "";
                //_code = "";
                _processStatus = "C";
            }
        }

        private void btnTransferIn_Click(object sender, EventArgs e)
        {
            FrmTransferIn frm = new FrmTransferIn("TASK");
            frm.ShowDialog();
        }

        private void btnTransferOut_Click(object sender, EventArgs e)
        {
            FrmTransferOut frm = new FrmTransferOut(_dtb, "TASK");
            //frm.DTB = Form_QD._dtb;
            frm.QD_CODE = txtCode.Text;
            frm.ShowDialog();
            //BUS.LIST_TASKControl ctr = new BUS.LIST_TASKControl();
            ////if (ctr.IsExist(ddlQD.Text, txtCode.Text))
            ////{
            ////DTO.LIST_TASKInfo inf = new DTO.LIST_TASKInfo();
            ////inf = GetDataFromForm(inf);
            //SaveFileDialog sfd = new SaveFileDialog();
            //sfd.Filter = "XML file(*.xml)|*.xml";
            //string sErr = "";
            //if (sfd.ShowDialog() == DialogResult.OK)
            //{
            //    DataTable dt = ctr.GetAll(ddlQD.Text, ref sErr);
            //    //dt.Rows.Add(inf.ToDataRow(dt));
            //    dt.WriteXml(sfd.FileName);
            //}
            //lbErr.Text = sErr;
            //}
        }





        private void btnQD_Click(object sender, EventArgs e)
        {
            if (_dtb != "")
            {
                Form_View a = new Form_View(_dtb, _user);
                a.database = _dtb;
                a.BringToFront();
                if (a.ShowDialog() == DialogResult.OK)
                {
                    AttQD_ID.Text = a.qdinfo.QD_ID;
                    //LoadQD(a.qdinfo);
                }
            }
            else
                lbErr.Text = "insert dtb";
        }

        private void dgvFrom_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void dgvFrom_DragEnter(object sender, DragEventArgs e)
        {

        }

        private void dgvFrom_DragDrop(object sender, DragEventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (_dtb != "")
            {
                Form_View a = new Form_View(_dtb, _user);
                a.database = _dtb;
                a.BringToFront();
                if (a.ShowDialog() == DialogResult.OK)
                {
                    CntQD_ID.Text = a.qdinfo.QD_ID;
                    //LoadQD(a.qdinfo);
                }
            }
            else
                lbErr.Text = "insert dtb";
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            frmEmail frm = new frmEmail(_dtb);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                LoadListEmail();
                 BUS.LIST_TASKControl ctr = new BUS.LIST_TASKControl();
                 if (ctr.IsExist(_dtb, txtCode.Text))
                 {
                     DTO.LIST_TASKInfo inf = new DTO.LIST_TASKInfo();
                     inf = GetDataFromForm(inf);
                     SetListEmail(inf.Emails);
                 }
            }
            
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (dgvList.Row >= 0)
            {
                DataRowView rowview = dgvList.CurrentRow.DataRow as DataRowView;
                DataTable dt = dgvList.DataSource as DataTable;
                DataTable dtU = dgvEmail.DataSource as DataTable;
                MoveRow(ref dt, ref dtU, rowview.Row);

                dgvEmail.DataSource = dtU;
                dgvList.DataSource = dt;

            }
        }

        private void MoveRow(ref DataTable dt, ref  DataTable dtU, DataRow dataRow)
        {
            if (dtU == null)
            {
                dtU = dt.Clone();
            }
            dtU.Rows.Add(dataRow.ItemArray);
            dt.Rows.Remove(dataRow);

        }

        private void btnFill_Click(object sender, EventArgs e)
        {

            DataTable dt = dgvList.DataSource as DataTable;

            DataTable dtU = dgvEmail.DataSource as DataTable;
            MoveRow(ref dt, ref dtU);
            dgvList.DataSource = dt;
            dgvEmail.DataSource = dtU;
        }

        private void MoveRow(ref DataTable dt, ref DataTable dtU)
        {
            if (dtU == null)
            {
                dtU = dt.Clone();
            }
            int cout = dt.Rows.Count - 1;
            for (int i = cout; i >= 0; i--)
            {
                DataRow dataRow = dt.Rows[i];
                dtU.Rows.Add(dataRow.ItemArray);
                dt.Rows.Remove(dataRow);
            }

        }

        private void btnRemoveAll_Click(object sender, EventArgs e)
        {
            DataTable dt = dgvList.DataSource as DataTable;

            DataTable dtU = dgvEmail.DataSource as DataTable;
            MoveRow(ref dtU, ref dt);
            dgvList.DataSource = dt;
            dgvEmail.DataSource = dtU;
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (dgvEmail.Row >= 0)
            {
                DataRowView rowview = dgvEmail.CurrentRow.DataRow as DataRowView;
                DataTable dt = dgvList.DataSource as DataTable;
                DataTable dtU = dgvEmail.DataSource as DataTable;
                MoveRow(ref dtU, ref dt, rowview.Row);

                dgvEmail.DataSource = dtU;
                dgvList.DataSource = dt;

            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            BUS.LIST_TASKControl ctr = new BUS.LIST_TASKControl();
            if (ctr.IsExist(_dtb, txtCode.Text))
            {
                DTO.LIST_TASKInfo inf = new DTO.LIST_TASKInfo();
                inf = GetDataFromForm(inf);
                frmTestMail frm = new frmTestMail(GetDataFromForm(inf));

                frm.ShowDialog();
            }
        }





    }
}
