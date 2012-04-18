using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.IO;
using System.Text.RegularExpressions;
using System.Data.SqlClient;

namespace dCube
{
    public partial class FrmSystem : Form
    {
        public string THEME = "";
        string _strConnect = "";
        DataSet dset = new DataSet("ROOT");
        public QDConfig _config = new QDConfig();
        public string Connecttion
        {
            get { return _strConnect; }
            set { _strConnect = value; }
        }
        public FrmSystem()
        {
            InitializeComponent();

        }
        private static string GetConnect(ref string type, string connect)
        {
            FrmConnection frm = new FrmConnection();

            //MSDASC.DataLinks mydlg = new MSDASC.DataLinks();
            //object cn = new ADODB.Connection();
            frm.Type = type;
            if (type == "QD" || type == null || type == "")
                frm.SetConnect("Provider=SQLOLEDB.1;" + connect);//((ADODB.Connection)cn).ConnectionString =
            else frm.SetConnect(connect);//((ADODB.Connection)cn).ConnectionString


            if (frm.ShowDialog() == DialogResult.OK)
            {
             
   
                string kq = frm.Connection;// ((ADODB.Connection)cn).ConnectionString;
                
                kq = BUS.CommonControl.RemoveAttribute(kq, "Persist Security Info");
                kq = BUS.CommonControl.RemoveAttribute(kq, "Extended Properties");
                kq = BUS.CommonControl.RemoveAttribute(kq, "Provider");
                if (type != "QD")
                    kq = "Provider=SQLOLEDB.1;" + kq;
                return kq;
            }
            return "";
            //else
            //    return ((ADODB.Connection)cn).ConnectionString;

        }
        private void dgvList_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > -1 && e.RowIndex > -1)
            {
                if (dgvList.Columns[e.ColumnIndex].Name == "BUILD")
                {

                   
                    string type = "";
                    string connect = "";
                    if (dgvList.Rows[e.RowIndex].Cells["TYPE"].Value != null)
                        type = dgvList.Rows[e.RowIndex].Cells["TYPE"].Value.ToString();
                    if (dgvList.Rows[e.RowIndex].Cells["CONTENT"].Value != null)
                        connect = dgvList.Rows[e.RowIndex].Cells["CONTENT"].Value.ToString();
                    connect = GetConnect(ref type, connect);

                    if (connect != "")
                    {
                        dgvList.Rows[e.RowIndex].Cells["CONTENT"].Value = connect;

                        dgvList.Rows[e.RowIndex].Cells["CONTENTEX"].Value = BUS.CommonControl.HiddenAttribute(connect, "Password");
                        dgvList.Rows[e.RowIndex].Cells["TEST"].Value = "Unknown";
                        dgvList.Rows[e.RowIndex].Cells["BUILD"].Value = "Build";
                        dgvList.Rows[e.RowIndex].Cells["TYPE"].Value = type;
                    }
                }
                else if (dgvList.Columns[e.ColumnIndex].Name == "TEST")
                {
                    if (dgvList.Rows[e.RowIndex].Cells["CONTENT"].Value != null && dgvList.Rows[e.RowIndex].Cells["TYPE"].Value != null)
                    {
                        if (dgvList.Rows[e.RowIndex].Cells["TYPE"].Value.ToString() == "QD")
                        {
                            System.Data.SqlClient.SqlConnection conn = new System.Data.SqlClient.SqlConnection(dgvList.Rows[e.RowIndex].Cells["CONTENT"].Value.ToString());
                            try { conn.Open(); dgvList.Rows[e.RowIndex].Cells["TEST"].Value = "OK"; }
                            catch
                            {
                                dgvList.Rows[e.RowIndex].Cells["TEST"].Value = "Fail";
                            }
                            finally { conn.Close(); }
                        }
                        else
                        {
                            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(dgvList.Rows[e.RowIndex].Cells["CONTENT"].Value.ToString());
                            try { conn.Open(); dgvList.Rows[e.RowIndex].Cells["TEST"].Value = "OK"; }
                            catch
                            {
                                dgvList.Rows[e.RowIndex].Cells["TEST"].Value = "Fail";
                            }
                            finally { conn.Close(); }
                        }

                    }
                    else
                    {
                        dgvList.Rows[e.RowIndex].Cells["TEST"].Value = "Fail";
                    }
                }
            }
        }

        private void FrmSystem_Load(object sender, EventArgs e)
        {
            dgvList.AutoGenerateColumns = false;

            //string startupPath = Application.StartupPath;
            //DataTable dt = new DataTable("ITEM");
            //DataColumn[] col = new DataColumn[]{new DataColumn("KEY"),
            //                        new DataColumn("DEFAULT"),
            //                        new DataColumn("CONTENT"),
            //                        new DataColumn("CONTENTEX"),
            //                        new DataColumn("TEST"),
            //                        new DataColumn("BUTTON"),
            //                        new DataColumn("TYPE")};


            //DataTable dtQD = new DataTable("DTB");
            //DataColumn[] colQD = new DataColumn[]{new DataColumn("QD"),
            //                        new DataColumn("AP")};

            //DataTable dtDIR = new DataTable("DIR");
            //DataColumn[] colDIR = new DataColumn[]{new DataColumn("TMP"),
            //                        new DataColumn("RPT")};

            //DataTable dtSYS = new DataTable("SYS");
            //DataColumn[] colSYS = new DataColumn[]{new DataColumn("FONT"),
            //                        new DataColumn("FORCECOLOR"),
            //                        new DataColumn("BACKCOLOR")};

            //dt.Columns.AddRange(col);
            //dtQD.Columns.AddRange(colQD);
            //dtDIR.Columns.AddRange(colDIR);
            //dtSYS.Columns.AddRange(colSYS);

            //dset.Tables.Add(dt);
            //dset.Tables.Add(dtQD);
            //dset.Tables.Add(dtDIR);
            //dset.Tables.Add(dtSYS);
            //string filename = GetDocumentDirec() + "\\Configuration\\xmlConnect.xml";
            _config.LoadConfig(GetDocumentDirec() + "\\Configuration\\QDConfig.tvc");
            bindingConfig.DataSource = _config;
            //iTEMBindingSource.DataSource = _config.ITEM;
            //if (File.Exists(filename))
            //{
            //    StreamReader sr = new StreamReader(filename);
            //    string result = sr.ReadToEnd();
            //    sr.Close();
            //    string kq = RC2.DecryptString(result, Form_QD._key, Form_QD._iv, Form_QD._padMode, Form_QD._opMode);
            //    StringReader stringReader = new StringReader(kq);
            //    dset.ReadXml(stringReader);
            //    stringReader.Close();
            //    //dset.ReadXml(filename);
            //}
            //dgvList.AutoGenerateColumns = false;
            //dgvList.DataSource = dt;
        }



        private string GetDocumentDirec()
        {
            return System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\" + Form_QD.DocumentFolder;
        }

        private void radPageViewPage2_Click(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;

            //dset = GetDataSet(dset);
            //_config = GetDataSetConfig(_config);
            //string filename = GetDocumentDirec() + "\\Configuration\\xmlConnect.xml";
            _config.SaveConfig(GetDocumentDirec() + "\\Configuration\\QDConfig.tvc");
            //StringBuilder sb = new StringBuilder();
            //StringWriter sw = new StringWriter(sb);
            //dset.WriteXml(sw);
            //sw.Close();
            //string temp = RC2.EncryptString(sb.ToString(), Form_QD._key, Form_QD._iv, Form_QD._padMode, Form_QD._opMode);
            //StreamWriter streamw = new StreamWriter(filename);
            //streamw.Write(temp);
            //streamw.Close();

            Close();
        }
        private QDConfig GetDataSetConfig(QDConfig dset)
        {
            if (dset.Tables["DTB"].Rows.Count > 0)
            {
                dset.Tables["DTB"].Rows[0]["QD"] = ddlQD.Text;
                dset.Tables["DTB"].Rows[0]["AP"] = ddlAP.Text;
            }
            else
            {
                DataRow row = dset.Tables["DTB"].NewRow();
                row["QD"] = ddlQD.Text;
                row["AP"] = ddlAP.Text;
                dset.Tables["DTB"].Rows.Add(row);
            }

            if (dset.Tables["DIR"].Rows.Count > 0)
            {
                dset.Tables["DIR"].Rows[0]["TMP"] = txtTMP.Text;
                dset.Tables["DIR"].Rows[0]["RPT"] = txtRPT.Text;
            }
            else
            {
                DataRow row = dset.Tables["DIR"].NewRow();
                row["TMP"] = txtTMP.Text;
                row["RPT"] = txtRPT.Text;
                dset.Tables["DIR"].Rows.Add(row);
            }
            if (dset.Tables["SYS"].Rows.Count > 0)
            {
                dset.Tables["SYS"].Rows[0]["FONT"] = txtFont.Text;
                if (panelForce.BackColor != SystemColors.Control)
                {
                    TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                    string color = tc.ConvertToString(panelForce.BackColor);
                    dset.Tables["SYS"].Rows[0]["FORCECOLOR"] = color;
                }
                else dset.Tables["SYS"].Rows[0]["FORCECOLOR"] = "";
                if (panelBack.BackColor != SystemColors.Control)
                {
                    TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                    string color = tc.ConvertToString(panelBack.BackColor);
                    dset.Tables["SYS"].Rows[0]["BACKCOLOR"] = color;
                }
                else dset.Tables["SYS"].Rows[0]["BACKCOLOR"] = "";

            }
            else
            {
                DataRow row = dset.Tables["SYS"].NewRow();
                row["FONT"] = txtFont.Text;
                if (panelForce.BackColor != SystemColors.Control)
                {
                    TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                    string color = tc.ConvertToString(panelForce.BackColor);
                    row["FORCECOLOR"] = color;
                }
                else row["FORCECOLOR"] = "";
                if (panelBack.BackColor != SystemColors.Control)
                {
                    TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                    string color = tc.ConvertToString(panelBack.BackColor);
                    row["BACKCOLOR"] = color;
                }
                else row["BACKCOLOR"] = "";
                dset.Tables["SYS"].Rows.Add(row);
            }
            return dset;
        }
        private DataSet GetDataSet(DataSet dset)
        {
            if (dset.Tables["DTB"].Rows.Count > 0)
            {
                dset.Tables["DTB"].Rows[0]["QD"] = ddlQD.Text;
                dset.Tables["DTB"].Rows[0]["AP"] = ddlAP.Text;
            }
            else
            {
                DataRow row = dset.Tables["DTB"].NewRow();
                row["QD"] = ddlQD.Text;
                row["AP"] = ddlAP.Text;
                dset.Tables["DTB"].Rows.Add(row);
            }

            if (dset.Tables["DIR"].Rows.Count > 0)
            {
                dset.Tables["DIR"].Rows[0]["TMP"] = txtTMP.Text;
                dset.Tables["DIR"].Rows[0]["RPT"] = txtRPT.Text;
            }
            else
            {
                DataRow row = dset.Tables["DIR"].NewRow();
                row["TMP"] = txtTMP.Text;
                row["RPT"] = txtRPT.Text;
                dset.Tables["DIR"].Rows.Add(row);
            }
            if (dset.Tables["SYS"].Rows.Count > 0)
            {
                dset.Tables["SYS"].Rows[0]["FONT"] = txtFont.Text;
                if (panelForce.BackColor != SystemColors.Control)
                {
                    TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                    string color = tc.ConvertToString(panelForce.BackColor);
                    dset.Tables["SYS"].Rows[0]["FORCECOLOR"] = color;
                }
                else dset.Tables["SYS"].Rows[0]["FORCECOLOR"] = "";
                if (panelBack.BackColor != SystemColors.Control)
                {
                    TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                    string color = tc.ConvertToString(panelBack.BackColor);
                    dset.Tables["SYS"].Rows[0]["BACKCOLOR"] = color;
                }
                else dset.Tables["SYS"].Rows[0]["BACKCOLOR"] = "";

            }
            else
            {
                DataRow row = dset.Tables["SYS"].NewRow();
                row["FONT"] = txtFont.Text;
                if (panelForce.BackColor != SystemColors.Control)
                {
                    TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                    string color = tc.ConvertToString(panelForce.BackColor);
                    row["FORCECOLOR"] = color;
                }
                else row["FORCECOLOR"] = "";
                if (panelBack.BackColor != SystemColors.Control)
                {
                    TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                    string color = tc.ConvertToString(panelBack.BackColor);
                    row["BACKCOLOR"] = color;
                }
                else row["BACKCOLOR"] = "";
                dset.Tables["SYS"].Rows.Add(row);
            }
            return dset;
        }
        bool flag = true;
        private void dgvList_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (flag && dgvList.Columns[e.ColumnIndex].Name == "DEFAULT" && dgvList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null && dgvList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "True")
            {
                flag = false;
                foreach (DataGridViewRow row in dgvList.Rows)
                {
                    if (row != dgvList.Rows[e.RowIndex])
                    {
                        row.Cells[e.ColumnIndex].Value = "False";
                    }
                }
                flag = true;
            }

        }

        private void dgvList_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            //for (int i = e.RowIndex; i < e.RowCount; i++)
            //{
            //    dgvList.Rows[i].Cells["BUILD"].Value = "Build";
            //    dgvList.Rows[i].Cells["TEST"].Value = "Unknown";
            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                txtTMP.Text = folderBrowserDialog1.SelectedPath + "\\";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                txtRPT.Text = folderBrowserDialog1.SelectedPath + "\\";
        }
        bool flagQD = true;
        bool flagAP = true;
        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (tabControl.SelectedTab == tabAP)
            //{
            //    string startupPath = Application.StartupPath;
            //DataTable dt = _config.Tables["ITEM"].Clone();
            //DataRow[] arrRow = _config.Tables["ITEM"].Copy().Select("TYPE='QD'");
            //foreach (DataRow row in arrRow)
            //{
            //    dt.ImportRow(row);
            //}
            //ddlQD.DataSource = dt;
            //ddlQD.DisplayMember = "KEY";
            //ddlQD.ValueMember = "KEY";

            //dt = _config.Tables["ITEM"].Clone();
            //arrRow = _config.Tables["ITEM"].Copy().Select("TYPE='AP'");
            //foreach (DataRow row in arrRow)
            //{
            //    dt.ImportRow(row);
            //}
            //ddlAP.DataSource = dt;
            //ddlAP.DisplayMember = "KEY";
            //ddlAP.ValueMember = "KEY";

            //    if (dset.Tables["DTB"].Rows.Count > 0)
            //    {
            //        ddlQD.Text = dset.Tables["DTB"].Rows[0]["QD"].ToString();
            //        ddlAP.Text = dset.Tables["DTB"].Rows[0]["AP"].ToString();
            //    }

            //    if (dset.Tables["DIR"].Rows.Count > 0)
            //    {
            //        if (txtTMP.Text == "")
            //            txtTMP.Text = dset.Tables["DIR"].Rows[0]["TMP"].ToString();
            //        if (txtRPT.Text == "")
            //            txtRPT.Text = dset.Tables["DIR"].Rows[0]["RPT"].ToString();
            //    }
            //    if (dset.Tables["SYS"].Rows.Count > 0)
            //    {
            //        if (txtFont.Text == "")
            //            txtFont.Text = dset.Tables["SYS"].Rows[0]["FONT"].ToString();
            //        if (dset.Tables["SYS"].Rows[0]["FORCECOLOR"].ToString() != "")
            //        {
            //            TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
            //            Color color = (Color)tc.ConvertFrom(dset.Tables["SYS"].Rows[0]["FORCECOLOR"].ToString());
            //            panelForce.BackColor = color;
            //        }
            //        if (dset.Tables["SYS"].Rows[0]["BACKCOLOR"].ToString() != "")
            //        {
            //            TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
            //            Color color = (Color)tc.ConvertFrom(dset.Tables["SYS"].Rows[0]["BACKCOLOR"].ToString());
            //            panelBack.BackColor = color;
            //        }
            //    }
            //}
        }

        private void dgvList_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow row in dgvList.Rows)
            {
                if (row.Cells["CONTENT"].Value != null)
                    row.Cells["ContentEx"].Value = Regex.Replace(row.Cells["CONTENT"].Value.ToString(), "[pP]+assword=.*$", "password=******");
            }
        }

        private void ddlQD_SelectedIndexChanged(object sender, EventArgs e)
        {
            flagQD = false;
        }

        private void ddlAP_SelectedIndexChanged(object sender, EventArgs e)
        {
            flagAP = false;
        }

        private void btnFont_Click(object sender, EventArgs e)
        {
            if (fontDialog1.ShowDialog() == DialogResult.OK)
            {
                TypeConverter tc = TypeDescriptor.GetConverter(typeof(Font));
                string fontString = tc.ConvertToString(fontDialog1.Font);
                txtFont.Font = fontDialog1.Font;
                txtFont.Text = fontString;
            }
        }


        private void panelForce_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                panelForce.BackColor = colorDialog1.Color;

                TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                string color = tc.ConvertToString(colorDialog1.Color);
                txtForce.Text = color;

            }
        }

        private void panelBack_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                panelBack.BackColor = colorDialog1.Color;
                TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                string color = tc.ConvertToString(colorDialog1.Color);
                txtBack.Text = color;
            }
        }

        private void btnForce_Click(object sender, EventArgs e)
        {
            panelForce.BackColor = SystemColors.Control;
            txtForce.Text = "";

        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            panelBack.BackColor = SystemColors.Control;
            txtBack.Text = "";
        }

        private void txtForce_TextChanged(object sender, EventArgs e)
        {

        }


    }
}