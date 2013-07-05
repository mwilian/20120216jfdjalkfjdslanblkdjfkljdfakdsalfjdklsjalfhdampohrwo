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
    public partial class frmImportDefinition : Form
    {
        string _sErr = "";
        DataTable _dtField = new DataTable("_TableName");
        string _processStatus = "";
        string _code = "";
        public QDConfig _config = null;
        string _dtb = "";

        public string DTB
        {
            get { return _dtb; }
            set { _dtb = value; }
        }
        public frmImportDefinition()
        {
            InitializeComponent();
        }

        private void frmQDADD_Load(object sender, EventArgs e)
        {
            InitConnection();
            EnableForm(false);
            txtdatabase.Text = _dtb;
        }
        private void RefreshForm(string str)
        {
            txtCode.Text = str;
            _code = str;
            txtDescription.Text = str;
            txtLookup.Text = str;
            _dtField.Rows.Clear();

        }
        private void EnableForm(bool val)
        {
            txtCode.Enabled = val;
            panelControl.Enabled = val;
            panelTab.Enabled = val;
            Janus.Windows.GridEX.InheritableBoolean tmp = Janus.Windows.GridEX.InheritableBoolean.False;
            if (val)
                tmp = Janus.Windows.GridEX.InheritableBoolean.True;
            dgvField.AllowEdit = dgvField.AllowDelete = dgvField.AllowAddNew = tmp;
            ddlQD.Enabled = val;
        }
        private void SetDataToForm(DTO.IMPORT_SCHEMAInfo inf)
        {
            RefreshForm("");
            txtdatabase.Text = inf.CONN_ID;
            ddlQD.SelectedText = inf.DEFAULT_CONN;
            txtDescription.Text = inf.DESCRIPTN;
            txtLookup.Text = inf.LOOK_UP;
            txtCode.Text = inf.SCHEMA_ID;
            txtDAG.Text = inf.DAG;
            _code = txtCode.Text.Trim();
            _dtField = IMPORT_SCHEMAControl.GetDataTableFromXML(_dtField, inf.FIELD_TEXT);

        }


        private DTO.IMPORT_SCHEMAInfo GetDataFromForm(DTO.IMPORT_SCHEMAInfo inf)
        {
            inf.DB = txtdatabase.Text;
            inf.CONN_ID = txtdatabase.Text;
            inf.DEFAULT_CONN = ddlQD.Text;
            inf.DESCRIPTN = txtDescription.Text;
            inf.LOOK_UP = txtLookup.Text;
            inf.DAG = txtDAG.Text;
            inf.UPDATED = DateTime.Today.Year * 10000 + DateTime.Today.Month * 100 + DateTime.Today.Day;
            inf.FIELD_TEXT = IMPORT_SCHEMAControl.GetXMLFromDataTable(_dtField);
            //inf.FIELD_TEXT = GetFieldCode(_dtField);
            inf.SCHEMA_ID = txtCode.Text;

            return inf;
        }


        private void InitConnection()
        {

            DataTable dtTmp = _config.Tables["ITEM"].Clone();
            DataRow[] arrRow = _config.Tables["ITEM"].Copy().Select("TYPE='AP'");
            foreach (DataRow row in arrRow)
            {
                dtTmp.ImportRow(row);
            }
            ddlQD.DataSource = dtTmp;
            ddlQD.DisplayMember = "KEY";
            ddlQD.ValueMember = "CONTENT";          



            _dtField = new DataTable("_TableName");
            DataColumn[] colfield = new DataColumn[] {  new DataColumn("PrimaryKey")
                ,new DataColumn("Key")
                , new DataColumn("DataTypeCode")
                , new DataColumn("Caption")
                , new DataColumn("DataMember")
                , new DataColumn("Position")
                , new DataColumn("AggregateFunction")
                , new DataColumn("IsNull")
                , new DataColumn("Visible")
                , new DataColumn("Tag")};
            _dtField.Columns.AddRange(colfield);

            bsField.DataSource = _dtField;
        }
        private string GetDocumentDirec()
        {
            return System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\" + Form_QD.DocumentFolder;
        }

        private void ddlQD_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddlQD.SelectedIndex >= 0)
                txtConnect.Text = ddlQD.SelectedValue.ToString();
        }

        private void btnSelectTable_Click(object sender, EventArgs e)
        {
            string key = "Table List";
            LookupDB(key);

        }

        private void LookupDB(string key)
        {
            frmLookup frm = new frmLookup();
            frm.Text = key;
            frm.Connect = txtConnect.Text;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                string tableName = frm.ReturnCode;
                if (tableName != "")
                {
                    txtDescription.Text = txtLookup.Text = txtCode.Text = tableName;

                    DataTable kq = new DataTable();
                    System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(frm.Connect);
                    try
                    {
                        conn.Open();
                        kq = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Columns, new Object[] { null, null, tableName, null });
                        //dgvField.AutoGenerateColumns = true;
                        _dtField.Clear();
                        for (int i = 0; i < kq.Rows.Count; i++)
                        {
                            DataRow row = kq.Rows[i];
                            DataRow newRow = _dtField.NewRow();
                            newRow["Key"] = newRow["DataMember"] = newRow["Caption"] = row["COLUMN_NAME"];

                            if (row["DATA_TYPE"].ToString() == "135")
                                newRow["DataTypeCode"] = "DateTime";
                            else if (row["DATA_TYPE"].ToString() == "5")
                                newRow["DataTypeCode"] = "Double";
                            else
                                newRow["DataTypeCode"] = "String";
                            newRow["Visible"] = "True";
                            newRow["Position"] = (i + 1).ToString();
                            newRow["IsNull"] = row["IS_NULLABLE"];
                            newRow["PrimaryKey"] = "False";
                            //newRow["Tag"] = "";
                            _dtField.Rows.Add(newRow);
                        }
                        //dgvField.DataSource = kq;
                    }
                    catch { }
                    finally { conn.Close(); }
                    dgvField.AutoSizeColumns();
                }

            }
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
            frmImportDefView frm = new frmImportDefView();
            frm.Connect = _dtb;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (frm.ReturnCode != "")
                {
                    BUS.IMPORT_SCHEMAControl ctr = new BUS.IMPORT_SCHEMAControl();
                    DTO.IMPORT_SCHEMAInfo inf = ctr.Get(_dtb, frm.ReturnCode, ref sErr);
                    SetDataToForm(inf);
                }
            }
            if (sErr == "")
            {
                EnableForm(false);
                _processStatus = "V";
            }
            dgvField.AutoSizeColumns();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            BUS.IMPORT_SCHEMAControl ctr = new BUS.IMPORT_SCHEMAControl();
            if (ctr.IsExist(ddlQD.Text, txtCode.Text))
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
            BUS.IMPORT_SCHEMAControl ctr = new BUS.IMPORT_SCHEMAControl();
            DTO.IMPORT_SCHEMAInfo inf = new DTO.IMPORT_SCHEMAInfo();

            if (_processStatus == "C")
            {
                if (!ctr.IsExist(ddlQD.Text, txtCode.Text))
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
            else MessageBox.Show(sErr);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            BUS.IMPORT_SCHEMAControl ctr = new BUS.IMPORT_SCHEMAControl();
            if (ctr.IsExist(ddlQD.Text, txtCode.Text))
            {
                if (MessageBox.Show("Do you want to delete " + txtCode.Text + " schema?", "Message", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    string sErr = ctr.Delete(ddlQD.Text, txtCode.Text);
                    RefreshForm("");
                    EnableForm(false);
                }
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            BUS.IMPORT_SCHEMAControl ctr = new BUS.IMPORT_SCHEMAControl();
            if (ctr.IsExist(ddlQD.Text, txtCode.Text))
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

        private void dgvField_MouseDown(object sender, MouseEventArgs e)
        {
            int index = dgvField.RowPositionFromPoint(e.X, e.Y);
            if (index > -1 && dgvField.ColumnFromPoint(e.X, e.Y) == null && dgvField.AllowEdit == Janus.Windows.GridEX.InheritableBoolean.True)
            {
                if (dgvField.GetRow(index).RowType == Janus.Windows.GridEX.RowType.Record)
                    dgvField.DoDragDrop(dgvField.GetRow(index), DragDropEffects.Move);
            }


        }

        private void dgvField_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(Janus.Windows.GridEX.GridEXRow)))
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        private void dgvField_DragDrop(object sender, DragEventArgs e)
        {
            Janus.Windows.GridEX.GridEXRow row = (Janus.Windows.GridEX.GridEXRow)e.Data.GetData(typeof(Janus.Windows.GridEX.GridEXRow));
            if (row != null)
            {
                int indexSource = row.RowIndex;
                Point p = dgvField.PointToClient(new Point(e.X, e.Y));
                int rowIndex = dgvField.RowPositionFromPoint(p.X, p.Y);
                if (rowIndex >= 0 && rowIndex < dgvField.RowCount && indexSource != rowIndex)
                {
                    DataRow rowmove = _dtField.NewRow();
                    rowmove.ItemArray = _dtField.Rows[indexSource].ItemArray;
                    _dtField.Rows.RemoveAt(indexSource);
                    _dtField.Rows.InsertAt(rowmove, rowIndex);
                }
            }
        }







        private void txtConnect_TextChanged(object sender, EventArgs e)
        {

            txtConectEx.Text = BUS.CommonControl.HiddenAttribute(txtConnect.Text, "Password");
        }

        private void lbErr_Click(object sender, EventArgs e)
        {
            MessageBox.Show(lbErr.Text);
        }

        private void dgvField_LinkClicked(object sender, Janus.Windows.GridEX.ColumnActionEventArgs e)
        {
            //frmValidatedList frm = new frmValidatedList(txtdatabase.Text, Form_QD._user);
            //if (frm.ShowDialog() == DialogResult.OK)
            //{

            //}
        }

        

        private void txtdatabase_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void dgvField_MouseUp(object sender, MouseEventArgs e)
        {
            GridEXColumn col = dgvField.ColumnFromPoint(e.X, e.Y);
            int rowInx = dgvField.RowPositionFromPoint(e.X, e.Y);
            if (rowInx >= 0 && col != null && col.Key == "Tag")
            {

                objValidatedList obj = null;
                frmValidatedList frm;
                if (dgvField.GetRow(rowInx).Cells[col].Value != DBNull.Value)
                {
                    obj = new objValidatedList(dgvField.GetRow(rowInx).Cells[col].Value);
                    frm = new frmValidatedList(txtdatabase.Text, Form_QD._user, obj);
                }
                else frm = new frmValidatedList(txtdatabase.Text, Form_QD._user);
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    dgvField.GetRow(rowInx).Cells[col].Value = frm.ObjReturn;
                }
            }
        }

        private void bt_group_Click(object sender, EventArgs e)
        {
            frmDAView frm = new frmDAView();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                txtDAG.Text = frm.Code;
            }
        }

        private void txtDAG_TextChanged(object sender, EventArgs e)
        {
            LIST_DAControl ctr = new LIST_DAControl();
            lbgroup.Text = ctr.Get(txtDAG.Text, ref _sErr).NAME;
        }






    }
}
