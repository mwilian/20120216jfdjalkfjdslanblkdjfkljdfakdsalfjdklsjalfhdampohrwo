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

namespace dCube
{
    public partial class frmQDADD : Form
    {
        string _db = "";
        DataSet _data = new DataSet("Schema");
        string _processStatus = "";
        string _code = "";
        public QDConfig _config = null;
        string _user = "";
        public frmQDADD(string db, string user)
        {
            InitializeComponent();
            _db = db;
            _user = user;
        }

        private void frmQDADD_Load(object sender, EventArgs e)
        {
            InitConnection();
            EnableForm(false);
        }
        private void RefreshForm(string str)
        {
            txtCode.Text = str;
            _code = str;
            txtDescription.Text = str;
            txtLookup.Text = str;
            _data.Tables["field"].Rows.Clear();
            _data.Tables["fromcode"].Rows.Clear();
            Group.Text = str;

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
            if (val) dgvField.ContextMenuStrip = contextMenuStrip1;
            else dgvField.ContextMenuStrip = null;
            dgvFrom.AllowEdit = dgvFrom.AllowDelete = dgvFrom.AllowAddNew = tmp;
            ddlQD.Enabled = val;
            Group.Enabled = val;
        }
        private void SetDataToForm(DTO.LIST_QD_SCHEMAInfo inf)
        {
            RefreshForm("");
            Group.Text = inf.DAG;
            ddlQD.Text = inf.DEFAULT_CONN;
            txtDescription.Text = inf.DESCRIPTN;
            txtLookup.Text = inf.LOOK_UP;
            txtCode.Text = inf.SCHEMA_ID;
            _code = txtCode.Text.Trim();
            ckbUse.Checked = inf.SCHEMA_STATUS.Trim() == "Y" ? true : false;
           
                _data = ReadScheme(inf);           



            //strR = new StringReader(inf.FROM_TEXT);
            //_data.Tables["fromcode"].ReadXml(strR);
            //strR.Close();

        }

        private DataSet ReadScheme(DTO.LIST_QD_SCHEMAInfo inf)
        {
            if (inf.FIELD_TEXT == "")
            {
                StringReader strR = new StringReader(inf.FROM_TEXT);
                try
                {
                    _data.ReadXml(strR);
                }
                catch (Exception ex)
                {
                    if (MessageBox.Show("Schema structure is erro!\n Do you want to delete it?", "Message", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        btnDelete_Click(null, null);
                    }
                }
                strR.Close();
                return _data;
            }
            else
            {
                XmlDocument xml = new XmlDocument();
                string strxml = string.Format("<?xml version=\"1.0\" encoding=\"utf-8\" ?><SUN_SCHEMA>{0}</SUN_SCHEMA>", inf.FROM_TEXT);
                xml.LoadXml(strxml);
                XmlElement doc = xml.DocumentElement; 
                DataTable dtfrom = _data.Tables["fromcode"];
                foreach (XmlElement ele in doc.ChildNodes)
                {
                    DataRow newRow = dtfrom.NewRow();
                    newRow["fromcode"] = ele.GetAttribute("fromcode");
                    newRow["lookup"] = ele.GetAttribute("lookup");

                    dtfrom.Rows.Add(newRow);
                }
                xml.LoadXml(inf.FIELD_TEXT);
                doc = xml.DocumentElement;
                DataTable dtfield = _data.Tables["field"];
                foreach (XmlElement ele in doc.ChildNodes)
                {
                    
                    DataRow newRow = dtfield.NewRow();
                    newRow["node"] = ele.GetAttribute("node");
                    newRow["name"] = ele.GetAttribute("name");
                    newRow["table"] = ele.GetAttribute("table");
                    newRow["nodeDesc"] = ele.GetAttribute("nodeDesc");
                    newRow["type"] = ele.GetAttribute("type");
                    dtfield.Rows.Add(newRow);
                }
                
                return _data;
            }
        }
        private DTO.LIST_QD_SCHEMAInfo GetDataFromForm(DTO.LIST_QD_SCHEMAInfo inf)
        {
            inf.CONN_ID = _db;
            inf.DEFAULT_CONN = ddlQD.Text;
            inf.DESCRIPTN = txtDescription.Text;
            inf.LOOK_UP = txtLookup.Text;
            inf.UPDATED = DateTime.Today.Year * 10000 + DateTime.Today.Month * 100 + DateTime.Today.Day;
            if (inf.FIELD_TEXT != "")
            {
                DataTable dtfield = _data.Tables["field"];
                string field = "<?xml version=\"1.0\" encoding=\"utf-8\" ?><SUN_SCHEMA>{0}</SUN_SCHEMA>";
                string tmp = "";
                string from = "";
                foreach (DataRow jrow in dtfield.Rows)
                {
                    tmp += string.Format("<row table=\"{0}\" node=\"{1}\" name=\"{2}\" nodeDesc=\"{3}\" type=\"{4}\" conn_id=\"{5}\"/>", jrow["table"], jrow["node"], jrow["name"], jrow["nodeDesc"], jrow["type"], inf.DEFAULT_CONN);
                }
                DataTable dtfrom = _data.Tables["fromcode"];
                field = string.Format(field, tmp);
                tmp = "";
                foreach (DataRow jrow in dtfrom.Rows)
                {
                    tmp += string.Format("<row fromcode=\"{0}\" lookup=\"{1}\" /> ", jrow["fromcode"], jrow["lookup"]);
                }
                from = tmp;

                //result = doc.InnerXml;
                inf.FIELD_TEXT = field;
                inf.FROM_TEXT = from;

            }
            else
            {
                StringBuilder sb = new StringBuilder();
                StringWriter str = new StringWriter(sb);
                _data.WriteXml(str);
                inf.FROM_TEXT = sb.ToString();
            }
            //inf.FIELD_TEXT = GetFieldCode(_data.Tables["field"]);
            inf.SCHEMA_ID = txtCode.Text;
            inf.SCHEMA_STATUS = ckbUse.Checked ? "Y" : "N";
            inf.DAG = Group.Text;
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
            DataTable dtfrom = new DataTable("fromcode");
            DataColumn[] colfrom = new DataColumn[] { new DataColumn("fromcode"), new DataColumn("lookup") };
            dtfrom.Columns.AddRange(colfrom);

            DataTable dtfield = new DataTable("field");
            DataColumn[] colfield = new DataColumn[] { new DataColumn("node")
            , new DataColumn("table")
            , new DataColumn("name")
            , new DataColumn("nodeDesc")
            , new DataColumn("type")};
            dtfield.Columns.AddRange(colfield);

            _data.Tables.Add(dtfrom);
            _data.Tables.Add(dtfield);
            DataRelation relation = new DataRelation("R_field", dtfrom.Columns["fromcode"], dtfield.Columns["table"], true);
            _data.Relations.Add(relation);
            bsFROMCODE.DataSource = dtfrom;
            bsField.DataSource = dtfield;
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
                    txtCode.Text = tableName;
                    _code = txtCode.Text.Trim();
                    DataTable kq = new DataTable();
                    System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(frm.Connect);
                    try
                    {
                        conn.Open();
                        kq = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Columns, new Object[] { null, null, tableName, null });
                        //dgvField.AutoGenerateColumns = true;
                        _data.Tables["field"].Clear();
                        _data.Tables["fromcode"].Clear();
                        DataRow rowtable = _data.Tables["fromcode"].NewRow();
                        rowtable["fromcode"] = rowtable["lookup"] = tableName;
                        _data.Tables["fromcode"].Rows.Add(rowtable);
                        foreach (DataRow row in kq.Rows)
                        {
                            DataRow newRow = _data.Tables["field"].NewRow();
                            newRow["table"] = tableName;
                            newRow["node"] = newRow["name"] = row["COLUMN_NAME"];
                            if (row["DATA_TYPE"].ToString() == "135")
                                newRow["type"] = "D";
                            else if (row["DATA_TYPE"].ToString() == "5")
                                newRow["type"] = "N";
                            _data.Tables["field"].Rows.Add(newRow);
                        }
                        //dgvField.DataSource = kq;
                    }
                    catch { }
                    finally { conn.Close(); }
                    dgvField.AutoSizeColumns();
                    dgvFrom.AutoSizeColumns();
                }

            }
        }

        private void btnSelectView_Click(object sender, EventArgs e)
        {
            LookupDB("View List");
            //frmLookup frm = new frmLookup();
            //frm.Text = "View List";
            //frm.Connect = txtConnect.Text;
            //if (frm.ShowDialog() == DialogResult.OK)
            //{
            //    string tableName = frm.ReturnCode;
            //    if (tableName != "")
            //    {
            //        txtCode.Text = tableName;
            //        _code = txtCode.Text.Trim();
            //        DataTable kq = new DataTable();
            //        System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(frm.Connect);
            //        try
            //        {
            //            conn.Open();
            //            kq = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Columns, new Object[] { null, null, tableName, null });
            //            //dgvField.AutoGenerateColumns = true;
            //            _data.Tables["field"].Clear();
            //            _data.Tables["fromcode"].Clear();
            //            DataRow rowtable = _data.Tables["fromcode"].NewRow();
            //            rowtable["fromcode"] = rowtable["lookup"] = tableName;
            //            _data.Tables["fromcode"].Rows.Add(rowtable);
            //            foreach (DataRow row in kq.Rows)
            //            {
            //                DataRow newRow = _data.Tables["field"].NewRow();
            //                newRow["table"] = tableName;
            //                newRow["node"] = newRow["name"] = row["COLUMN_NAME"];
            //                _data.Tables["field"].Rows.Add(newRow);
            //            }
            //            //dgvField.DataSource = kq;
            //        }
            //        catch { }
            //        finally { conn.Close(); }
            //    }
            //}

        }

        private void dgvField_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void btnRelation_Click(object sender, EventArgs e)
        {
            string sErr = "";
            frmQDADDView frmview = new frmQDADDView(_db, _user);

            frmview.Conn_ID = ddlQD.Text;

            if (frmview.ShowDialog() == DialogResult.OK && frmview.ReturnCode != "")
            {
                frmRelation frm = new frmRelation();
                frm.DTOriginal = _data.Tables["field"];
                BUS.LIST_QD_SCHEMAControl ctr = new BUS.LIST_QD_SCHEMAControl();
                DTO.LIST_QD_SCHEMAInfo inf = ctr.Get(_db, frmview.ReturnCode, ref sErr);
                if (inf.SCHEMA_ID != "")
                {
                    try
                    {
                        string result = "<?xml version=\"1.0\" encoding=\"utf-8\" ?><SUN_SCHEMA></SUN_SCHEMA>";
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(result);
                        XmlElement docele = doc.DocumentElement;

                        //<row table="M5" node="Lookup" name="Lookup" type=""/>  	
                        string schema = inf.FROM_TEXT;
                        DataSet dset = new DataSet("Schema");
                        //DataTable dtfrom = new DataTable("fromcode");
                        //DataColumn[] colfrom = new DataColumn[] { new DataColumn("fromcode"), new DataColumn("lookup") };
                        //dtfrom.Columns.AddRange(colfrom);

                        DataTable dtfield = new DataTable("field");
                        DataColumn[] colfield = new DataColumn[] { new DataColumn("node")
                    , new DataColumn("table")
                    , new DataColumn("name")
                    , new DataColumn("nodeDesc")
                    , new DataColumn("type")};
                        dtfield.Columns.AddRange(colfield);

                        //dset.Tables.Add(dtfrom);
                        dset.Tables.Add(dtfield);
                        //DataRelation relation = new DataRelation("R_field", dtfrom.Columns["fromcode"], dtfield.Columns["table"], true);
                        //dset.Relations.Add(relation);
                        StringReader strR = new StringReader(schema);
                        dset.ReadXml(strR);
                        frm.DTLooup = dset.Tables["field"];
                        if (frm.ShowDialog() == DialogResult.OK)
                        {
                            DataTable dtR = frm.Relation;
                            DataRow fromRow = _data.Tables["fromcode"].NewRow();
                            DataRow fieldRow = _data.Tables["field"].NewRow();
                            fieldRow["node"] = frmview.ReturnCode;
                            fieldRow["name"] = frmview.ReturnCode + "Record";
                            fieldRow["type"] = "S";
                            fieldRow["table"] = txtCode.Text.Trim();
                            fromRow["fromcode"] = txtCode.Text.Trim() + "\\" + frmview.ReturnCode;
                            string lookup = "";
                            foreach (DataRow row in dtR.Rows)
                            {
                                if (row["Original"].ToString().Contains(" "))
                                    row["Original"] = "[" + row["Original"].ToString() + "]";
                                if (row["lookup"].ToString().Contains(" "))
                                    row["lookup"] = "[" + row["lookup"].ToString() + "]";
                                lookup += " and " + txtCode.Text.Trim() + "." + row["Original"].ToString() + " = " + frmview.ReturnCode + "." + row["lookup"].ToString();
                            }
                            fromRow["lookup"] = lookup.Substring(5);
                            _data.Tables["fromcode"].Rows.Add(fromRow);
                            if (dgvField.CurrentRow != null && dgvField.CurrentRow.RowIndex >= 0)
                            {
                                _data.Tables["field"].Rows.InsertAt(fieldRow, dgvField.CurrentRow.RowIndex);
                            }
                            else
                                _data.Tables["field"].Rows.Add(fieldRow);
                        }
                    }
                    catch (Exception ex) { lbErr.Text = ex.Message; }
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
            frmQDADDView frm = new frmQDADDView(_db, _user);
            //frm.Connect = _db;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (frm.ReturnCode != "")
                {
                    BUS.LIST_QD_SCHEMAControl ctr = new BUS.LIST_QD_SCHEMAControl();
                    DTO.LIST_QD_SCHEMAInfo inf = ctr.Get(_db, frm.ReturnCode, ref sErr);
                    SetDataToForm(inf);
                }
            }
            if (sErr == "")
            {
                EnableForm(false);
                _processStatus = "V";
            }
            dgvFrom.AutoSizeColumns();
            dgvField.AutoSizeColumns();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            BUS.LIST_QD_SCHEMAControl ctr = new BUS.LIST_QD_SCHEMAControl();
            if (ctr.IsExist(_db, txtCode.Text))
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
            BUS.LIST_QD_SCHEMAControl ctr = new BUS.LIST_QD_SCHEMAControl();
            DTO.LIST_QD_SCHEMAInfo inf = new DTO.LIST_QD_SCHEMAInfo();

            if (_processStatus == "C")
            {
                if (!ctr.IsExist(_db, txtCode.Text))
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
            BUS.LIST_QD_SCHEMAControl ctr = new BUS.LIST_QD_SCHEMAControl();
            if (ctr.IsExist(_db, txtCode.Text))
            {
                if (MessageBox.Show("Do you want to delete " + txtCode.Text + " schema?", "Message", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    string sErr = ctr.Delete(_db, txtCode.Text);
                    RefreshForm("");
                    EnableForm(false);
                }
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            BUS.LIST_QD_SCHEMAControl ctr = new BUS.LIST_QD_SCHEMAControl();
            if (ctr.IsExist(_db, txtCode.Text))
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
            FrmTransferIn frm = new FrmTransferIn("QDADD");
            frm.ShowDialog();
        }

        private void btnTransferOut_Click(object sender, EventArgs e)
        {
            FrmTransferOut frm = new FrmTransferOut(_db, "QDADD");
            frm.QD_CODE = txtCode.Text;
            //frm.DTB = _db;
            frm.ShowDialog();
            //BUS.LIST_QD_SCHEMAControl ctr = new BUS.LIST_QD_SCHEMAControl();
            ////if (ctr.IsExist(ddlQD.Text, txtCode.Text))
            ////{
            ////DTO.LIST_QD_SCHEMAInfo inf = new DTO.LIST_QD_SCHEMAInfo();
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

        private void txtCode_TextChanged(object sender, EventArgs e)
        {
            bool flag = false;
            for (int i = 0; i < dgvFrom.RowCount; i++)
            {
                dgvFrom.GetRow(i).BeginEdit();
                if (dgvFrom.GetRow(i).Cells["fromcode"].Value != null)
                {
                    string code = dgvFrom.GetRow(i).Cells["fromcode"].Value.ToString().Trim();
                    if (code == _code)
                    {
                        dgvFrom.GetRow(i).Cells["fromcode"].Value = txtCode.Text;
                        flag = true;
                    }
                    else if (flag)
                    {
                        dgvFrom.GetRow(i).Cells["fromcode"].Value = code.Replace(_code + "\\", txtCode.Text.Trim() + "\\");
                        string lookup = dgvFrom.GetRow(i).Cells["lookup"].Value.ToString();
                        dgvFrom.GetRow(i).Cells["lookup"].Value = lookup.Replace(_code + ".", txtCode.Text.Trim() + ".");
                    }
                }
                dgvFrom.GetRow(i).EndEdit();
            }
            _code = txtCode.Text.Trim();


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
                    DataRow rowmove = _data.Tables["field"].NewRow();
                    rowmove.ItemArray = _data.Tables["field"].Rows[indexSource].ItemArray;
                    _data.Tables["field"].Rows.RemoveAt(indexSource);
                    _data.Tables["field"].Rows.InsertAt(rowmove, rowIndex);
                }
            }
        }


        private void dgvFrom_DragDrop(object sender, DragEventArgs e)
        {
            Janus.Windows.GridEX.GridEXRow row = (Janus.Windows.GridEX.GridEXRow)e.Data.GetData(typeof(Janus.Windows.GridEX.GridEXRow));
            if (row != null)
            {
                int indexSource = row.RowIndex;
                Point p = dgvFrom.PointToClient(new Point(e.X, e.Y));
                int rowIndex = dgvFrom.RowPositionFromPoint(p.X, p.Y);
                if (rowIndex >= 0 && rowIndex < dgvFrom.RowCount && indexSource != rowIndex)
                {
                    DataRow rowmove = _data.Tables["fromcode"].NewRow();
                    rowmove.ItemArray = _data.Tables["fromcode"].Rows[indexSource].ItemArray;
                    _data.Tables["fromcode"].Rows.RemoveAt(indexSource);
                    _data.Tables["fromcode"].Rows.InsertAt(rowmove, rowIndex);
                }
            }
        }

        private void dgvFrom_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(Janus.Windows.GridEX.GridEXRow)))
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        private void dgvFrom_MouseDown(object sender, MouseEventArgs e)
        {
            int index = dgvFrom.RowPositionFromPoint(e.X, e.Y);
            if (index > -1 && dgvFrom.ColumnFromPoint(e.X, e.Y) == null && dgvFrom.AllowEdit == Janus.Windows.GridEX.InheritableBoolean.True)
            {
                if (dgvFrom.GetRow(index).RowType == Janus.Windows.GridEX.RowType.Record)
                    dgvFrom.DoDragDrop(dgvFrom.GetRow(index), DragDropEffects.Move);
            }
        }

        private void txtConnect_TextChanged(object sender, EventArgs e)
        {
            //int indexS = txtConnect.Text.IndexOf("Password=", 0);
            //if (indexS != -1)
            //{
            //    int indexE = txtConnect.Text.IndexOf(";", indexS + 9);
            //    string kq = txtConnect.Text.Substring(0, indexS + 9) + "*****" + txtConnect.Text.Substring(indexE);
            //    txtConectEx.Text = kq;
            //}
            //else
            //    txtConectEx.Text = "Password=*****;" + txtConnect.Text;
            txtConectEx.Text = BUS.CommonControl.HiddenAttribute(txtConnect.Text, "Password");
        }

        private void lbErr_Click(object sender, EventArgs e)
        {
            MessageBox.Show(lbErr.Text);
        }

        private void btnXML_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog frm = new FolderBrowserDialog();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                string[] arrFilename = Directory.GetFiles(frm.SelectedPath);
                int fromIndex = -1;
                for (int j = 0; j < arrFilename.Length; j++)
                    if (Path.GetFileNameWithoutExtension(arrFilename[j]).ToUpper() == "FROM")
                        fromIndex = j;
                if (fromIndex != -1)
                    for (int i = 0; i < arrFilename.Length; i++)
                    {
                        if (fromIndex != i)
                        {
                        }
                    }
            }
        }

        private void dgvAddRow_Click(object sender, EventArgs e)
        {
            _data.Tables["field"].Rows.InsertAt(_data.Tables["field"].NewRow(), dgvField.CurrentRow.RowIndex);
        }

        private void btnQD_Click(object sender, EventArgs e)
        {
            frmDAView frm = new frmDAView();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                Group.Text = frm.Code;
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            string sErr = "";
            BUS.LIST_QD_SCHEMAControl ctr = new BUS.LIST_QD_SCHEMAControl();
            DataTable dt = ctr.GetAll(_db, ref sErr);
            foreach (DataRow row in dt.Rows)
            {
                DTO.LIST_QD_SCHEMAInfo inf = new DTO.LIST_QD_SCHEMAInfo(row);
                try
                {
                    _data.Tables["field"].Rows.Clear();
                    _data.Tables["fromcode"].Rows.Clear();
                    _data= ReadScheme(inf);
                    DataTable dtfield = _data.Tables["field"];
                    string field = "<?xml version=\"1.0\" encoding=\"utf-8\" ?><SUN_SCHEMA>{0}</SUN_SCHEMA>";
                    string tmp = "";
                    string from = "";
                    foreach (DataRow jrow in dtfield.Rows)
                    {
                        tmp += string.Format("<row table=\"{0}\" node=\"{1}\" name=\"{2}\" nodeDesc=\"{3}\" type=\"{4}\" conn_id=\"{5}\"/>", jrow["table"], jrow["node"], jrow["name"], jrow["nodeDesc"], jrow["type"], inf.DEFAULT_CONN);
                    }
                    DataTable dtfrom = _data.Tables["fromcode"];
                    field = string.Format(field, tmp);
                    tmp = "";
                    foreach (DataRow jrow in dtfrom.Rows)
                    {
                        tmp += string.Format("<row fromcode=\"{0}\" lookup=\"{1}\"/> ", jrow["fromcode"], jrow["lookup"]);
                    }
                    from = tmp;

                    //result = doc.InnerXml;
                    inf.FIELD_TEXT = field;
                    inf.FROM_TEXT = from;
                    ctr.Update(inf);
                }
                catch (Exception ex)
                { }
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            QueryBuilder.SchemaDefinition.InvalidateCache();
        }


    }
}
