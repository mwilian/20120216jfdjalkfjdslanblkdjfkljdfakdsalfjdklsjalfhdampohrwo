using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FlexCel.XlsAdapter;
using System.IO;
using FlexCel.Core;
using System.Collections;
using System.Xml;
using BUS;

namespace dCube
{
    public partial class frmImport : Form
    {
        string _dtb = "";

        public string DTB
        {
            get { return _dtb; }
            set { _dtb = value; }
        }
        string _sErr = "";
        public frmImport(string db)
        {
            InitializeComponent();
            _dtb = db;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ofdImport.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofdImport.FileName;
            }
        }

        /* public DataTable PreviewExcelFile(DataTable dt, string filename, string fontCode)
         {
             //DataTable dt = new DataTable();
             if (File.Exists(filename))
             {
                 XlsFile xlsFile = new XlsFile(textBox1.Text, true);
                 ArrayList arrRange = new ArrayList();

                 int maxCol = 1;
                 int maxRow = 1;
                 for (int i = 0; i < xlsFile.NamedRangeCount; i++)
                 {
                     string name = xlsFile.GetNamedRange(i + 1)._Name;
                     if (name.Length > 3 && (name.Substring(0, 3) == "d__" || name.Substring(0, 3) == "f__") && (dt.Columns.Contains(name.Substring(3))))
                     {
                         TXlsNamedRange range = xlsFile.GetNamedRange(i + 1);
                         //if (maxCol < range.ColCount) maxCol = range.ColCount;
                         if (maxRow < range.RowCount) maxCol = range.RowCount;
                         if (name.Substring(0, 3) == "d__")
                         {
                             arrRange.Insert(0, range);
                         }
                         else
                         {
                             if (maxCol < range.ColCount) maxCol = range.ColCount;
                             arrRange.Add(range);
                         }

                         //if (name.Substring(3) == "TimesheetDate")
                         //dt.Columns.Add(name.Substring(3), Type.GetType("System." + dgvList.RootTable.Columns[name.Substring(3)].DataTypeCode.ToString()));
                         //else
                         //    dt.Columns.Add(name.Substring(3));

                     }
                 }
                 if (arrRange.Count == 0)
                     return dt;
                 DataRow newRow;
                 int index_c = 0;
                 for (int i_d = 0; i_d < ((TXlsNamedRange)arrRange[0]).RowCount; i_d++)
                 {
                     bool flag = true;

                     while (flag)
                     {
                         int top = ((TXlsNamedRange)arrRange[0]).Top + i_d;
                         if (top >= ((TXlsNamedRange)arrRange[0]).RowCount + ((TXlsNamedRange)arrRange[0]).Top)
                             top = ((TXlsNamedRange)arrRange[0]).Top;
                         int left = ((TXlsNamedRange)arrRange[0]).Left + index_c;
                         if (left >= ((TXlsNamedRange)arrRange[0]).ColCount + ((TXlsNamedRange)arrRange[0]).Left)
                         {
                             left = ((TXlsNamedRange)arrRange[0]).Left;
                         }
                         object dObject = xlsFile.GetCellValue(top, left);
                         if (dObject != null && dObject.ToString() != "")
                         {
                             newRow = dt.NewRow();
                             for (int i = 0; i < arrRange.Count; i++)
                             {
                                 if (dt.Columns.Contains(((TXlsNamedRange)arrRange[i])._Name.Substring(3)))
                                 {
                                     //newRow[((TXlsNamedRange)arrRange[0])._Name.Substring(3)] = dObject;
                                     if (((TXlsNamedRange)arrRange[i]).IsOneCell)
                                     {
                                         if (dgvList.RootTable.Columns[((TXlsNamedRange)arrRange[i])._Name.Substring(3)].DataTypeCode == TypeCode.DateTime)
                                             newRow[((TXlsNamedRange)arrRange[i])._Name.Substring(3)] = FlxDateTime.FromOADate((double)xlsFile.GetCellValue(((TXlsNamedRange)arrRange[i]).Top, ((TXlsNamedRange)arrRange[i]).Left), false);
                                         else
                                             newRow[((TXlsNamedRange)arrRange[i])._Name.Substring(3)] = GetObject(xlsFile.GetCellValue(((TXlsNamedRange)arrRange[i]).Top, ((TXlsNamedRange)arrRange[i]).Left), fontCode);
                                     }
                                     else
                                     {
                                         top = ((TXlsNamedRange)arrRange[i]).Top + i_d;
                                         if (top >= ((TXlsNamedRange)arrRange[i]).RowCount + ((TXlsNamedRange)arrRange[i]).Top)
                                             top = ((TXlsNamedRange)arrRange[i]).Top;
                                         left = ((TXlsNamedRange)arrRange[i]).Left + index_c;
                                         if (left >= ((TXlsNamedRange)arrRange[i]).ColCount + ((TXlsNamedRange)arrRange[i]).Left)
                                         {
                                             left = ((TXlsNamedRange)arrRange[i]).Left;
                                         }
                                         if (dgvList.RootTable.Columns.Contains(((TXlsNamedRange)arrRange[i])._Name.Substring(3)))
                                         {
                                             try
                                             {
                                                 if (dgvList.RootTable.Columns[((TXlsNamedRange)arrRange[i])._Name.Substring(3)].DataTypeCode == TypeCode.DateTime)
                                                     newRow[((TXlsNamedRange)arrRange[i])._Name.Substring(3)] = FlxDateTime.FromOADate((double)xlsFile.GetCellValue(top, left), false);
                                                 else
                                                     newRow[((TXlsNamedRange)arrRange[i])._Name.Substring(3)] = GetObject(xlsFile.GetCellValue(top, left), fontCode);

                                             }
                                             catch { }
                                         }
                                     }
                                 }
                             }
                             dt.Rows.Add(newRow);
                             index_c++;
                             if (index_c < maxCol)
                                 flag = true;
                             else
                             {
                                 index_c = 0;
                                 flag = false;
                             }
                         }
                         else break;

                     }
                 }


             }
             return dt;
         }*/


        private void btnPreview_Click(object sender, EventArgs e)
        {
            _flagimport = true;
            //button2.Enabled = false;
            if (ddlImport.Text == "")
            {
                MessageBox.Show("Please choose a Import Code");
                return;
            }
            string filename = textBox1.Text;
            string sErr = "";
            for (int i = 0; i < tcMain.TabPages.Count; i++)
            {

                TabPage page = tcMain.TabPages[i];
                if (page.Controls.Count == 1 && page.Controls[0] is DataGridView)
                {
                    DataGridView dgv = page.Controls[0] as DataGridView;

                    if (dgv.Tag is IMPORT_SCHEMAControl)
                    {
                        IMPORT_SCHEMAControl _importCtr = dgv.Tag as IMPORT_SCHEMAControl;
                        DataTable dtX = IMPORT_SCHEMAControl.GetDataTableStruct(_importCtr.DtStruct, _importCtr.Lookup);
                        try
                        {
                            DataTable dt = clsTransfer.PreviewExcelFile(dtX, filename, cboConvertor.Text, ddlImport.SelectedValue.ToString());
                            dgv.DataSource = dt;
                            lbErr.Text = "You have " + dt.Rows.Count + " records from file";
                            btnImport.Enabled = false;
                            btnGroup.Enabled = dt.Rows.Count > 0;
                        }
                        catch (Exception ex) { sErr += ex.Message; }
                    }
                }
                lbErr.Text = sErr;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            BUS.IMPORT_SCHEMAControl ctr = new BUS.IMPORT_SCHEMAControl();
            string sErr = "";
            DataTable dt = ctr.GetAll(_dtb, ref sErr);
            BUS.LIST_DAControl daCtr = new LIST_DAControl();
            DataTable dtPermision = daCtr.GetPermission(Form_QD._user, ref sErr);
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                DTO.IMPORT_SCHEMAInfo impInf = new DTO.IMPORT_SCHEMAInfo(dt.Rows[i]);
                string flag = "";
                bool ie = true;
                foreach (DataRow row in dtPermision.Rows)
                {
                    if (impInf.DAG.Trim() != "")
                    {
                        if (row["DAG_ID"].ToString() == impInf.DAG)
                        {
                            flag = row["EI"].ToString();
                        }
                        else if (row["EI"].ToString() == "I")
                        {
                            ie = false;
                        }
                    }
                }
                if ((flag == "" && ie) || flag == "I")
                {
                }
                else
                {
                    dt.Rows.Remove(dt.Rows[i]);
                }

            }
            ddlImport.DataSource = dt;
            ddlImport.ValueMember = "SCHEMA_ID";
            ddlImport.DisplayMember = "DESCRIPTN";
            cboConvertor.SelectedIndex = 0;
            //_importCtr.StrConn = Form_QD._strConnectDes;
        }

        private BUS.IMPORT_SCHEMAControl AddValidatedList(BUS.IMPORT_SCHEMAControl _importCtr, string db, string xml)
        {
            _importCtr.ListV.Clear();
            _importCtr.LKey.Clear();
            _importCtr.DtStruct = BUS.IMPORT_SCHEMAControl.GetStruct(xml);
            foreach (DataRow row in _importCtr.DtStruct.Rows)
            {
                if ((row["Tag"] != DBNull.Value && row["Tag"].ToString() != "") || (row["IsNull"] == DBNull.Value || row["IsNull"].ToString() == "False"))
                {
                    if (row["Tag"] != DBNull.Value && row["Tag"].ToString() != "")
                    {
                        objValidatedList objVal = new objValidatedList(row["Tag"]);
                        ValueList validate = new ValueList();
                        if (row["IsNull"] != DBNull.Value && row["IsNull"].ToString() == "True")
                            validate.IsNull = true;
                        else
                            validate.IsNull = false;
                        validate.Key = row["Key"].ToString();
                        validate.Message = objVal.Message;
                        BUS.LIST_QDControl ctr = new BUS.LIST_QDControl();
                        DTO.LIST_QDInfo inf = ctr.Get_LIST_QD(db, objVal.QD, ref _sErr);
                        QueryBuilder.SQLBuilder sqlB = QueryBuilder.SQLBuilder.LoadSQLBuilderFromDataBase(inf.QD_ID, inf.DTB, inf.ANAL_Q0);
                        sqlB.StrConnectDes = Form_QD._strConnectDes;
                        DataTable dt = sqlB.BuildDataTable(inf.SQL_TEXT);
                        foreach (DataRow aRow in dt.Rows)
                        {
                            if (!validate.Content.Contains(aRow[objVal.Field].ToString().Trim()))
                                validate.Content.Add(aRow[objVal.Field].ToString().Trim());
                        }
                        _importCtr.ListV.Add(validate);

                    }
                    else
                    {
                        ValueList validate = new ValueList();
                        //if (row["IsNull"] == DBNull._Value || row["IsNull"].ToString() == "False")
                        validate.IsNull = false;
                        validate.Key = row["Key"].ToString();
                        _importCtr.ListV.Add(validate);
                    }

                }
                if (row["PrimaryKey"].ToString() == "True")
                    _importCtr.LKey.Add(row["Key"].ToString());
            }
            return _importCtr;
        }
        //private void dgv_LoadingRow(object sender, Janus.Windows.GridEX.RowLoadEventArgs e)
        //{
        //    if (e.Row.RowType == RowType.Record)
        //        ValidatedRow(_importCtr, e.Row);
        //}
        bool _flagimport = true;
        private void ValidatedRow(IMPORT_SCHEMAControl _importCtr, DataGridViewRow gridEXRow)
        {
            bool flag = true;
            string sErr = "";
            foreach (DataGridViewCell cell in gridEXRow.Cells)
            {
                string message = "";
                cell.ErrorText = "";
                if (!_importCtr.ContrainList(gridEXRow.DataGridView.Columns[cell.ColumnIndex].Name, cell.Value, ref message))
                {
                    cell.ErrorText = message;
                    sErr += message;
                    //cell.ToolTipText = message;
                    flag = false;
                }
            }
            gridEXRow.ErrorText = sErr;
            _flagimport = _flagimport & flag;
        }

        private void dllImport_ValueChanged(object sender, EventArgs e)
        {

        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            string sErr = "";
            if (MessageBox.Show("Do you want to import these correct records?", "Warning", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                int dem = 0;
                for (int i = 0; i < tcMain.TabPages.Count; i++)
                {

                    TabPage page = tcMain.TabPages[i];
                    if (page.Controls.Count == 1 && page.Controls[0] is DataGridView)
                    {

                        DataGridView dgv = page.Controls[0] as DataGridView;
                        IMPORT_SCHEMAControl _importCtr = dgv.Tag as IMPORT_SCHEMAControl;

                        DataTable dt = dgv.DataSource as DataTable;
                        for (int j = dgv.RowCount - 1; j >= 0; j--)
                        {
                            DataGridViewRow row = dgv.Rows[j];
                            if (row.DataBoundItem is DataRowView)
                            {
                                DataRow dtrow = ((DataRowView)row.DataBoundItem).Row;
                                string tmp = "";
                                int result = _importCtr.Import(dtrow, checkBox1.Checked, checkBox2.Checked, ref tmp);
                                if (!sErr.Contains(tmp))
                                    sErr += tmp;

                                if (result == 1)
                                {
                                    dt.Rows.Remove(dtrow);
                                    dem++;
                                }
                            }

                        }
                    }
                }
                btnImport.Enabled = false;
                btnGroup.Enabled = false;
                if (sErr == "")
                    lbErr.Text = "Have " + dem + " update records";
                else
                {
                    lbErr.Text = sErr;
                }
            }
            //}
            //_importCtr.Import(dgvList.DataSource as DataTable, checkBox1.Checked, checkBox2.Checked);
        }

        private void lbErr_Click(object sender, EventArgs e)
        {
            MessageBox.Show(lbErr.Text);
        }

        private void btnGroup_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < tcMain.TabPages.Count; i++)
            {

                TabPage page = tcMain.TabPages[i];
                if (page.Controls.Count == 1 && page.Controls[0] is DataGridView)
                {
                    DataGridView dgv = page.Controls[0] as DataGridView;
                    IMPORT_SCHEMAControl _importCtr = dgv.Tag as IMPORT_SCHEMAControl;
                    if (dgv.RowCount > 0)
                    {
                        DataSet dset = null;
                        DataTable dt = dgv.DataSource as DataTable;

                        DataSetHelper dsHelper = new DataSetHelper(ref dset);
                        string strField = "";
                        string filter = "";
                        string groupField = "";
                        bool flag = false;
                        foreach (DataRow row in _importCtr.DtStruct.Rows)
                        {
                            if (row["AggregateFunction"] != DBNull.Value && row["AggregateFunction"].ToString() != "")
                            {
                                strField += "," + row["AggregateFunction"].ToString().Trim().ToLower() + "(" + row["Key"].ToString() + ") " + row["Key"].ToString();
                                flag = true;
                            }
                            else
                            {
                                strField += "," + row["Key"].ToString();
                                groupField += "," + row["Key"].ToString();
                            }
                        }
                        strField = strField.Substring(1);
                        groupField = groupField.Substring(1);
                        if (flag)
                        {
                            DataTable dtgroup = dsHelper.SelectGroupByInto("Group", dt, strField, filter, groupField);
                            dgv.DataSource = dtgroup;
                            lbErr.Text = "You have " + dtgroup.Rows.Count + " records by Group";
                            btnImport.Enabled = dtgroup.Rows.Count > 0;
                        }
                        else btnImport.Enabled = dt.Rows.Count > 0;

                    }
                }
            }

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            tcMain.TabPages.Clear();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (ddlImport.Text == "")
            {
                MessageBox.Show("Please choose a Import Code");
                return;
            }

            string code = ddlImport.SelectedValue.ToString();
            string text = ddlImport.Text;
            if (!tcMain.TabPages.ContainsKey(code))
            {
                DataRow row = ((DataRowView)ddlImport.SelectedItem).Row;
                DTO.IMPORT_SCHEMAInfo importInf = new DTO.IMPORT_SCHEMAInfo(row);
                IMPORT_SCHEMAControl ctr = new IMPORT_SCHEMAControl();

                string key = importInf.DEFAULT_CONN;
                ctr.StrConn = Form_QD.Config.GetConnection(ref key, "AP");
                ctr.Lookup = row["LOOK_UP"].ToString();
                ctr = AddValidatedList(ctr, row["DB"].ToString(), row["FIELD_TEXT"].ToString());
                //StringReader sb = new StringReader(row["FIELD_TEXT"].ToString());
                byte[] byteArray = Encoding.ASCII.GetBytes(row["FIELD_TEXT"].ToString());
                MemoryStream stream = new MemoryStream(byteArray);
                stream.Close();
                DataGridView dgv = new DataGridView();
                dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                dgv.Name = "dgv" + code;
                dgv.Dock = System.Windows.Forms.DockStyle.Fill;
                dgv.Location = new System.Drawing.Point(3, 3);
                dgv.Size = new System.Drawing.Size(1058, 325);
                dgv.AllowUserToAddRows = false;
                dgv.ReadOnly = true;
                //tcMain.TabPages[tcMain.TabPages.Count - 1].Tag = row["LOOK_UP"].ToString();
                dgv.Tag = ctr;
                dgv.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(dgv_DataBindingComplete);

                tcMain.TabPages.Add(code, text);
                tcMain.TabPages[tcMain.TabPages.Count - 1].Controls.Add(dgv);
            }

        }

        void dgv_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (dgv.Tag is IMPORT_SCHEMAControl)
            {
                IMPORT_SCHEMAControl _importCtr = dgv.Tag as IMPORT_SCHEMAControl;
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    ValidatedRow(_importCtr, row);
                }
            }
        }
    }

}
