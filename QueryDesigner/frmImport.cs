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
using Janus.Windows.GridEX;
using System.Xml;
using BUS;

namespace dCube
{
    public partial class frmImport : Form
    {
        BUS.IMPORT_SCHEMAControl _importCtr = new BUS.IMPORT_SCHEMAControl();
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
        public DataTable PreviewExcelFile(DataTable dt, string filename, string fontCode, string importCode)
        {
            //DataTable dt = new DataTable();
            if (File.Exists(filename))
            {
                XlsFile xlsFile = new XlsFile(textBox1.Text, true);
                List<TXlsNamedRange> arrRange = new List<TXlsNamedRange>();

                int maxCol = 1;
                int maxRow = 1;
                for (int i = 0; i < xlsFile.SheetCount; i++)
                {
                    TXlsNamedRange range = xlsFile.GetNamedRange("importData", i);

                    if (range != null)
                    {
                        xlsFile.ActiveSheet = range.SheetIndex;
                        for (int r = 1; r < range.RowCount - 1; r++)
                        {
                            DataRow newRow = dt.NewRow();
                            for (int c = 0; c < range.ColCount; c++)
                            {
                                if (xlsFile.GetCellValue(range.Top, range.Left + c) != null && dt.Columns.Contains(xlsFile.GetCellValue(range.Top, range.Left + c).ToString()))
                                {
                                    GetValue(ref newRow, xlsFile.GetCellValue(range.Top, range.Left + c).ToString(), xlsFile, range.Top + r, range.Left + c, fontCode);
                                }
                            }
                            dt.Rows.Add(newRow);
                        }
                    }

                }
                if (dt.Rows.Count > 0)
                    return dt;
                InitAdvance(dt, xlsFile, arrRange, ref maxCol, ref maxRow);
                if (arrRange.Count == 0)
                    return dt;
                GetAdvance(dt, fontCode, xlsFile, arrRange, maxCol, maxRow);
            }
            return dt;
        }

        private void GetAdvance(DataTable dt, string fontCode, XlsFile xlsFile, List<TXlsNamedRange> arrRange, int maxCol, int maxRow)
        {
            for (int i = 0; i < maxCol; i++)
            {
                for (int j = 0; j < maxRow; j++)
                {
                    DataRow newRow = dt.NewRow();
                    bool flag = true;
                    for (int index = 0; index < arrRange.Count; index++)
                    {
                        xlsFile.ActiveSheet = arrRange[index].SheetIndex;
                        int top = Math.Min(j, arrRange[index].RowCount - 1) + arrRange[index].Top;
                        int left = Math.Min(i, arrRange[index].ColCount - 1) + arrRange[index].Left;
                        if (index == 0)
                        {
                            int atop = top;
                            int aleft = left;
                            if (!xlsFile.CellMergedBounds(top, left).IsOneCell)
                            {
                                atop = xlsFile.CellMergedBounds(top, left).Top;
                                aleft = xlsFile.CellMergedBounds(top, left).Left;
                            }
                            object dObject = GetObject(xlsFile.GetCellValue(atop, aleft), fontCode);
                            if (dObject == null || dObject.ToString() == "" || dObject == DBNull.Value)
                            {
                                flag = false;
                                break;
                            }
                            else
                            {
                                GetValue(ref newRow, arrRange[index].Name.Substring(3), xlsFile, top, left, fontCode);
                            }
                        }
                        else
                        {
                            GetValue(ref newRow, arrRange[index].Name.Substring(3), xlsFile, top, left, fontCode);
                        }

                    }
                    if (flag)
                        dt.Rows.Add(newRow);
                }
            }
        }

        private static void InitAdvance(DataTable dt, XlsFile xlsFile, List<TXlsNamedRange> arrRange, ref int maxCol, ref int maxRow)
        {
            for (int i = 0; i < xlsFile.NamedRangeCount; i++)
            {
                string name = xlsFile.GetNamedRange(i + 1).Name;
                if (name.Length > 3 && (name.Substring(0, 3) == "d__" || name.Substring(0, 3) == "f__") && (dt.Columns.Contains(name.Substring(3))))
                {
                    TXlsNamedRange range = xlsFile.GetNamedRange(i + 1);
                    if (maxCol < range.ColCount) maxCol = range.ColCount;
                    if (maxRow < range.RowCount) maxRow = range.RowCount;
                    if (name.Substring(0, 3) == "d__")
                    {
                        arrRange.Insert(0, range);
                    }
                    else
                    {
                        //if (maxCol < range.ColCount) maxCol = range.ColCount;
                        arrRange.Add(range);
                    }

                    //if (name.Substring(3) == "TimesheetDate")
                    //dt.Columns.Add(name.Substring(3), Type.GetType("System." + dgvList.RootTable.Columns[name.Substring(3)].DataTypeCode.ToString()));
                    //else
                    //    dt.Columns.Add(name.Substring(3));

                }
            }
        }

        private void GetValue(ref DataRow newRow, string fieldName, XlsFile xlsFile, int top, int left, string fontCode)
        {
            try
            {
                if (!xlsFile.CellMergedBounds(top, left).IsOneCell)
                {
                    top = xlsFile.CellMergedBounds(top, left).Top;
                    left = xlsFile.CellMergedBounds(top, left).Left;
                }
                if (newRow.Table.Columns[fieldName].DataType == typeof(DateTime))
                    newRow[fieldName] = GetDateTimeObj(xlsFile, top, left);
                else
                    newRow[fieldName] = GetObject(xlsFile.GetCellValue(top, left), fontCode);
            }
            catch { throw new Exception("Object(" + xlsFile.ActiveSheetByName + "," + top + "," + left + ") is not valided"); }
        }

        private static object GetDateTimeObj(XlsFile xlsFile, int top, int left)
        {
            if (xlsFile.GetCellVisibleFormatDef(top, left).Format != "")
            {
                if (xlsFile.GetCellValue(top, left) != null || xlsFile.GetCellValue(top, left).ToString() == "" || xlsFile.GetCellValue(top, left).ToString() == "NULL")
                    return FlxDateTime.FromOADate((double)xlsFile.GetCellValue(top, left), false);
                else return DBNull.Value;
            }
            else
            {
                Object obj = xlsFile.GetCellValue(top, left);
                if (obj is double)
                {
                    int sunDate = Convert.ToInt32(obj);
                    int year = sunDate / 10000;
                    int month = (sunDate - year * 10000) / 100;
                    int day = sunDate - year * 10000 - month * 100;
                    return new DateTime(year, month, day);
                }
                else
                    return DBNull.Value;
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
                     string name = xlsFile.GetNamedRange(i + 1).Name;
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
                                 if (dt.Columns.Contains(((TXlsNamedRange)arrRange[i]).Name.Substring(3)))
                                 {
                                     //newRow[((TXlsNamedRange)arrRange[0]).Name.Substring(3)] = dObject;
                                     if (((TXlsNamedRange)arrRange[i]).IsOneCell)
                                     {
                                         if (dgvList.RootTable.Columns[((TXlsNamedRange)arrRange[i]).Name.Substring(3)].DataTypeCode == TypeCode.DateTime)
                                             newRow[((TXlsNamedRange)arrRange[i]).Name.Substring(3)] = FlxDateTime.FromOADate((double)xlsFile.GetCellValue(((TXlsNamedRange)arrRange[i]).Top, ((TXlsNamedRange)arrRange[i]).Left), false);
                                         else
                                             newRow[((TXlsNamedRange)arrRange[i]).Name.Substring(3)] = GetObject(xlsFile.GetCellValue(((TXlsNamedRange)arrRange[i]).Top, ((TXlsNamedRange)arrRange[i]).Left), fontCode);
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
                                         if (dgvList.RootTable.Columns.Contains(((TXlsNamedRange)arrRange[i]).Name.Substring(3)))
                                         {
                                             try
                                             {
                                                 if (dgvList.RootTable.Columns[((TXlsNamedRange)arrRange[i]).Name.Substring(3)].DataTypeCode == TypeCode.DateTime)
                                                     newRow[((TXlsNamedRange)arrRange[i]).Name.Substring(3)] = FlxDateTime.FromOADate((double)xlsFile.GetCellValue(top, left), false);
                                                 else
                                                     newRow[((TXlsNamedRange)arrRange[i]).Name.Substring(3)] = GetObject(xlsFile.GetCellValue(top, left), fontCode);

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

        private object GetObject(object p, string fontCode)
        {
            //object p = null;
            if (p == null || p.ToString().Trim().TrimStart() == "")
                return DBNull.Value;
            if (fontCode == "None" || fontCode == "Unicode")
            {
                //if (p is string)
                return p.ToString().Trim().TrimStart();
                //return p;
            }
            else
            {
                if (p is String)
                {
                    if (fontCode == "TVCN3")
                        return VNConvertor.ConvertTCVN3ToUnicode(p.ToString().Trim().TrimStart());
                    else if (fontCode == "VNI")
                        return VNConvertor.ConvertVNI2Unicode(p.ToString().Trim().TrimStart());
                    //else 
                    //    return 
                }

            }
            return p;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            _flagimport = true;
            //button2.Enabled = false;
            if (ddlImport.Text == "")
            {
                MessageBox.Show("Please choose a Import Code");
                return;
            }
            string filename = textBox1.Text;
            DataTable dtX = IMPORT_SCHEMAControl.GetDataTableStruct(_importCtr.DtStruct, _importCtr.Lookup);
            try
            {
                DataTable dt = PreviewExcelFile(dtX, filename, cboConvertor.SelectedValue.ToString(), ddlImport.SelectedValue.ToString());
                dgvList.DataSource = dt;
                dgvList.AllowEdit = InheritableBoolean.False;
                lbErr.Text = "You have " + dt.Rows.Count + " records from file";
                btnImport.Enabled = false;
                btnGroup.Enabled = dt.Rows.Count > 0;
            }
            catch (Exception ex) { lbErr.Text = ex.Message; }
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
            _importCtr.StrConn = Form_QD._strConnectDes;
        }

        private void AddValidatedList(string db, string xml)
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
                        //if (row["IsNull"] == DBNull.Value || row["IsNull"].ToString() == "False")
                        validate.IsNull = false;
                        validate.Key = row["Key"].ToString();
                        _importCtr.ListV.Add(validate);
                    }

                }
                if (row["PrimaryKey"].ToString() == "True")
                    _importCtr.LKey.Add(row["Key"].ToString());
            }
        }
        private void dgvList_LoadingRow(object sender, Janus.Windows.GridEX.RowLoadEventArgs e)
        {
            if (e.Row.RowType == RowType.Record)
                ValidatedRow(_importCtr, e.Row);
        }
        bool _flagimport = true;
        private void ValidatedRow(IMPORT_SCHEMAControl _importCtr, GridEXRow gridEXRow)
        {
            bool flag = true;
            foreach (GridEXCell cell in gridEXRow.Cells)
            {
                string message = "";

                if (!_importCtr.ContrainList(cell.Column.Key, cell.Value, ref message))
                {
                    cell.ImageIndex = 0;
                    cell.Column.ColumnType = ColumnType.ImageAndText;
                    cell.ToolTipText = message;
                    gridEXRow.RowStyle = dgvList.RowWithErrorsFormatStyle;
                    flag = false;
                }
            }
            _flagimport = _flagimport & flag;
        }

        private void multiColumnCombo1_ValueChanged(object sender, EventArgs e)
        {
            DataRow row = ((DataRowView)ddlImport.SelectedItem).Row;
            DTO.IMPORT_SCHEMAInfo importInf = new DTO.IMPORT_SCHEMAInfo(row);

            string key = importInf.DEFAULT_CONN;
            _importCtr.StrConn = Form_QD.Config.GetConnection(ref key, "AP");

            //StringReader sb = new StringReader(row["FIELD_TEXT"].ToString());
            byte[] byteArray = Encoding.ASCII.GetBytes(row["FIELD_TEXT"].ToString());
            MemoryStream stream = new MemoryStream(byteArray);

            //MemoryStream stream = new MemoryStream(
            ////FileStream file = new FileStream(Application.StartupPath + "\\pbs.BO.HR.TSHInfoList", FileMode.Open);
            dgvList.LoadLayoutFile(stream);
            foreach (GridEXColumn col in dgvList.RootTable.Columns)
                col.DataMember = col.Key;
            //sstreamb.Close();
            stream.Close();
            _importCtr.Lookup = row["LOOK_UP"].ToString();
            AddValidatedList(row["DB"].ToString(), row["FIELD_TEXT"].ToString());
            dgvList.TotalRow = InheritableBoolean.True;
            dgvList.GroupTotals = GroupTotals.Always;
            dgvList.DataSource = new DataTable();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int dem = 0;
            //if (_flagimport)
            //{
            //    foreach (GridEXRow row in dgvList.GetRows())
            //    {
            //        if (row.RowType == RowType.Record && row.RowStyle != dgvList.RowWithErrorsFormatStyle)
            //        {
            //            string sErr = "";
            //            int result = _importCtr.Import(((DataRowView)row.DataRow).Row, checkBox1.Checked, checkBox2.Checked, ref sErr);
            //            if (result == 1)
            //                dem++;
            //        }

            //    }
            //    MessageBox.Show("Have " + dem + " update records");
            //}
            //else
            //{
            if (MessageBox.Show("Do you want to import these correct records?", "Warning", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                DataTable dt = dgvList.DataSource as DataTable;
                for (int i = dgvList.RowCount - 1; i >= 0; i--)
                {
                    GridEXRow row = dgvList.GetRow(i);
                    if (row.RowType == RowType.Record && row.RowStyle != dgvList.RowWithErrorsFormatStyle)
                    {
                        DataRow dtrow = ((DataRowView)row.DataRow).Row;
                        string sErr = "";
                        int result = _importCtr.Import(dtrow, checkBox1.Checked, checkBox2.Checked, ref sErr);

                        if (result == 1)
                        {
                            dt.Rows.Remove(dtrow);
                            dem++;
                        }
                    }

                }

                btnImport.Enabled = false;
                btnGroup.Enabled = false;
                lbErr.Text = "Have " + dem + " update records";
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
            if (dgvList.RowCount > 0)
            {
                DataSet dset = null;
                DataTable dt = dgvList.DataSource as DataTable;

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
                    dgvList.DataSource = dtgroup;
                    lbErr.Text = "You have " + dtgroup.Rows.Count + " records by Group";
                    btnImport.Enabled = dtgroup.Rows.Count > 0;
                }
                else btnImport.Enabled = dt.Rows.Count > 0;

            }
        }
    }

}
