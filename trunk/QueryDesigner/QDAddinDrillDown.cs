using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;

using QueryBuilder;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;


using System.IO;
using BUS;
using Janus.Windows.GridEX;

namespace QueryDesigner
{
    public partial class QDAddinDrillDown : Form
    {
        private QueryBuilder.SQLBuilder _sqlBuilder;
        bool flag_view = false;
        String sErr = "";
        string idHandler = "";
        String _ttFormular = "";
        Excel._Application _xlsApp;
        String _status = "F";
        string THEME = "Breeze";
        string nodeCode = "";
        public DataTable _dataTable = new DataTable();
        Point downPt;
        Node[] _arrNodes = null;
        //string _strConnectDes = "";
        QDConfig _config;

        public QDConfig Config
        {
            get { return _config; }
            set
            {
                _config = value;
                BUS.LIST_QD_SCHEMAControl schCtr = new LIST_QD_SCHEMAControl();

                DTO.LIST_QD_SCHEMAInfo schInf = schCtr.Get(_sqlBuilder.Database, _sqlBuilder.Table, ref sErr);
                string keyconn = schInf.DEFAULT_CONN;
                string connectstring = _config.GetConnection(ref keyconn, "AP");
                _sqlBuilder.ConnID = keyconn;
                _sqlBuilder.StrConnectDes = connectstring;
                //_strConnectDes = connectstring;
            }
        }
        public String Status
        {

            get { return _status; }
            set { _status = value; }
        }
        public Excel._Application XlsApp
        {
            get { return _xlsApp; }
            set { _xlsApp = value; }
        }
        public String TTFormular
        {
            get { return _ttFormular; }
            set { _ttFormular = value; }
        }
        public string Pos
        {
            get { return _sqlBuilder.Pos; }
            set { _sqlBuilder.Pos = value; }
        }
        public QDAddinDrillDown(string connectDesc)
        {
            InitializeComponent();
            //_strConnectDes = connectDesc;
            _sqlBuilder = new QueryBuilder.SQLBuilder(processingMode.Balance);
            _sqlBuilder.StrConnectDes = connectDesc;
            //TopMost = true;
            //ThemeResolutionService.ApplyThemeToControlTree(this, THEME);
        }
        public QDAddinDrillDown(string Pos, Excel._Application xls, string formular, string connectDesc)
        {
            InitializeComponent();
            //_strConnectDes = connectDesc;
            Init(Pos, xls, formular);
            _sqlBuilder.StrConnectDes = connectDesc;
            //ThemeResolutionService.ApplyThemeToControlTree(this, THEME);

        }
        public void Init(string Pos, Excel._Application xls, string formular)
        {
            _sqlBuilder = new QueryBuilder.SQLBuilder(processingMode.Balance);
            //((GridViewComboBoxColumn)dgvSelectNodes.Columns["Agregate"]).DataSource = Parsing.GetListNumberAgregate();
            _sqlBuilder.Pos = Pos;
            _xlsApp = xls;
            //TopMost = true;
            GetQueryBuilderFromFomular(formular);
            if (xls == null)
                btPivotTable.Visible = false;
        }
        public QDAddinDrillDown(string Pos, Excel._Application xls, QueryBuilder.SQLBuilder sqBuilder, string connectDesc)
        {
            InitializeComponent();
            Init(Pos, xls, sqBuilder);
            //_strConnectDes = sqBuilder.StrConnectDes;
            //ThemeResolutionService.ApplyThemeToControlTree(this, THEME);

        }
        public void Init(string Pos, Excel._Application xls, QueryBuilder.SQLBuilder sqBuilder)
        {
            _sqlBuilder = new SQLBuilder(sqBuilder);
            //((GridViewComboBoxColumn)dgvSelectNodes.Columns["Agregate"]).DataSource = Parsing.GetListNumberAgregate();
            _sqlBuilder.Pos = Pos;
            _xlsApp = xls;
            //TopMost = true;
            //GetQueryBuilderFromFomular(formular);
        }

        public void GetQueryBuilderFromFomular(string formular)
        {
            if (formular.Contains("TT_XLB_EB") || formular.Contains("USER TABLE"))
            {
                if (_xlsApp != null)
                {
                    Excel._Worksheet sheet = (Excel._Worksheet)_xlsApp.ActiveWorkbook.ActiveSheet;
                    Excel._Workbook wbook = (Excel._Workbook)_xlsApp.ActiveWorkbook;

                    string vParamsString = Regex.Match(formular, @"\" +  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ System.Convert.ToChar(34) + @"\,.+?\)").Value.ToString();

                    // fill to parameter Array
                    int i = 0, n = 0;

                    if (!(string.IsNullOrEmpty(vParamsString)))
                    {
                        vParamsString =  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vParamsString.Substring(1);
                        vParamsString = vParamsString.Substring(1, vParamsString.Length - 2);// Strings.Mid(vParamsString, 1,  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vParamsString.Length - 1); 
                        vParamsString = vParamsString + ","; //  them dau , cho de xu ly
                        n = Regex.Matches(vParamsString, ".*?,").Count; //  cac tham so
                        if (n > 0)
                        {
                            string[] vParameter = new string[n]; // tham so dau tien la vi tri cua cong thuc
                            foreach (System.Text.RegularExpressions.Match p in Regex.Matches(vParamsString, ".*?,"))
                            {
                                i = i + 1;
                                if (i == 1)
                                {
                                    _sqlBuilder.Pos = p.Value.ToString().Replace(",", string.Empty);

                                }
                                else
                                {

                                    string address = p.Value.ToString().Replace(",", string.Empty);
                                    string value = "";
                                    //try
                                    //{
                                    //    value = sheet.get_Range(address, Type.Missing).get_Value(Type.Missing).ToString();

                                    //    _sqlBuilder.ParaValueList[i - 1] = value;
                                    foreach (Excel._Worksheet isheet in wbook.Sheets)
                                    {
                                        try
                                        {
                                            value = isheet.get_Range(address, Type.Missing).get_Value(Type.Missing).ToString();

                                            _sqlBuilder.ParaValueList[i - 1] = value;
                                            break;

                                        }
                                        catch
                                        {
                                        }
                                    }

                                    //}
                                    //catch
                                    //{
                                    //    value = ((Excel.Range)wbook.Sheets.get_Item(address)).get_Value(Type.Missing).ToString();

                                    //    _sqlBuilder.ParaValueList[i - 1] = value;
                                    //}
                                    //vParameter[i - 1] = p.Value.ToString().Replace(",", string.Empty);
                                }

                            }
                        }
                    }
                }
                Parsing.Formular2SQLBuilder(formular, ref _sqlBuilder);



                //SetDataToForm();
            }
        }
        public void SetValueFocus(string address, string value)
        {
            if (idHandler == "DB")
            {
                //txtDB.Text = address;
                //lbDBDesc.Text = value;
                //_sqlBuilder.Database = address;
                //_sqlBuilder.DatabaseV = value;
                //_sqlBuilder.DatabaseP = "{P}";
            }
            else if (idHandler == "LD")
            {
                //txtLedger.Text = address;
                //lbLedgerDesc.Text = value;
                //_sqlBuilder.Ledger = address;
                //_sqlBuilder.LedgerV = value;
                //_sqlBuilder.LedgerP = "{P}";
            }
            else if (idHandler == "VF")
            {
                //txtFilterFrom.Text = address;
                //lbValueFrom.Text = value;
                //if (dgvFilter.CurrentRow != null)
                //{
                //    QueryBuilder.Filter filter = dgvFilter.CurrentRow.DataBoundItem as QueryBuilder.Filter;
                //    filter.FilterFrom = address;
                //    filter.ValueFrom = value;
                //    filter.FilterFromP = "{P}";
                //}
            }
            else if (idHandler == "VT")
            {
                //txtFilterTo.Text = address;
                //lbValueTo.Text = value;
                //if (dgvFilter.CurrentRow != null)
                //{
                //    QueryBuilder.Filter filter = dgvFilter.CurrentRow.DataBoundItem as QueryBuilder.Filter;
                //    filter.FilterTo = address;
                //    filter.ValueTo = value;
                //    filter.FilterToP = "{P}";
                //}
            }

        }
        private QueryBuilder.SQLBuilder GetDataFromForm()
        {
            return _sqlBuilder;
        }
        private void SetDataToForm()
        {
            //txtDB.Text = _sqlBuilder.Database;
            ////lbDBDesc.Text = _sqlBuilder.Database;
            //txtLedger.Text = _sqlBuilder.Ledger;
            ////lbLedgerDesc.Text = _sqlBuilder.Ledger;
            //txtTable.Text = _sqlBuilder.Table;
            //lbTableDesc_TextChanged(null, null);
            //dgvFilter.DataSource = _sqlBuilder.Filters;
            //dgvSelectNodes.DataSource = _sqlBuilder.SelectedNodes;
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void QDAddin_Load(object sender, EventArgs e)
        {
            try
            {
                dgvResult.AllowDrop = true;

                DialogResult = DialogResult.Yes;
                BindingList<Node> list = SchemaDefinition.GetDecorateTableByCode(_sqlBuilder.Table, _sqlBuilder.Database);
                //twSchema = RadTreeViewLoader.LoadTree(ref twSchema, list, _sqlBuilder.Table, "");
                twSchema1 = TreeViewLoader.LoadTree(ref twSchema1, list, _sqlBuilder.Table, "");
                //BUS.LIST_QD_SCHEMAControl ctr = new LIST_QD_SCHEMAControl();
                //DTO.LIST_QD_SCHEMAInfo inf = ctr.Get(_sqlBuilder.Database, _sqlBuilder.Table, ref sErr);
                //string key = inf.DEFAULT_CONN;
                //_strConnectDes = _sqlBuilder.StrConnectDes = _config.GetConnection(ref key, "AP");
                LoadDataGrid();
            }
            catch (Exception ex) { lbErr.Text = ex.Message; }
            //if (_xlsApp == null)
            //    //TopMost = false;
            //dgvResult.ad
        }

        private void LoadDataGrid()
        {
            BUS.CommonControl commo = new BUS.CommonControl();
            //_sqlBuilder.StrConnectDes = _strConnectDes;
            DataTable dt = _sqlBuilder.BuildDataTable("");
            dgvResult.DataSource = dt;
            dgvResult.RetrieveStructure();
            dgvResult.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.False;
            dgvResult.AutoSizeColumns();
            for (int j = 0; j < _sqlBuilder.SelectedNodes.Count; j++)
            {
                //if (dgvResult.RootTable.Columns.Contains(_sqlBuilder.SelectedNodes[j].MyCode))
                //    dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].Caption = _sqlBuilder.SelectedNodes[j].Description;

                if (_sqlBuilder.SelectedNodes[j].Agregate != "")
                {
                    if (dgvResult.RootTable.Columns.Contains(_sqlBuilder.SelectedNodes[j].MyCode))
                    {
                        switch (_sqlBuilder.SelectedNodes[j].Agregate)
                        {
                            case "SUM":
                                dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Sum;
                                break;
                            case "COUNT":
                                dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Count;
                                break;
                            case "AVG":
                                dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Average;
                                break;
                            case "MAX":
                                dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Max;
                                break;
                            case "MIN":
                                dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Min;
                                break;
                        }

                    }
                    else
                    {
                        if (dgvResult.RootTable.Columns.Contains(_sqlBuilder.SelectedNodes[j].Description))
                        {
                            switch (_sqlBuilder.SelectedNodes[j].Agregate)
                            {
                                case "SUM":
                                    dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].Description].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Sum;
                                    break;
                                case "COUNT":
                                    dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].Description].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Count;
                                    break;
                                case "AVG":
                                    dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].Description].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Average;
                                    break;
                                case "MAX":
                                    dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].Description].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Max;
                                    break;
                                case "MIN":
                                    dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].Description].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Min;
                                    break;
                            }

                        }
                    }
                }
            }
            for (int i = 0; i < dgvResult.RootTable.Columns.Count; i++)
            {
                //dgvResult.RootTable.Columns[i].AutoSizeMode = Janus.Windows.GridEX.ColumnAutoSizeMode.DisplayedCellsAndHeader;
                if (dgvResult.RootTable.Columns[i].DataTypeCode == TypeCode.Decimal || dgvResult.RootTable.Columns[i].DataTypeCode == TypeCode.Double)
                {
                    dgvResult.RootTable.Columns[i].FormatString = dgvResult.RootTable.Columns[i].TotalFormatString = "###,###.##";
                    dgvResult.RootTable.Columns[i].TotalFormatString = "###,###.##";
                }
                // dgvResult.Columns[i].A
            }

            dgvResult.RootTable.GroupTotals = GroupTotals.Always;
            dgvResult.RootTable.TotalRow = InheritableBoolean.True;

        }



        private void btnPrint_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel File(*.xls)|*.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string fileName = sfd.FileName;

                    FileStream sr = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write);
                    gridEXExporter1.Export(sr);
                    sr.Close();
                    if (MessageBox.Show("Do you want to open this document?", "Message", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        System.Diagnostics.Process.Start(sfd.FileName);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void btPivotTable_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            BUS.CommonControl commo = new BUS.CommonControl();
            //_sqlBuilder.StrConnectDes = _strConnectDes;
            _dataTable = _sqlBuilder.BuildDataTable("");
            //_dataTable = _sqlBuilder.BuildDataTable("");
            Close();
        }


        private void btnExpandAll_Click(object sender, EventArgs e)
        {
            dgvResult.ExpandGroups();
        }
        private void btnCollapseAll_Click(object sender, EventArgs e)
        {
            dgvResult.CollapseGroups();
        }

        private void QDAddinDrillDown_MouseUp(object sender, MouseEventArgs e)
        {
            nodeCode = "";
        }
        /*
        private void twSchema_ItemDrag(object sender, RadTreeViewEventArgs e)
        {
            RadSelectedNodesCollection dragNodes = twSchema.SelectedNodes;
            _arrNodes = new Node[dragNodes.Count];
            int i = 0;
            foreach (RadTreeNode dragNode in dragNodes)
            {
                if (dragNode.Nodes.Count == 0)
                {
                    string[] arrNode = dragNode.Tag.ToString().Split(';');
                    _arrNodes[i] = new Node(arrNode[0], dragNode.Name, dragNode.Text, arrNode[1], arrNode[2]);
                    i++;
                }
            }
            if (_arrNodes[0] == null)
                _arrNodes = null;
        }
        private void twSchema_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            RadTreeNode tmpNode = twSchema.SelectedNode;

            if (tmpNode != null && dgvResult.AllowDrop == true && tmpNode.Nodes.Count == 0)
            {
                bool flag = true;
                string[] arrNode = tmpNode.Tag.ToString().Split(';');
                Node a = new Node(arrNode[0], tmpNode.Name, tmpNode.Text, arrNode[1], arrNode[2]);
                for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
                    if (_sqlBuilder.SelectedNodes[i].Code == a.Code)
                    {
                        flag = false;
                        break;
                    }
                if (flag)
                    _sqlBuilder.SelectedNodes.Add(a);

                DataTable dt = Parsing.GetListNumberAgregate();
                Node aa = _sqlBuilder.SelectedNodes[_sqlBuilder.SelectedNodes.Count - 1];
                if (aa.FType == "" || aa.FType[0] != 'N')
                    dt = Parsing.GetListStringAgregate();

                LoadDataGrid();
            }
        }
        private void twSchema_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void twSchema_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(string)))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.All;

            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
            //e.Effect = DragDropEffects.Copy | DragDropEffects.All;
        }

        private void twSchema_DragDrop(object sender, DragEventArgs e)
        {
            if (twSchema.AllowDragDrop == true)
            {
                string code = (string)e.Data.GetData(typeof(string));
                if (code != null)
                {
                    int vitri = -1;
                    for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
                        if (_sqlBuilder.SelectedNodes[i].MyCode == code)
                        {
                            vitri = i;
                        }
                    if (vitri != -1)
                    {
                        dgvResult.RootTable.Columns.Remove(_sqlBuilder.SelectedNodes[vitri].MyCode);
                        _sqlBuilder.SelectedNodes.RemoveAt(vitri);
                    }


                    LoadDataGrid();
                }
            }
        }



        private void twSchema_MouseUp(object sender, MouseEventArgs e)
        {
            if (_arrNodes != null)
            {
                //TopMost = true;
                Rectangle rect = this.dgvResult.RectangleToScreen(this.dgvResult.ClientRectangle);
                if (rect.Contains(Cursor.Position))
                {

                    bool flag = true;
                    Node[] arrNode = _arrNodes;
                    for (int j = 0; j < arrNode.Length; j++)
                    {
                        Node a = arrNode[j];
                        for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
                            if (_sqlBuilder.SelectedNodes[i].Code == a.Code)
                            {
                                flag = false;
                                break;
                            }
                        if (flag)
                        {
                            _sqlBuilder.SelectedNodes.Add(a);
                            LoadDataGrid();
                        }
                    }

                }
                _arrNodes = null;
            }
        }

        private void twSchema_MouseMove(object sender, MouseEventArgs e)
        {
            if (_arrNodes != null)
            {
                //TopMost = false;
                dgvResult.Cursor = Cursors.Hand;
            }
        }
        */
        private void dgvResult_MouseMove(object sender, MouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Left)
            //    twSchema.Cursor = Cursors.Arrow;
        }

        private void dgvResult_MouseLeave(object sender, EventArgs e)
        {
            //if (nodeCode != "")
            //    DoDragDrop(nodeCode, DragDropEffects.Copy);
            //nodeCode = "";
        }
        private void dgvResult_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                //RadElement element = this.dgvResult.ElementTree.GetElementAtPoint(e.Location);
                ////GridDataCellElement cell = element as GridDataCellElement;
                //if (element is GridHeaderCellElement)
                //{
                //    GridHeaderCellElement cell = element as GridHeaderCellElement;
                //    if (cell.ColumnInfo is GridViewDataColumn)
                //        nodeCode = ((GridViewDataColumn)cell.ColumnInfo).FieldName;
                //    //    DoDragDrop(((GridViewDataColumn)cell.ColumnInfo).FieldName, DragDropEffects.All);
                //}

            }
        }

        private void dgvResult_DraggingColumn(object sender, Janus.Windows.GridEX.ColumnActionCancelEventArgs e)
        {
            nodeCode = e.Column.Key;
        }

        private void dgvResult_MouseUp(object sender, MouseEventArgs e)
        {
            if (nodeCode != "")
            {
                Rectangle rect = this.dgvResult.RectangleToScreen(this.dgvResult.ClientRectangle);
                if (!rect.Contains(Cursor.Position))
                {

                    int vitri = -1;
                    for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
                        if (_sqlBuilder.SelectedNodes[i].MyCode == nodeCode || _sqlBuilder.SelectedNodes[i].Description == nodeCode)
                        {
                            vitri = i;
                        }
                    if (vitri != -1)
                    {
                        dgvResult.RootTable.Columns.Remove(_sqlBuilder.SelectedNodes[vitri].MyCode);
                        _sqlBuilder.SelectedNodes.RemoveAt(vitri);
                    }


                    LoadDataGrid();


                }
            }
            nodeCode = "";
        }
        private void dgvResult_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(Node)))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void dgvResult_DragDrop(object sender, DragEventArgs e)
        {
            bool flag = true;
            Node a = (Node)e.Data.GetData(typeof(Node));
            for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
                if (_sqlBuilder.SelectedNodes[i].Code == a.Code)
                {
                    flag = false;
                    break;
                }
            if (flag)
            {
                _sqlBuilder.SelectedNodes.Add(a);
                LoadDataGrid();
            }
        }

        private void twSchema1_ItemDrag(object sender, ItemDragEventArgs e)
        {
            if (e.Item != null)
            {
                Node _node = (Node)((TreeNode)e.Item).Tag;
                if (_node.FType != "S")
                    twSchema1.DoDragDrop(_node, DragDropEffects.Copy);
            }
        }

        private void twSchema1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TreeNode tmpNode = twSchema1.SelectedNode;

                if (tmpNode != null)
                {
                    bool flag = true;
                    string[] arrNode = tmpNode.Tag.ToString().Split(';');
                    Node a = ((Node)tmpNode.Tag).CloneNode();
                    for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
                        if (_sqlBuilder.SelectedNodes[i].Code == a.Code)
                        {
                            flag = false;
                            break;
                        }
                    if (flag)
                    {
                        _sqlBuilder.SelectedNodes.Add(a);
                        //if (dgvPreview.RootTable.Columns.IndexOf(a.MyCode) < 0)
                        //    dgvPreview.RootTable.Columns.Add(a.MyCode);
                    }

                    //DataTable dt = Parsing.GetListNumberAgregate();
                    //Node aa = _sqlBuilder.SelectedNodes[_sqlBuilder.SelectedNodes.Count - 1];
                    //if (aa.FType == "" || aa.FType[0] != 'N')
                    //    dt = Parsing.GetListStringAgregate();

                    //customerColumn.DataSource = dt;
                    //dgvSelectNodes.MasterGridViewTemplate.AutoGenerateColumns = false;
                    //dgvSelectNodes.DataSource = _sqlBuilder.SelectedNodes;
                    //dgvSelectNodes.Refresh();
                }
            }
        }

        private void twSchema1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            TreeNode tmpNode = twSchema1.SelectedNode;

            if (tmpNode != null && tmpNode.Nodes.Count == 0)
            {
                bool flag = true;
                string[] arrNode = tmpNode.Tag.ToString().Split(';');
                Node a = ((Node)tmpNode.Tag).CloneNode();
                for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
                    if (_sqlBuilder.SelectedNodes[i].Code == a.Code)
                    {
                        flag = false;
                        break;
                    }
                if (flag)
                {
                    _sqlBuilder.SelectedNodes.Add(a);
                    LoadDataGrid();
                    //if (dgvPreview.RootTable.Columns.IndexOf(a.MyCode) < 0)
                    //    dgvPreview.RootTable.Columns.Add(a.MyCode);
                }

                //DataTable dt = Parsing.GetListNumberAgregate();
                //Node aa = _sqlBuilder.SelectedNodes[_sqlBuilder.SelectedNodes.Count - 1];
                //if (aa.FType == "" || aa.FType[0] != 'N')
                //    dt = Parsing.GetListStringAgregate();

                //customerColumn.DataSource = dt;
                //dgvSelectNodes.MasterGridViewTemplate.AutoGenerateColumns = false;
                //dgvSelectNodes.DataSource = _sqlBuilder.SelectedNodes;
                //dgvSelectNodes.Refresh();
            }
        }

        private void lbErr_Click(object sender, EventArgs e)
        {
            MessageBox.Show(lbErr.Text);
        }





    }
}
