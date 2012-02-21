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
using System.Management;
using System.IO;
using BUS;
namespace QueryDesigner
{
    public partial class QDAddIn : Form
    {
        private QueryBuilder.SQLBuilder _sqlBuilder;
        bool flag_view = false;
        Node[] _arrNodes = null;
        String _sErr = "";
        string idHandler = "";
        String _ttFormular = "";
        Excel._Application _xlsApp;
        String _status = "F";
        string THEME = "Breeze";
        string _strConnect = "";
        string _strConnectDes = "";
        DataTable _dataReturn;
        string _currentAddress = "A1";
        QDConfig _config = null;

        public QDConfig Config
        {
            get { return _config; }
            set
            {
                _config = value;
                BUS.LIST_QD_SCHEMAControl schCtr = new LIST_QD_SCHEMAControl();

                DTO.LIST_QD_SCHEMAInfo schInf = schCtr.Get(_sqlBuilder.Database, _sqlBuilder.Table, ref _sErr);
                string keyconn = schInf.DEFAULT_CONN;
                string connectstring = _config.GetConnection(ref keyconn, "AP");
                _sqlBuilder.ConnID = keyconn;
                _sqlBuilder.StrConnectDes = connectstring;
            }
        }
        public string CurrentAddress
        {
            get { return _currentAddress; }
            set
            {
                _currentAddress = value;
            }
        }
        public DataTable DataReturn
        {
            get { return _dataReturn; }
            set { _dataReturn = value; }
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
        public QDAddIn(string connect, string connectDes)
        {
            InitializeComponent();
            ////ThemeResolutionService.ApplyThemeToControlTree(this, THEME);
            _sqlBuilder = new QueryBuilder.SQLBuilder(processingMode.Balance);
            ((DataGridViewComboBoxColumn)dgvSelectNodes.Columns["colAgregate"]).DataSource = Parsing.GetListNumberAgregate();
            ((DataGridViewComboBoxColumn)dgvSelectNodes.Columns["colAgregate"]).DisplayMember = "Description";
            ((DataGridViewComboBoxColumn)dgvSelectNodes.Columns["colAgregate"]).ValueMember = "Code";
            //TopMost = true;
            _strConnect = connect;
            _strConnectDes = connectDes;
        }
        public QDAddIn(string Pos, Excel._Application xls, string formular, string connect, string connectDes)
        {
            InitializeComponent();
            ////ThemeResolutionService.ApplyThemeToControlTree(this, THEME);

            _strConnect = connect;
            _strConnectDes = connectDes;
            Init(Pos, xls, formular);
        }

        public void Init(string Pos, Excel._Application xls, string formular)
        {
            _sqlBuilder = new QueryBuilder.SQLBuilder(processingMode.Balance);
            ((DataGridViewComboBoxColumn)dgvSelectNodes.Columns["colAgregate"]).DataSource = Parsing.GetListNumberAgregate();
            ((DataGridViewComboBoxColumn)dgvSelectNodes.Columns["colAgregate"]).DisplayMember = "Description";
            ((DataGridViewComboBoxColumn)dgvSelectNodes.Columns["colAgregate"]).ValueMember = "Code";
            _sqlBuilder.Pos = Pos;
            _xlsApp = xls;
            //TopMost = true;
            GetQueryBuilderFromFomular(formular);
        }

        public void GetQueryBuilderFromFomular(string formular)
        {
            if (formular.Contains("TT_XLB_EB") || formular.Contains("USER TABLE"))
            {
                Excel._Worksheet sheet = (Excel._Worksheet)_xlsApp.ActiveWorkbook.ActiveSheet;
                string vParamsString = Regex.Match(formular, @"\" +  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ System.Convert.ToChar(34) + @"\,.+?\)").Value.ToString();
                Excel._Workbook wbook = (Excel._Workbook)_xlsApp.ActiveWorkbook;
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
                                //vParameter[i - 1] = p.Value.ToString().Replace(",", string.Empty);
                            }

                        }
                    }
                }
                Parsing.Formular2SQLBuilder(formular, ref _sqlBuilder);



                SetDataToForm();
            }
        }
        public void SetValueFocus(string address, string value)
        {
            _currentAddress = address;
            if (idHandler == "DB")
            {
                txtDB.Text = address;
                lbDBDesc.Text = value;
                UpdateDatabase(true);
                //_sqlBuilder.Database = address;
                //_sqlBuilder.DatabaseV = value;
                //_sqlBuilder.DatabaseP = "{P}";
            }
            else if (idHandler == "LD")
            {
                txtLedger.Text = address;
                lbLedgerDesc.Text = value;
                UpdateLedger(true);
                //_sqlBuilder.Ledger = address;
                //_sqlBuilder.LedgerV = value;
                //_sqlBuilder.LedgerP = "{P}";
            }
            else if (idHandler == "VF")
            {
                txtFilterFrom.Text = address;
                lbValueFrom.Text = value;
                UpdateFilterFrom(true);
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
                txtFilterTo.Text = address;
                lbValueTo.Text = value;
                UpdateFilterTo(true);
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

            txtDB.Text = _sqlBuilder.Database;
            //lbDBDesc.Text = _sqlBuilder.Database;
            txtLedger.Text = _sqlBuilder.Ledger;
            //lbLedgerDesc.Text = _sqlBuilder.Ledger;
            txtTable.Text = _sqlBuilder.Table;

            BUS.LIST_QD_SCHEMAControl schctr = new LIST_QD_SCHEMAControl();
            DTO.LIST_QD_SCHEMAInfo schInf = schctr.Get(_sqlBuilder.Database, _sqlBuilder.Table, ref _sErr);
            _sqlBuilder.ConnID = schInf.DEFAULT_CONN;
            GetDescSource();

            lbTableDesc_TextChanged(null, null);
            //dgvFilter.DataSource = _sqlBuilder.Filters;
            //dgvSelectNodes.DataSource = _sqlBuilder.SelectedNodes;
        }



        /*
        private void twSchema_ItemDrag(object sender, RadTreeViewEventArgs e)
        {
            TreeNode dragNodes = twSchema1.SelectedNodes;
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
        private void twSchema_MouseUp(object sender, MouseEventArgs e)
        {
            //TopMost = true;
            if (_arrNodes != null)
            {
                Rectangle rect = this.dgvSelectNodes.RectangleToScreen(this.dgvSelectNodes.ClientRectangle);
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
                            _sqlBuilder.SelectedNodes.Add(a);
                    }
                }


                rect = this.dgvFilter.RectangleToScreen(this.dgvFilter.ClientRectangle);
                if (rect.Contains(Cursor.Position))
                {
                    bool flag = true;
                    Node[] arrNode = _arrNodes;
                    for (int j = 0; j < arrNode.Length; j++)
                    {
                        Node a = arrNode[j];
                        QueryBuilder.Filter tmp = new QueryBuilder.Filter(a);

                        _sqlBuilder.Filters.Add(tmp);
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
                dgvSelectNodes.Cursor = Cursors.Hand;
                dgvFilter.Cursor = Cursors.Hand;
            }
        }
        private void twSchema_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            RadTreeNode tmpNode = twSchema.SelectedNode;

            if (tmpNode != null && dgvSelectNodes.AllowDrop == true)
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

                //DataTable dt = Parsing.GetListNumberAgregate();
                //Node aa = _sqlBuilder.SelectedNodes[_sqlBuilder.SelectedNodes.Count - 1];
                //if (aa.FType == "" || aa.FType[0] != 'N')
                //    dt = Parsing.GetListStringAgregate();

                //customerColumn.DataSource = dt;
                ////dgvSelectNodes.MasterGridViewTemplate.AutoGenerateColumns = false;
                ////dgvSelectNodes.DataSource = _sqlBuilder.SelectedNodes;
            }
        }
        */

        private void btOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            _ttFormular = _sqlBuilder.BuildTTformula(_sqlBuilder.Pos);
            Close();
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btDB_Click(object sender, EventArgs e)
        {
            QueryDesigner.Form_DTBView a = new QueryDesigner.Form_DTBView();
            //a.themname = this.ThemeName;
            a.BringToFront();
            a.ShowDialog(this);
            txtDB.Text = a.Code_DTB;
            lbDBDesc.Text = a.Description_DTB;
        }


        private void btTable_Click(object sender, EventArgs e)
        {
            QueryDesigner.Form_TableView a = new QueryDesigner.Form_TableView();
            //a.themname = this.ThemeName;
            a.Code_DTB = txtDB.Text;
            a.BringToFront();
            if (a.ShowDialog(this) == DialogResult.OK)
            {
                txtTable.Text = a.Code_DTB;
                lbTableDesc.Text = a.Description_DTB;
            }
        }

        public void lbTableDesc_TextChanged(object sender, EventArgs e)
        {
            if (lbTableDesc.Text != "_" && txtTable.Text.Trim() != twSchema1.Name)
            {
                _sqlBuilder.Table = txtTable.Text.Trim();
                BindingList<Node> list = SchemaDefinition.GetDecorateTableByCode(_sqlBuilder.Table, _sqlBuilder.Database);
                //twSchema = RadTreeViewLoader.LoadTree(ref twSchema, list, txtTable.Text.Trim(), "");
                twSchema1 = TreeViewLoader.LoadTree(ref twSchema1, list, _sqlBuilder.Table, "");
                //_sqlBuilder.SelectedNodes.Clear();
                //_sqlBuilder.Filters.Clear();
            }
        }



        private void txtTable_Leave(object sender, EventArgs e)
        {
            //BindingList<QueryBuilder.TableItem> tmp = QueryBuilder.SchemaDefinition.GetTableList();
            //foreach (QueryBuilder.TableItem x in tmp)
            //{
            //    if (x.Code == txtTable.Text.Trim())
            //    {
            //        lbTableDesc.Text = x.Description;
            //        return;
            //    }
            //}
            //lbTableDesc.Text = "_";
        }

        private void txtDB_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                btDB_Click(null, null);
        }

        private void txtTable_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                btTable_Click(null, null);
        }



        private void txtDB_Leave(object sender, EventArgs e)
        {
            //DBInfoList tmp = DBInfoList.GetDBInfoList();
            //foreach (QueryBuilder.DBInfo x in tmp)
            //{
            //    if (x.Code == txtDB.Text.Trim())
            //    {
            //        lbDBDesc.Text = x.Description;
            //        return;
            //    }
            //}
            //lbDBDesc.Text = "_";
        }
        private void QDAddin_Load(object sender, EventArgs e)
        {
            dgvFilter.AutoGenerateColumns = false;
            dgvSelectNodes.AutoGenerateColumns = false;
            SQLBuilderBindingSource.DataSource = _sqlBuilder;
            DialogResult = DialogResult.Yes;
            btnUserTable.Visible = ValidateLicense(System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\TVC-QD\\Configuration\\license.bin");
            _status = "I";

            Text = "TTFomular - " + _sqlBuilder.Pos;
        }


        private void txtDB_Enter(object sender, EventArgs e)
        {
            idHandler = "DB";
        }

        private void txtLedger_Enter(object sender, EventArgs e)
        {
            idHandler = "LD";
        }

        private void txtFilterFrom_Enter(object sender, EventArgs e)
        {
            idHandler = "VF";
        }

        private void txtFilterTo_Enter(object sender, EventArgs e)
        {
            idHandler = "VT";
        }

        private void txtTable_Enter(object sender, EventArgs e)
        {
            idHandler = "";
        }



        private void txtFilterFrom_TextChanged(object sender, EventArgs e)
        {
            //Excel._Worksheet sheet = (Excel._Worksheet)_xlsApp.ActiveWorkbook.ActiveSheet;
            //UpdateFilterFrom(false);
        }

       

        private void txtFilterTo_TextChanged(object sender, EventArgs e)
        {
            //Excel._Worksheet sheet = (Excel._Worksheet)_xlsApp.ActiveWorkbook.ActiveSheet;
            //UpdateFilterTo(false);
        }



        private void txtLedger_TextChanged(object sender, EventArgs e)
        {
            //Excel._Worksheet sheet = (Excel._Worksheet)_xlsApp.ActiveWorkbook.ActiveSheet;
            //UpdateLedger();


        }



        private void txtDB_TextChanged(object sender, EventArgs e)
        {
            //Excel._Worksheet sheet = (Excel._Worksheet)_xlsApp.ActiveWorkbook.ActiveSheet;
            //UpdateDatabase();
            if (lbDBDesc.Text == "" || lbDBDesc.Text == "_")
            {
                _sqlBuilder.Database = txtDB.Text;
            }
        }



        private void btTest_Click(object sender, EventArgs e)
        {
            string query = _sqlBuilder.BuildSQLEx("");
            MessageBox.Show(query);
            //Clipboard.SetData("System.String", query);
        }



        private void txtTable_TextChanged(object sender, EventArgs e)
        {

        }

        private void GetDescSource()
        {
            BindingList<QueryBuilder.TableItem> tmp = QueryBuilder.SchemaDefinition.GetTableList(_sqlBuilder.Database);
            foreach (QueryBuilder.TableItem x in tmp)
            {
                if (x.Code == txtTable.Text.Trim())
                {
                    lbTableDesc.Text = x.Description;
                    return;
                }
            }
        }
        /*
         private void dgvFilter_CellValueChanged(object sender, GridViewCellEventArgs e)
        {
            QueryBuilder.Filter x = e.Row.DataBoundItem as QueryBuilder.Filter;
            if (x != null)
            {
                if (dgvFilter.Columns[e.ColumnIndex].FieldName == "FilterFrom")
                    x.FilterFrom = x.ValueFrom = e.Value.ToString();
                else if (dgvFilter.Columns[e.ColumnIndex].FieldName == "FilterTo")
                    x.FilterTo = x.ValueTo = e.Value.ToString();
            }
        }

private void dgvFilter_RowsChanged(object sender, GridViewCollectionChangedEventArgs e)
        {
            //if (dgvFilter.CurrentRow != null)
            //{
            //    QueryBuilder.Filter x = dgvFilter.CurrentRow.DataBoundItem as QueryBuilder.Filter;
            //    txtFilterFrom.Text = x.FilterFrom;
            //    txtFilterTo.Text = x.FilterTo;
            //}
        }
 private void dgvFilter_CurrentRowChanged(object sender, EventArgs e)
        {

        }
        private void dgvFilter_SelectionChanged(object sender, EventArgs e)
        {
            //if (dgvFilter.CurrentRow != null)
            //{
            //    QueryBuilder.Filter x = dgvFilter.CurrentRow.DataBoundItem as QueryBuilder.Filter;
            //    txtFilterFrom.Text = x.FilterFrom;
            //    txtFilterTo.Text = x.FilterTo;
            //}
        }
        private void dgvFilter_CellClick(object sender, GridViewCellEventArgs e)
        {
            //if (dgvFilter.CurrentRow != null)
            //{
            //    QueryBuilder.Filter x = dgvFilter.CurrentRow.DataBoundItem as QueryBuilder.Filter;
            //    txtFilterFrom.Text = x.FilterFrom;
            //    txtFilterTo.Text = x.FilterTo;
            //}
        }

        private void dgvFilter_CurrentRowChanged(object sender, CurrentRowChangedEventArgs e)
        {
            if (dgvFilter.CurrentRow != null)
            {
                QueryBuilder.Filter x = dgvFilter.CurrentRow.DataBoundItem as QueryBuilder.Filter;
                txtFilterFrom.Text = x.FilterFrom;
                txtFilterTo.Text = x.FilterTo;
            }
        }
        */
        private void btnUserTable_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            _ttFormular = _sqlBuilder.BuildTTformula(_sqlBuilder.Pos);
            _ttFormular = _ttFormular.Replace("=TT_XLB_EB", "USER TABLE");
            Close();
        }

        private void btnCommend_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Status = "C";
            _ttFormular = _sqlBuilder.BuildTTformula(_sqlBuilder.Pos);

            Close();
        }



        private void radPanel1_Enter(object sender, EventArgs e)
        {
            if (!(sender is TextBox))
                idHandler = "";
        }

        private void btnList_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Status = "L";

            BUS.CommonControl commo = new BUS.CommonControl();
            _sqlBuilder.StrConnectDes = _strConnectDes;
            _dataReturn = _sqlBuilder.BuildDataTable("");
            Close();
            //FrmListReport frm = new FrmListReport(_dataTable);
            ////TopMost = false;
            //frm.FormClosed += new FormClosedEventHandler(frm_FormClosed);
            //frm.ShowDialog();
        }

        void frm_FormClosed(object sender, FormClosedEventArgs e)
        {
            //TopMost = true;
            if (_status == "I")
                _status = "F";
        }

        private void btnAnalysis_Click(object sender, EventArgs e)
        {
            string key = _sqlBuilder.ConnID;
            _strConnectDes = _sqlBuilder.StrConnectDes = _config.GetConnection(ref key, "AP");
            _sqlBuilder.ConnID = key;
            QDAddinDrillDown frmD = new QDAddinDrillDown(_sqlBuilder.Pos, _xlsApp, _sqlBuilder, _strConnectDes);

            //TopMost = false;
            frmD.FormClosed += new FormClosedEventHandler(frm_FormClosed);
            frmD.Show(this);
        }

        private bool ValidateLicense(string _pathLicense)
        {
            if (File.Exists(_pathLicense))
            {
                StreamReader reader = new StreamReader(_pathLicense);
                string result = reader.ReadLine();
                string kq = RC2.DecryptString(result, Form_QD._key, Form_QD._iv, Form_QD._padMode, Form_QD._opMode);
                string[] tmp = kq.Split(';');
                DTO.License license = new DTO.License();
                license.CompanyName = tmp[0];
                license.ExpiryDate = Convert.ToInt32(tmp[1]);
                license.Modules = tmp[2];
                license.NumUsers = Convert.ToInt32(tmp[3]);
                license.SerialNumber = tmp[4];
                license.Key = tmp[5];
                //license.SerialCPU = tmp[6];
                license.SerialCPU = GetProcessorId();
                reader.Close();


                string param = license.CompanyName + license.SerialNumber + license.NumUsers.ToString() + license.Modules + license.ExpiryDate.ToString() + license.SerialCPU;


                string temp = RC2.EncryptString(param, Form_QD._key, Form_QD._iv, Form_QD._padMode, Form_QD._opMode);
                string key = Convert.ToBase64String(new System.Security.Cryptography.SHA1CryptoServiceProvider().ComputeHash(Encoding.ASCII.GetBytes(temp)));
                if (key == license.Key)
                {
                    int now = Convert.ToInt32(DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00"));
                    BUS.CommonControl ctr = new CommonControl();
                    object dt = ctr.executeScalar("select getdate()", _strConnect);
                    if (dt != null && dt is DateTime)
                    {
                        now = Convert.ToInt32(((DateTime)dt).Year.ToString() + ((DateTime)dt).Month.ToString("00") + ((DateTime)dt).Day.ToString("00"));
                    }
                    if (now > license.ExpiryDate)
                    {

                        _sErr = "Your license is expired!";
                    }
                    else
                    {
                        //if (license.Modules.Length == 4 && license.Modules.Substring(3) == "Y")
                        return true;
                        //else _flagQDADD = false;

                    }
                }
                else
                {

                    _sErr = "Application have not license!";
                }

            }
            else
            {

                _sErr = "Application have not license!";
            }
            return false;
        }
        public static string GetProcessorId()
        {
            string strCPU = "";
            string Key = "Win32_DiskDrive";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("select * from " + Key + " where InterfaceType ='IDE'");
            try
            {
                foreach (ManagementObject share in searcher.Get())
                {
                    if (share.Properties.Count <= 0)
                    {
                        return "";
                    }
                    foreach (PropertyData PC in share.Properties)
                    {
                        if (PC.Name.Contains("SerialNumber") || PC.Name.Contains("SerialNumber"))
                        {
                            strCPU += Convert.ToString(PC.Value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return strCPU;
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
                //for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
                //    if (_sqlBuilder.SelectedNodes[i].Code == a.Code)
                //    {
                //        flag = false;
                //        break;
                //    }
                if (flag)
                {
                    _sqlBuilder.SelectedNodes.Add(a);
                    //LoadDataGrid();
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

        private void dgvSelectNodes_DragDrop(object sender, DragEventArgs e)
        {
            bool flag = true;
            Node a = (Node)e.Data.GetData(typeof(Node));
            //for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
            //    if (_sqlBuilder.SelectedNodes[i].Code == a.Code)
            //    {
            //        flag = false;
            //        break;
            //    }
            if (flag)
            {
                _sqlBuilder.SelectedNodes.Add(a);
                //LoadDataGrid();
            }
        }

        private void dgvSelectNodes_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(Node)))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }
        private void dgvSelectNodes_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void dgvFilter_DragDrop(object sender, DragEventArgs e)
        {
            bool flag = true;
            Node a = (Node)e.Data.GetData(typeof(Node));

            _sqlBuilder.Filters.Add(new QueryBuilder.Filter(a));
            //LoadDataGrid();

        }


        private void dgvFilter_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                ((QueryBuilder.Filter)dgvFilter.Rows[e.RowIndex].DataBoundItem).FilterFrom =
                    ((QueryBuilder.Filter)dgvFilter.Rows[e.RowIndex].DataBoundItem).ValueFrom;
                ((QueryBuilder.Filter)dgvFilter.Rows[e.RowIndex].DataBoundItem).FilterFrom =
                    ((QueryBuilder.Filter)dgvFilter.Rows[e.RowIndex].DataBoundItem).ValueTo;
            }
        }

        private void dgvFilter_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(Node)))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void dgvFilter_CurrentCellChanged(object sender, EventArgs e)
        {
            //if (dgvFilter.CurrentCell != null)
            //{
            //    QueryBuilder.Filter x = dgvFilter.CurrentRow.DataBoundItem as QueryBuilder.Filter;
            //    txtFilterFrom.Text = x.FilterFrom;
            //    txtFilterTo.Text = x.FilterTo;
            //}
        }

        private void txtTable_Validated(object sender, EventArgs e)
        {
            BUS.LIST_QD_SCHEMAControl schctr = new LIST_QD_SCHEMAControl();
            DTO.LIST_QD_SCHEMAInfo schInf = schctr.Get(_sqlBuilder.Database, _sqlBuilder.Table, ref _sErr);
            _sqlBuilder.ConnID = schInf.DEFAULT_CONN;
            GetDescSource();

        }

        private void lbDBDesc_TextChanged(object sender, EventArgs e)
        {
            //_sqlBuilder.Database = lbDBDesc.Text;
        }

        public void UpdateFilterFrom(bool flag)
        {
            //Excel._Workbook wbook = (Excel._Workbook)_xlsApp.ActiveWorkbook;
            string address = "";
            string value = "";
            //try
            //{
            address = txtFilterFrom.Text;
            value = lbValueFrom.Text;
            //    //foreach (Excel._Worksheet isheet in wbook.Sheets)
            //    //{
            //    //    try
            //    //    {
            //    //        value = isheet.get_Range(address, Type.Missing).get_Value(Type.Missing).ToString();

            //    //        //_sqlBuilder.ParaValueList[i - 1] = value;
            //    //        break;

            //    //    }
            //    //    catch
            //    //    {
            //    //    }
            //    //}

            //    //lbValueFrom.Text = value;

            //}
            //catch
            //{

            //    lbValueFrom.Text = "_";
            //}

            if (dgvFilter.CurrentRow != null)
            {
                QueryBuilder.Filter filter = dgvFilter.CurrentRow.DataBoundItem as QueryBuilder.Filter;
                if (flag)
                {

                    if (value != "" && address != "")
                    {
                        //filter.FilterFrom = address;
                        filter.ValueFrom = value;
                        filter.FilterFromP = "{P}";
                        txtFilterFrom.ForeColor = Color.BlueViolet;
                    }
                    else
                    {
                        //filter.FilterFrom = txtFilterFrom.Text;
                        filter.ValueFrom = filter.FilterFrom;
                        filter.FilterFromP = "";
                        txtFilterFrom.ForeColor = Color.Black;
                    }
                    //dgvFilter.CurrentRow.InvalidateRow();
                }
                else
                {
                    filter.ValueFrom = txtFilterFrom.Text;
                    filter.FilterFromP = "";
                    txtFilterFrom.ForeColor = Color.Black;
                }
            }

        }
        public void UpdateFilterTo(bool flag)
        {
            //Excel._Workbook wbook = (Excel._Workbook)_xlsApp.ActiveWorkbook;
            string address = "";
            string value = "";
            //try
            //{
            address = txtFilterTo.Text;
            value = lbValueTo.Text;
            //    foreach (Excel._Worksheet isheet in wbook.Sheets)
            //    {
            //        try
            //        {
            //            value = isheet.get_Range(address, Type.Missing).get_Value(Type.Missing).ToString();

            //            //_sqlBuilder.ParaValueList[i - 1] = value;
            //            break;

            //        }
            //        catch
            //        {
            //        }
            //    }
            //    lbValueTo.Text = value;

            //}
            //catch
            //{

            //    lbValueTo.Text = "_";
            //}
            if (dgvFilter.CurrentRow != null)
            {
                QueryBuilder.Filter filter = dgvFilter.CurrentRow.DataBoundItem as QueryBuilder.Filter;
                if (flag)
                {

                    if (value != "" && address != "")
                    {
                        //filter.FilterTo = address;
                        filter.ValueTo = value;
                        filter.FilterToP = "{P}";
                        txtFilterTo.ForeColor = Color.BlueViolet;
                    }
                    else
                    {
                        //filter.FilterTo = txtFilterTo.Text;
                        filter.ValueTo = filter.FilterTo;
                        filter.FilterToP = "";
                        txtFilterTo.ForeColor = Color.Black;
                    }
                }
                else
                {
                    filter.ValueTo = txtFilterTo.Text;
                    filter.FilterToP = "";
                    txtFilterTo.ForeColor = Color.Black;
                }
                //dgvFilter.CurrentRow.InvalidateRow();
            }
        }
        private void UpdateLedger(bool flag)
        {
            string address = "";
            string value = "";
            //try
            //{
            value = lbLedgerDesc.Text;
            address = txtLedger.Text;
            //    lbLedgerDesc.Text = value;
            //    txtLedger.ForeColor = Color.BlueViolet;
            //}
            //catch
            //{
            //    txtLedger.ForeColor = Color.Black;
            //    lbLedgerDesc.Text = address;
            //}

            if (flag)
            {
                if (value != "" && address != "")
                {
                    _sqlBuilder.Ledger = address;
                    _sqlBuilder.LedgerV = value;
                    _sqlBuilder.LedgerP = "{P}";
                    txtLedger.ForeColor = Color.BlueViolet;
                }
                else
                {
                    _sqlBuilder.Ledger = txtLedger.Text;
                    _sqlBuilder.LedgerV = txtLedger.Text;
                    _sqlBuilder.LedgerP = "";
                    txtLedger.ForeColor = Color.Black;
                }
            }
            else
            {
                _sqlBuilder.Ledger = txtLedger.Text;
                _sqlBuilder.LedgerV = txtLedger.Text;
                _sqlBuilder.LedgerP = "";
                txtLedger.ForeColor = Color.Black;
            }
        }
        public void UpdateDatabase(bool flag)
        {
            //Excel._Workbook wbook = (Excel._Workbook)_xlsApp.ActiveWorkbook;
            string address = "";
            string value = "";
            //try
            //{
            address = txtDB.Text;
            value = lbValueTo.Text;
            //    foreach (Excel._Worksheet isheet in wbook.Sheets)
            //    {
            //        try
            //        {
            //            value = isheet.get_Range(address, Type.Missing).get_Value(Type.Missing).ToString();

            //            //_sqlBuilder.ParaValueList[i - 1] = value;
            //            break;

            //        }
            //        catch
            //        {
            //        }
            //    }
            //    lbDBDesc.Text = value;

            //}
            //catch
            //{
            //    lbDBDesc.Text = "_";
            //}
            if (flag)
            {
                if (value != "" && address != "")
                {
                    _sqlBuilder.Database = address;
                    _sqlBuilder.DatabaseV = value;
                    _sqlBuilder.DatabaseP = "{P}";
                    txtDB.ForeColor = Color.BlueViolet;
                }
                else
                {
                    txtDB.ForeColor = Color.Black;
                    DBInfoList tmp = DBInfoList.GetDBInfoList();
                    lbDBDesc.Text = "_";
                    foreach (QueryBuilder.DBInfo x in tmp)
                    {
                        if (x.Code == txtDB.Text.Trim())
                        {
                            lbDBDesc.Text = x.Description;

                            break;
                        }
                    }
                    _sqlBuilder.Database = txtDB.Text;
                    _sqlBuilder.DatabaseV = txtDB.Text;
                    _sqlBuilder.DatabaseP = "";

                }
            }
            else
            {
                txtDB.ForeColor = Color.Black;
                DBInfoList tmp = DBInfoList.GetDBInfoList();
                lbDBDesc.Text = "_";
                foreach (QueryBuilder.DBInfo x in tmp)
                {
                    if (x.Code == txtDB.Text.Trim())
                    {
                        lbDBDesc.Text = x.Description;

                        break;
                    }
                }
                _sqlBuilder.Database = txtDB.Text;
                _sqlBuilder.DatabaseV = txtDB.Text;
                _sqlBuilder.DatabaseP = "";
            }
        }

        private void txtFilterFrom_Validated(object sender, EventArgs e)
        {
            UpdateFilterFrom(false);
        }

        private void txtFilterTo_Validated(object sender, EventArgs e)
        {
            UpdateFilterTo(false);
        }

    }
}
