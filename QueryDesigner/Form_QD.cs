using BUS;
using DTO;
using FlexCel.Core;
using FlexCel.Pdf;
using FlexCel.Render;
using FlexCel.XlsAdapter;
using Janus.Windows.GridEX;
using QueryBuilder;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.IO;
using System.Management;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;

namespace QueryDesigner
{
    public partial class Form_QD : Form
    {
        #region IDHandle
        const int _hQD = 1;
        const int _hCreate = 2;
        const int _hAmmend = 3;
        const int _hDelete = 4;
        const int _hTransIn = 5;
        const int _hTransOut = 6;
        const int _hTemplate = 7;
        const int _hSystem = 8;
        const int _hAddress = 9;
        const int _hOperator = 10;
        const int _hImportDef = 11;
        const int _hImport = 12;
        const int _hTask = 13;
        #endregion IDHandle
        static QDConfig _config = new QDConfig();

        public static QDConfig Config
        {
            get { return Form_QD._config; }
            set { Form_QD._config = value; }
        }
        ReportGenerator _rpGen = null;
        clsChartProperty _propertyChart = new clsChartProperty();
        private SQLBuilder _sqlBuilder = new SQLBuilder(processingMode.Details);
        //GridViewComboBoxColumn customerColumn = new GridViewComboBoxColumn();
        public static string _strConnect = "";
        public static string _strConnectDes = "";
        //bool flag_view = false;
        String sErr = "";
        string _strY = "Y";
        string _strN = "N";
        string THEME = "Office2010";
        Node[] _arrNodes = null;
        string _hashTmp = null;
        public static string _key = "newoppo123456789";
        public static string _iv = "12345678";
        public static string _padMode = "PKCS7";
        public static string _opMode = "CBC";
        string owner = "";
        string _appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Replace("file:\\", "");
        string _pathLicense = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\License.bin";
        string _datePaterm = @"^([0]?[1-9]|[1|2][0-9]|[3][0|1])[./-]([0]?[1-9]|[1][0-2])[./-]([0-9]{4}|[0-9]{3}|[0-9]{2}|[0-9]{1})$";
        public static string __documentDirectory = string.Empty;
        public static string __reportPath = "";
        public static string __templatePath = "";
        string _processStatus = "";
        bool flagOpen = true;
        TreeNode _currentNode = null;
        static string _dtb = "";
        public static string DB
        {
            get { return _dtb; }
            set
            {
                if (!string.IsNullOrEmpty(value) && _dtb != value)
                {
                    _dtb = value;
                    CmdManager.Db = value;
                }
            }
        }
        public static string _user;
        //public static string _connDes;
        string _pass;
        string _DataXML = ""; string _filehtml = "";
        ExcelFile _xlsFile = null;
        //public Form_QD(1)
        //{
        //    InitializeComponent();
        //    InitDocument();
        //    //customerColumn.UniqueName = "Agregate";
        //    //customerColumn.HeaderText = "Agregate";

        //    //customerColumn.ValueMember = "Code";
        //    //customerColumn.DisplayMember = "Description";
        //    //customerColumn.Width = 100;
        //    //((GridViewComboBoxColumn)dgvSelectNodes.Columns["Agregate"]).DataSource = Parsing.GetListNumberAgregate();
        //    //ThemeResolutionService.ApplyThemeToControlTree(this, THEME);
        //    LoadConfig("");
        //}
        public Form_QD(string[] agrs)
        {
            _pathLicense = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\\License.bin";
            InitializeComponent();
            InitDocument();
            //string user = "";
            //string pass = "";
            if (agrs.Length >= 6)
            {
                string conID = agrs[0];
                if (conID == "ZZZ")
                    conID = "";
                string dtb = agrs[1];
                DB = dtb;

                _user = agrs[2];
                _pass = agrs[3];
                string qdid = agrs[4];
                LoadConfig(conID);
                if (_pass != "TVCSYS")
                {
                    flagOpen = false;

                }
                else
                {
                    BUS.PODControl podCtr = new PODControl();
                    DTO.PODInfo podInf = podCtr.Get(_user, ref sErr);
                    if (podInf.DB_DEFAULT != "")
                    {
                        DB = podInf.DB_DEFAULT;


                    }
                    else
                    {
                        DB = dtb;
                    }
                }
                //string permis = agrs[5];


                //if (agrs.Length >= 7)
                //    owner = agrs[6];


                BUS.LIST_QDControl qdCtr = new LIST_QDControl();
                DTO.LIST_QDInfo qdInfo = qdCtr.Get_LIST_QD(dtb, qdid, ref sErr);
                if (qdInfo.QD_ID != "")
                    LoadQD(qdInfo);
                else
                    _processStatus = "C";

                //SetPermission(permis);


                //btnEdit_Click(null, null);
                if (qdInfo.QD_ID != "")
                    tsMain.SelectedTab = tabItemQD;
            }
            else
            {
                _processStatus = "C";
                flagOpen = false;
                LoadConfig("");
                frmLogin frm = new frmLogin();
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    _user = frm.User;
                    _pass = frm.Pass;
                    DB = frm.DB;
                    //if (DB != "")
                    //{
                    //    DB = DB;
                    //    //txtdatabase.ReadOnly = true;
                    //    //bt_database.Enabled = false;
                    //}
                    flagOpen = true;
                }
                else
                {

                }
                //return;

                //customerColumn.UniqueName = "Agregate";
                //customerColumn.HeaderText = "Agregate";

                //customerColumn.ValueMember = "Code";
                //customerColumn.DisplayMember = "Description";
                //customerColumn.Width = 100;
                //((GridViewComboBoxColumn)dgvSelectNodes.Columns["Agregate"]).DataSource = Parsing.GetListNumberAgregate();
                ////ThemeResolutionService.ApplyThemeToControlTree(this, THEME);

            }
            _sqlBuilder.Database = DB;
        }

        //void SqlBuilder_Change()
        //{
        //    if (_rpGen != null && !(_rpGen.IsClose()))
        //    {
        //        _rpGen.Close();
        //        _filehtml = "";
        //        _DataXML = "";
        //    }
        //}

        private void SetPermissionSetup(string user)
        {
            BUS.PODControl podCtr = new PODControl();
            PODInfo podInf = podCtr.Get(user, ref sErr);
            BUS.POPControl popCtr = new POPControl();
            POPInfo popInf = popCtr.Get(podInf.ROLE_ID, DB, ref sErr);
            if (popInf.ROLE_ID == "")
                popInf = popCtr.Get(podInf.ROLE_ID, "", ref sErr);
            popInf.PERMISSION = popInf.PERMISSION.Replace(" ", popInf.DEFAULT_VALUE);
            if (popInf.PERMISSION == "")
                popInf.PERMISSION = new string('N', 100);
            if (popInf.PERMISSION.Substring(_hQD, 1) == "Y")
            {
                if (popInf.PERMISSION.Substring(_hSystem, 1) == "N")
                    btnSystem.Enabled = false;
                if (popInf.PERMISSION.Substring(_hAddress, 1) == "N")
                    _flagQDADD = false;
                if (popInf.PERMISSION.Substring(_hAmmend, 1) == "N")
                    btnEdit.Enabled = false;
                if (popInf.PERMISSION.Substring(_hCreate, 1) == "N")
                {
                    btnNew.Enabled = false;
                    btnCopy.Enabled = false;
                }
                if (popInf.PERMISSION.Substring(_hAmmend, 1) == "N")
                    if (popInf.PERMISSION.Substring(_hCreate, 1) == "N")
                        btnSave.Enabled = false;
                if (popInf.PERMISSION.Substring(_hDelete, 1) == "N")
                    btnDelete.Enabled = false;
                if (popInf.PERMISSION.Substring(_hOperator, 1) == "N")
                    btnOperator.Enabled = false;
                if (popInf.PERMISSION.Substring(_hTemplate, 1) == "N")
                    btnTemplate.Enabled = false;
                if (popInf.PERMISSION.Substring(_hTransIn, 1) == "N")
                    btnTransferIn.Enabled = false;
                if (popInf.PERMISSION.Substring(_hTransOut, 1) == "N")
                    btnTransferOut.Enabled = false;

                if (popInf.PERMISSION.Substring(_hImportDef, 1) == "N")
                    importDefinitionToolStripMenuItem.Visible = false;
                if (popInf.PERMISSION.Substring(_hImport, 1) == "N")
                    importToolStripMenuItem1.Visible = false;
                if (popInf.PERMISSION.Substring(_hTask, 1) == "N")
                    taskToolStripMenuItem.Visible = false;
            }
            else Close();
        }

        private void SetPermission(string permis)
        {
            dgvSelectNodes.Enabled = false;
            btUserFunc.Enabled = false;
            if (permis != null && permis != "" && permis.Length >= 6)
            {
                string insert = permis.Substring(1, 1);
                if (insert != _strY)
                {
                    btnNew.Visible = false;
                    btnCopy.Visible = false;
                }
                string update = permis.Substring(2, 1);
                if (update != _strY)
                {
                    btnEdit.Visible = false;
                    btnTemplate.Visible = false;
                }
                if (insert != _strY && update != _strY)
                {
                    btnSave.Visible = false;
                }
                if (update == _strY)
                {
                    dgvSelectNodes.Enabled = true;
                    btUserFunc.Enabled = true;
                }

                string delete = permis.Substring(3, 1);
                if (delete != _strY)
                {
                    btnDelete.Visible = false;
                }
                string tranferOut = permis.Substring(4, 1);
                if (tranferOut != _strY)
                {
                    btnTransferOut.Visible = false;
                }

                string tranferIn = permis.Substring(5, 1);
                if (tranferIn != _strY)
                {
                    btnTransferIn.Visible = false;
                }
                string read = permis.Substring(0, 1);
                if (read != _strY)
                {
                    btnView.Visible = false;
                    DisableForm();
                    //txtdatabase.Enabled = false;
                    bt_datasource.Enabled = false;
                    btnQD.Enabled = false;
                    twSchema.Enabled = true;
                    dgvFilter.Enabled = true;
                    btnParameter.Enabled = true;
                    btnEdit.Enabled = false;
                    btnDelete.Enabled = false;
                    btnTransferIn.Enabled = false;
                    btnTransferOut.Enabled = false;
                    btnSave.Enabled = false;
                    btnNew.Enabled = false;
                    btnTemplate.Enabled = false;
                    //bt_database.Enabled = false;
                }
                else
                {
                    twSchema.Enabled = true;
                    dgvFilter.Enabled = true;
                    btnParameter.Enabled = true;
                }
            }
            else
                flagOpen = false;

        }
        private void LoadConfig(string strAP)
        {
            try
            {
                if (File.Exists(__documentDirectory + "\\Configuration\\QDConfig.tvc"))
                {
                    _config.LoadConfig(__documentDirectory + "\\Configuration\\QDConfig.tvc");

                    string key = "";
                    _strConnect = _config.GetConnection(ref key, "QD");
                    SQLBuilder.SetConnection(_strConnect);
                    CommonControl.SetConnection(_strConnect);
                    _strConnectDes = _config.GetConnection(ref strAP, "AP");
                    _sqlBuilder.ConnID = strAP;

                    if (_config.DIR.Rows.Count > 0)
                    {
                        __templatePath = _config.DIR.Rows[0]["TMP"].ToString();
                        __reportPath = _config.DIR.Rows[0]["RPT"].ToString();
                    }

                    if (_config.SYS.Rows.Count > 0)
                    {
                        if (_config.SYS.Rows[0]["FONT"].ToString() != "")
                        {
                            TypeConverter tc = TypeDescriptor.GetConverter(typeof(Font));
                            Font font = (Font)tc.ConvertFrom(_config.SYS.Rows[0]["FONT"].ToString());
                            this.Font = font;
                        }
                        if (_config.SYS.Rows[0]["FORCECOLOR"].ToString() != "")
                        {
                            TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                            Color color = (Color)tc.ConvertFrom(_config.SYS.Rows[0]["FORCECOLOR"].ToString());
                            this.ForeColor = color;
                        }
                        else this.ForeColor = SystemColors.ControlText;
                        if (_config.SYS.Rows[0]["BACKCOLOR"].ToString() != "")
                        {
                            TypeConverter tc = TypeDescriptor.GetConverter(typeof(Color));
                            Color color = (Color)tc.ConvertFrom(_config.SYS.Rows[0]["BACKCOLOR"].ToString());
                            this.BackColor = color;
                        }
                        this.BackColor = SystemColors.Control;
                        ReportGenerator.User2007 = (bool)_config.SYS.Rows[0][_config.SYS.USE2007Column];
                    }
                    //string filename = GetDocumentDirec() + "\\Configuration\\xmlConnect.xml";

                    _config.SaveConfig(__documentDirectory + "\\Configuration\\QDConfig.tvc");

                    //Close();
                }

                Configuration.clsConfigurarion config = new Configuration.clsConfigurarion();
                config.GetDataTableDictionary(_appPath + "/Configuration/Languages.xml");
                //cboLanguage.Items. = config.DtDictionary;
                //cboLanguage.ValueMember = "Value";
                //cboLanguage.DisplayMember = "Code";
                dgvSelectNodes.AutoGenerateColumns = false;
                dgvFilter.AutoGenerateColumns = false;
                nodeBindingSource.DataSource = _sqlBuilder.SelectedNodes;
                filterBindingSource.DataSource = _sqlBuilder.Filters;
                DataTable dtcom = QueryBuilder.Parsing.GetListNumberAgregate();
                DataRow newrow = dtcom.NewRow();
                newrow["Code"] = newrow["Description"] = "";
                dtcom.Rows.Add(newrow);
                nodeAgregate.DataSource = dtcom;
                nodeAgregate.DisplayMember = "Code";
                nodeAgregate.ValueMember = "Code";

                ReportGenerator.Config = _config;
                CmdManager.Db = DB;
                CmdManager.AppConnect = _strConnect;
                CmdManager.RepConnect = _strConnectDes;
                CmdManager.ReptPath = __reportPath;
                CmdManager.TempPath = __templatePath;
            }
            catch (Exception ex)
            {
                lb_Err.Text = ex.Message;
            }
        }

        private string GetConnectionDes(string strAP)
        {
            if (_config.DTB.Rows.Count > 0)
            {
                foreach (DataRow row in _config.ITEM.Rows)
                {
                    if (strAP == "")
                    {
                        if (row["KEY"].ToString() == _config.DTB.Rows[0]["AP"].ToString())
                        {
                            _strConnectDes = row["CONTENT"].ToString();
                            _sqlBuilder.ConnID = row["KEY"].ToString();

                            //_connDes = row["KEY"].ToString();
                        }
                    }
                    else
                    {
                        //_connDes = strAP;
                        if (row["KEY"].ToString() == strAP)
                        {
                            _config.DTB.Rows[0]["AP"] = strAP;
                            _strConnectDes = row["CONTENT"].ToString();
                            _sqlBuilder.ConnID = strAP;

                        }
                    }
                }
            }
            return _strConnectDes;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            tsMain.TabPages.Remove(tabItemSQLText);

            webGadget.AllowNavigation = wbChart.AllowNavigation = true;
            dgvFilter.AutoGenerateColumns = false;
            dgvSelectNodes.AutoGenerateColumns = false;
            QueryBuilder.SQLBuilder.SQLDebugMode = Properties.Settings.Default.SQLDebugMode;
            IsNot.DataSource = QueryBuilder.Parsing.GetListIsNot();
            IsNot.ValueMember = "Code";
            IsNot.DisplayMember = "Description";
            Operate.DataSource = QueryBuilder.Parsing.GetListOperator("");
            Operate.ValueMember = "Code";
            Operate.DisplayMember = "Description";
            txtowner.Text = _user;
            if (_user != "TVC" || _pass != "TVCSYS")
            {
                SetPermissionSetup(_user);

                //dgvSelectNodes.Visible = false;
                //btUserFunc.Enabled = btnSave.Enabled = btnTemplate.Enabled = btnDelete.Enabled = false;
            }
            if (flagOpen == false)
                Close();
            //SQLBuilder.SQLDebugMode = Properties.Settings.Default.SQLDebugMode;
            //ThemeResolutionService.ApplicationThemeName = "Office2007Blue";          
            //TopMost = true;
            //ResetForm();



            try
            {
                if (!Directory.Exists(__templatePath))
                    Directory.CreateDirectory(__templatePath);
                if (!Directory.Exists(__reportPath))
                    Directory.CreateDirectory(__reportPath);
                string ext = "";

                if (!File.Exists(__templatePath + "-.template" + ReportGenerator.Ext))
                {
                    File.Copy(_appPath + "\\-.template" + ReportGenerator.Ext, __templatePath + "-.template" + ReportGenerator.Ext);
                }
            }
            catch (Exception ex)
            {
                lb_Err.Text = ex.Message;
            }
            dgvFilter.Invalidate();
            dgvSelectNodes.Invalidate();
            ValidateLicense();
            //TopMost = false;
            if (sErr != "")
                lb_Err.Text = sErr;
            splitContainer3.SplitterDistance = 1124;
            Text = "Query Desinger for WinForm - " + _user + "@" + DB;
            //frmLoading frm = new frmLoading();

            //frm.Show();
        }

        private void InitDocument()
        {
            string filename = _appPath + "\\Configuration\\xmlConnect.xml";
            __documentDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\TVC-QD";
            if (!Directory.Exists(__documentDirectory))
            {
                Directory.CreateDirectory(__documentDirectory);
            }
            string configureDirectory = __documentDirectory + "\\Configuration";
            if (!Directory.Exists(configureDirectory))
            {
                Directory.CreateDirectory(configureDirectory);
            }
            string layoutDirectory = __documentDirectory + "\\Layout";
            if (!Directory.Exists(layoutDirectory))
            {
                Directory.CreateDirectory(layoutDirectory);
            }
            string logFolder = __documentDirectory + "\\Log";
            if (!Directory.Exists(logFolder))
            {
                Directory.CreateDirectory(logFolder);
            }

            string connectionFile = configureDirectory + "\\xmlConnect.xml";
            if (!File.Exists(connectionFile))
            {
                File.Copy(filename, connectionFile);
            }
            string tmpLicense = _appPath + "\\license.bin";
            string fileLicense = configureDirectory + "\\license.bin";
            if (!File.Exists(fileLicense))
                if (File.Exists(tmpLicense))
                {
                    File.Copy(tmpLicense, fileLicense);
                }
        }


        //private void OnRadMenuItemChangeTheme_Click(object sender, EventArgs e)
        //{
        //    RadMenuItem menuItem = (RadMenuItem)sender;

        //    foreach (RadMenuItem sibling in menuItem.ParentItem.Items)
        //    {
        //        sibling.IsChecked = false;
        //    }

        //    menuItem.IsChecked = true;

        //    string themeName = (string)(menuItem).Tag;
        //    ChangeThemeName(this, themeName);
        //}

        //private void ChangeThemeName(Control control, string themeName)
        //{
        //    IComponentTreeHandler radControl = control as IComponentTreeHandler;
        //    if (radControl != null)
        //    {
        //        radControl.ThemeName = themeName;
        //    }

        //    foreach (Control child in control.Controls)
        //    {
        //        ChangeThemeName(child, themeName);
        //    }
        //}

        private void LoadQD(LIST_QDInfo info)
        {
            _sqlBuilder.Filters.Clear();
            _sqlBuilder.SelectedNodes.Clear();
            txt_sql.Text = "";
            Load_QDinfo(info);
            tsMain.SelectedTab = tabItemGeneral;
            //radTabStrip1.Se
            //txtdatabase.Focus();
            tabItemGeneral.Select();
            tsMain.Enabled = true;
            _sqlBuilder = SQLBuilder.LoadSQLBuilderFromDataBase(txtqd_id.Text.Trim(), DB.Trim(), txtdatasource.Text.Trim());
            nodeBindingSource.DataSource = _sqlBuilder.SelectedNodes;
            filterBindingSource.DataSource = _sqlBuilder.Filters;
            _propertyChart.ReadProperty(info.FOOTER_TEXT);
            _xlsFile = null;
        }


        private void bt_database_Click(object sender, EventArgs e)
        {

            Form_DTBView a = new Form_DTBView();
            a.themname = THEME;
            a.BringToFront();
            if (a.ShowDialog(this) == DialogResult.OK)
            {
                DB = a.Code_DTB;
                //txt_database.Text = a.Description_DTB;
            }
        }

        private void bt_datasource_Click(object sender, EventArgs e)
        {
            Form_TableView a = new Form_TableView(_dtb, _user);
            a.Code_DTB = DB;
            a.themname = THEME;
            a.BringToFront();
            if (a.ShowDialog(this) == DialogResult.OK)
            {
                txtdatasource.Text = a.Code_DTB;
                txt_datasource.Text = a.Description_DTB;
                txt_datasource.Focus();
            }
        }

        public void ResetForm()
        {
            _sqlBuilder.Filters.Clear();
            _sqlBuilder.SelectedNodes.Clear();
            _sqlBuilder.Table = "";
            _sqlBuilder.Database = "";

            //txt_database.Text = "";
            txt_datasource.Text = "";
            lbgroup.Text = "";

            txtqd_id.Text = "";
            txtdesr.Text = "";
            txtowner.Text = _user;
            txtANAL_Q2.Text = "";
            //if (DB != null)
            //    DB = DB;
            txtdatasource.Text = "";
            txt_sql.Text = "";
            lb_Err.Text = "";
            txtTmp.Text = "";

            ckbShared.Checked = false;
            //_sqlBuilder.SelectedNodes.Clear();
            //_sqlBuilder.Filters.Clear();
        }

        public void EnableForm()
        {
            //txt_database.Enabled = true;
            txt_datasource.Enabled = true;
            lbgroup.Enabled = true;

            txtqd_id.Enabled = true;
            txtdesr.Enabled = true;
            txtowner.Enabled = true;
            txtANAL_Q2.Enabled = true;
            //txtdatabase.Enabled = true;
            txtdatasource.Enabled = true;

            //bt_database.Enabled = true;
            bt_datasource.Enabled = true;
            bt_group.Enabled = true;

            //tsMain.Enabled = true;
            dgvPreview.Enabled = true;
            btUserFunc.Enabled = true;
            twSchema.Enabled = true;
            dgvFilter.Enabled = true;
            dgvSelectNodes.Enabled = true;
            txt_sql.Enabled = true;
            btUserFunc.Enabled = true;
            txtANAL_Q1.Enabled = true;
            btnParameter.Enabled = true;
            ckbShared.Enabled = true;
        }

        public void DisableForm()
        {
            txtANAL_Q1.Enabled = false;
            //txt_database.Enabled = false;
            txt_datasource.Enabled = false;
            lbgroup.Enabled = false;
            btUserFunc.Enabled = false;
            txtqd_id.Enabled = false;
            txtdesr.Enabled = false;
            //txtowner.Enabled = false;
            txtANAL_Q2.Enabled = false;
            //txtdatabase.Enabled = false;
            txtdatasource.Enabled = false;

            //bt_database.Enabled = false;
            bt_datasource.Enabled = false;
            bt_group.Enabled = false;

            //tsMain.Enabled = false;
            //dgvPreview.Enabled = false;
            btUserFunc.Enabled = false;
            //twSchema.Enabled = false;
            //dgvFilter.Enabled = false;
            dgvSelectNodes.Enabled = false;
            txt_sql.Enabled = false;
            btnParameter.Enabled = false;
            ckbShared.Enabled = false;
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            ResetForm();
            EnableForm();
            DB = DB;
            tsMain.SelectedTab = tabItemGeneral;
            if (owner != "")
            {
                txtowner.Text = owner;
                txtowner.Enabled = false;
            }
            _processStatus = "C";
        }

        private int IsExist(BindingList<QueryBuilder.Filter> bindingList, string para)
        {
            for (int i = 0; i < bindingList.Count; i++)
                if (bindingList[i].Code == para)
                    return i;
            return -1;
        }
        //  SQLBuilder.LoadSQLBuilderFromDataBase

        private ArrayList GetFieldName()
        {
            ArrayList arr = new ArrayList();
            //CommoControl commo = new CommoControl();
            //string connnectString = commo.CreateConnectString(Properties.Settings.Default.Server
            //            , Properties.Settings.Default.User
            //            , Properties.Settings.Default.Pass
            //            , Properties.Settings.Default.DBName);
            _sqlBuilder.StrConnectDes = _strConnectDes;
            DataTable dt = _sqlBuilder.BuildDataTable(txt_sql.Text);

            foreach (DataColumn colum in dt.Columns)
            {
                arr.Add(colum.ColumnName);
            }

            return arr;
        }


        private void LoadSQLBuilderFromDataBase()
        {
            foreach (DataGridViewRow row in dgvFilter.Rows)
            {
                if (row.DataBoundItem is QueryBuilder.Filter)
                {
                    QueryBuilder.Filter a = (QueryBuilder.Filter)row.DataBoundItem;
                    if (txt_sql.Text != "" && !Regex.IsMatch(a.Code, @"^@"))
                        row.Visible = false;
                    else
                        row.Visible = true;
                }
            }
            foreach (DataGridViewRow row in dgvSelectNodes.Rows)
            {
                if (row.DataBoundItem is QueryBuilder.Node)
                {
                    //QueryBuilder.Node a = (QueryBuilder.Node)row.DataBoundItem;
                    if (txt_sql.Text != "")
                        row.Visible = false;
                    else
                        row.Visible = true;
                }

            }
            dgvSelectNodes.AutoGenerateColumns = false;
            dgvFilter.AutoGenerateColumns = false;
            nodeBindingSource.DataSource = _sqlBuilder.SelectedNodes;
            filterBindingSource.DataSource = _sqlBuilder.Filters;
        }

        public string get_para(string x)
        {
            string[] kq = Regex.Split(x, @"(^[a-zA-Z_]+)");
            return kq[1];
            //string[] kq = x.Split(new char[] { '\'', ' ', '\n' });
            //return kq[0];

        }



        private bool CheckError_TabGeneral()
        {
            bool flag = false;
            String err = "";
            if (DB == "")
            {
                err = err + "- Database required !!!\n";
                flag = true;
            }
            if (txtqd_id.Text.Trim() == "")
            {
                err = err + "- Inquiry Code required !!!\n";
                flag = true;
            }
            if (txtdesr.Text == "")
            {
                err = err + "- Description required !!!\n";
                flag = true;
            }
            if (txtowner.Text == "")
            {
                err = err + "- Owner required !!!\n";
                flag = true;
            }
            if (txtdatasource.Text == "" && txt_sql.Text == "")//
            {
                err = err + "- DataSource required !!!\n";
                flag = true;
            }
            lb_Err.Text = err;


            return flag;
        }

        private bool CheckError_TabSQL()
        {
            bool flag = false;
            String err = "";
            if (DB == "")
            {
                err = err + "- Database required !!!\n";
                flag = true;
            }
            if (txtqd_id.Text.Trim() == "")
            {
                err = err + "- Inquiry Code required !!!\n";
                flag = true;
            }
            if (txtdesr.Text == "")
            {
                err = err + "- Description required !!!\n";
                flag = true;
            }
            if (txtowner.Text == "")
            {
                err = err + "- Owner required !!!\n";
                flag = true;
            }
            lb_Err.Text = err;


            return flag;
        }

        private void bt_Err_Click(object sender, EventArgs e)
        {
            if (lb_Err.Text != "")
                MessageBox.Show(lb_Err.Text, "Warning");
        }



        private void btnView_Click(object sender, EventArgs e)
        {
            if (DB != "")
            {
                Form_View a = new Form_View(DB, _user);
                a.themname = THEME;
                a.database = DB;
                a.BringToFront();
                if (a.ShowDialog() == DialogResult.OK)
                {
                    LoadQD(a.qdinfo);
                    txtdatasource_Validated(null, null);
                    _processStatus = "V";
                }
            }
            else
                lb_Err.Text = "insert dtb";
        }
        public void Load_QDinfo(LIST_QDInfo qdinfo)
        {
            if (qdinfo.QD_ID != "")
            {
                DB = qdinfo.DTB.Trim();
                txtqd_id.Text = qdinfo.QD_ID.Trim();
                txtdesr.Text = qdinfo.DESCRIPTN.Trim();
                txtowner.Text = qdinfo.OWNER.Trim();
                txtdatasource.Text = qdinfo.ANAL_Q0.Trim();
                txtANAL_Q2.Text = qdinfo.ANAL_Q2.Trim();
                txtANAL_Q1.Text = qdinfo.ANAL_Q1.Trim();
                txt_sql.Text = qdinfo.SQL_TEXT.Trim();
                ckbShared.Checked = qdinfo.SHARED;
                LIST_TEMPLATEControl ctr = new LIST_TEMPLATEControl();
                LIST_TEMPLATEInfo info = ctr.Get(_dtb, qdinfo.QD_ID, ref sErr);
                txtTmp.Text = info.Code;
                _processStatus = "V";
                DisableForm();
            }


        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            //  radTabStrip1.SelectedTab = tabItemGeneral;
            BUS.LIST_QDControl ctr = new LIST_QDControl();

            if (ctr.IsExist(DB, txtqd_id.Text))
            {
                EnableForm();
                txtqd_id.Enabled = false;
                _processStatus = "A";
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            sErr = "";
            if (_sqlBuilder.SelectedNodes.Count > 0 || txt_sql.Text != "")
            {
                LIST_QDControl qdControl = new LIST_QDControl();
                LIST_QDInfo qdInfo = new LIST_QDInfo();
                qdInfo.ANAL_Q0 = txtdatasource.Text;
                qdInfo.ANAL_Q1 = txtANAL_Q1.Text;
                qdInfo.DESCRIPTN = txtdesr.Text;
                qdInfo.DTB = DB;
                qdInfo.OWNER = txtowner.Text;
                qdInfo.QD_ID = txtqd_id.Text.Trim();
                qdInfo.ANAL_Q2 = txtANAL_Q2.Text;
                qdInfo.FOOTER_TEXT = _propertyChart.GetProperty();
                qdInfo.SHARED = ckbShared.Checked;
                //qdInfo.
                //      qdInfo.SQL_TEXT = _sqlBuilder.BuildSQL();
                qdInfo.SQL_TEXT = txt_sql.Text;
                if (_processStatus == "C" && !qdControl.IsExist(DB, qdInfo.QD_ID))
                    qdControl.Add_LIST_QD(qdInfo, ref sErr);
                else if (_processStatus == "A")
                    qdControl.Update_LIST_QD(qdInfo);
                else
                {
                    lb_Err.Text = "";
                    return;
                }
                LIST_QDDControl qddControl = new LIST_QDDControl();
                LIST_QDD_FILTERControl qddFilterCtr = new LIST_QDD_FILTERControl();
                if (sErr == "")
                {
                    qddControl.Delete_LIST_QDD_By_QD_ID(qdInfo.QD_ID, qdInfo.DTB, ref sErr);
                    qddFilterCtr.DeleteByQD_ID(qdInfo.QD_ID, qdInfo.DTB, ref sErr);
                    sErr = "";
                    int index = 1;
                    for (int i = 0; i < _sqlBuilder.Filters.Count; i++)
                    {

                        //if (!string.IsNullOrEmpty(_sqlBuilder.Filters[i].FilterFrom))
                        //{

                        LIST_QDDInfo qddInfo = new LIST_QDDInfo(qdInfo.DTB,
                            qdInfo.QD_ID, index,
                            _sqlBuilder.Filters[i].Node.Code,
                            _sqlBuilder.Filters[i].Node.Description,
                            _sqlBuilder.Filters[i].Node.FTypeFull,
                            (i + 1).ToString(),
                            _sqlBuilder.Filters[i].Node.Agregate,
                            _sqlBuilder.Filters[i].Node.NodeDesc,
                            _sqlBuilder.Filters[i].FilterFrom,
                            _sqlBuilder.Filters[i].FilterTo,
                            true);

                        qddControl.Add_LIST_QDD(qddInfo, ref sErr);
                        LIST_QDD_FILTERInfo qddFilter = new LIST_QDD_FILTERInfo(qdInfo.DTB
                            , qdInfo.QD_ID
                            , index
                            , _sqlBuilder.Filters[i].Operate
                            , _sqlBuilder.Filters[i].IsNot);
                        if (_sqlBuilder.Filters[i].Operate != "-" && _sqlBuilder.Filters[i].Operate != "")
                        {
                            qddFilterCtr.Add(qddFilter, ref sErr);
                            sErr = "";
                        }
                        index++;
                        //}
                    }
                    for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
                    {
                        LIST_QDDInfo qddInfo = new LIST_QDDInfo(qdInfo.DTB,
                            qdInfo.QD_ID, index,
                            _sqlBuilder.SelectedNodes[i].Code,
                            _sqlBuilder.SelectedNodes[i].Description,
                            _sqlBuilder.SelectedNodes[i].FTypeFull,
                            _sqlBuilder.SelectedNodes[i].Sort,
                            _sqlBuilder.SelectedNodes[i].Agregate,
                            _sqlBuilder.SelectedNodes[i].Expresstion,
                            "",
                            "",
                            false);
                        qddControl.Add_LIST_QDD(qddInfo, ref sErr);
                        index++;
                    }
                }
                if (sErr == "")
                {
                    UpdateTemplateToDB();
                    DisableForm();
                    _processStatus = "V";
                    //tabItemGeneral.IsSelected = true;
                }
                else
                    lb_Err.Text = sErr;
            }
        }



        private DataTable CreateFilterTable()
        {
            DataTable dt = new DataTable();
            DataRow row = dt.NewRow();
            if (_sqlBuilder.Filters.Count > 0)
            {
                dt.Rows.Add(row);

                for (int i = 0; i < _sqlBuilder.Filters.Count; i++)
                {
                    dt.Columns.Add(_sqlBuilder.Filters[i].Code + "_From");
                    dt.Columns.Add(_sqlBuilder.Filters[i].Code + "_To");

                    dt.Rows[0][_sqlBuilder.Filters[i].Code + "_From"] = _sqlBuilder.Filters[i].FilterFrom;
                    dt.Rows[0][_sqlBuilder.Filters[i].Code + "_To"] = _sqlBuilder.Filters[i].FilterTo;
                }
            }
            return dt;
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            BUS.POSControl posCtr = new BUS.POSControl();
            DTO.POSInfo infPOS = new POSInfo(_user, DB, "Query Designer", "QD", DateTime.Now.ToString("yyyy-MM-dd hh:mm"));
            posCtr.InsertUpdate(infPOS);
            //SaveFileDialog frm = new SaveFileDialog();
            //frm.Filter = "Excel file (*.xls)|*.xls";
            //if (frm.ShowDialog() == DialogResult.OK)
            //{
            //XlsFile xls = new XlsFile();
            //BUS.CommoControl commo = new CommoControl();
            //string connnectString = commo.CreateConnectString(Properties.Settings.Default.Server
            //             , Properties.Settings.Default.User
            //             , Properties.Settings.Default.Pass
            //             , Properties.Settings.Default.DBName);

            ReportGenerator rpGen = new ReportGenerator(_sqlBuilder, txtqd_id.Text, txt_sql.Text, _strConnectDes, __templatePath, __reportPath);
            LIST_TEMPLATEControl tmpCtr = new LIST_TEMPLATEControl();
            LIST_TEMPLATEInfo tempInfo = tmpCtr.Get(_dtb, txtqd_id.Text, ref sErr);

            if (tempInfo.Code != "")
            {
                rpGen.STemp = tempInfo.Data;
                rpGen.LengthTemp = tempInfo.Length;
            }
            //ExcelFile test = null;
            try
            {
                //string ext = "";

                string filename = __reportPath + txtdesr.Text + ReportGenerator.Ext;
                rpGen.Name = txtdesr.Text;
                if (CheckChange(_xlsFile, __templatePath, txtqd_id.Text))
                    _xlsFile = rpGen.CreateReport();
                _xlsFile.Save(filename, ReportGenerator._format);

                //AddData(reportPath + txtqd_id.Text.Trim() + ".xls");
                Process.Start(filename);

            }
            catch (Exception ex) { lb_Err.Text = ex.Message; }
            //}


        }

        private bool CheckChange(ExcelFile _xlsFile, string path, string qdid)
        {
            string ext = ".xls";
            if (_config.SYS[0].USE2007)
                ext = ".xlsx";
            string str1 = null;
            try
            {
                using (FileStream rd = new FileStream(path + qdid + ".template" + ext, FileMode.Open, FileAccess.Read, System.IO.FileShare.ReadWrite))
                {
                    str1 = MyHash.Hash(rd, "MD5");
                }
            }
            catch { return true; }

            if (str1.Equals(_hashTmp))
            {
                _hashTmp = str1;
                return true;
            }
            else if (_xlsFile == null)
            {
                return true;
            }
            return false;
        }


        private void btnTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                btnEdit_Click(null, null);

                if (txtTmp.Text == "")
                {
                    //      File.Delete(saveFileDialog1.FileName);
                    string currentPath = _appPath + "\\";

                    if (!File.Exists(__templatePath + txtqd_id.Text.Trim() + ".template" + ReportGenerator.Ext))
                    {
                        XlsFile xlsTemp = new XlsFile(currentPath + "-.template" + ReportGenerator.Ext);
                        xlsTemp.SetCellValue(xlsTemp.GetSheetIndex("<#Config>"), 10, 2, txtqd_id.Text.Trim(), 0);
                        xlsTemp.SetCellValue(xlsTemp.GetSheetIndex("<#Config>"), 11, 2, "FilterPara", 0);
                        xlsTemp.SetCellValue(xlsTemp.GetSheetIndex("<#Config>"), 12, 2, "params", 0);

                        xlsTemp.Save(__templatePath + txtqd_id.Text.Trim() + ".template" + ReportGenerator.Ext);
                    }
                    Process.Start(__templatePath + txtqd_id.Text.Trim() + ".template" + ReportGenerator.Ext);

                }
                else if (File.Exists(txtTmp.Text))
                {
                    Process.Start(txtTmp.Text);
                }
                else
                {
                    LIST_TEMPLATEControl ctr = new LIST_TEMPLATEControl();
                    LIST_TEMPLATEInfo info = ctr.Get(_dtb, txtqd_id.Text, ref sErr);
                    string filename = __templatePath + txtqd_id.Text.Trim() + ".template" + ReportGenerator.Ext;
                    using (FileStream fs = new FileStream(filename, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Write))
                    {
                        fs.Write(info.Data, 0, info.Data.Length);
                    }
                    txtTmp.Text = filename;
                    Process.Start(filename);
                }
                //flexCelReport1.AddTable(temp);
                //AutoRun(temp.Tables[0], flag_filter);
                //Set_TT_XLB_EB(path_template);

                //}
                //else
                //{
                //}



                //ListFieldData a = new ListFieldData();
                #region Parameter
                DataTable dt_filter = new DataTable();
                dt_filter.TableName = "parameter";
                dt_filter.Columns.Add("Name");
                dt_filter.Columns.Add("Code");

                if (_sqlBuilder.Filters.Count > 0)
                {
                    for (int i = 0; i < _sqlBuilder.Filters.Count; i++)
                    {
                        dt_filter.Rows.Add(new string[] { _sqlBuilder.Filters[i].Description + "_From", "parameter." + _sqlBuilder.Filters[i].Description + "_From" });
                        dt_filter.Rows.Add(new string[] { _sqlBuilder.Filters[i].Description + "_To", "parameter." + _sqlBuilder.Filters[i].Description + "_To" });
                    }
                    //a.dt_Filter = dt_filter;
                }
                DataTable dt_param = new DataTable();
                DataColumn[] cols = new DataColumn[] { new DataColumn("Code")
                    , new DataColumn("Name")};
                dt_param.Columns.AddRange(cols);
                dt_param.TableName = "params";

                dt_param.Rows.Add("Code", "Code");
                dt_param.Rows.Add("Description", "Description");
                dt_param.Rows.Add("ValueFrom", "ValueFrom");
                dt_param.Rows.Add("ValueTo", "ValueTo");
                dt_param.Rows.Add("IsNot", "IsNot");
                dt_param.Rows.Add("Operate", "Operate");

                #endregion Parameter
                #region Field
                DataTable dt_list = new DataTable();
                if (_sqlBuilder.SelectedNodes.Count > 0)
                {
                    //CommoControl commo = new CommoControl();
                    //string connnectString = commo.CreateConnectString(Properties.Settings.Default.Server
                    //         , Properties.Settings.Default.User
                    //         , Properties.Settings.Default.Pass
                    //         , Properties.Settings.Default.DBName);
                    DataTable rs = _sqlBuilder.BuildDataTableStruct(txt_sql.Text, _strConnectDes);

                    //a.THEME = this.THEME;
                    dt_list.TableName = "data";
                    dt_list.Columns.Add("Name");
                    dt_list.Columns.Add("Code");

                    foreach (DataColumn colum in rs.Columns)
                    {
                        string desc = colum.ColumnName;
                        foreach (Node node in _sqlBuilder.SelectedNodes)
                        {
                            if (node.MyCode == colum.ColumnName)
                            {
                                desc = node.Description;
                                break;
                            }
                        }
                        dt_list.Rows.Add(new string[] { desc, colum.ColumnName });
                    }


                    //a.dt_list = dt;


                }
                else
                {

                    //a.THEME = this.ThemeName;

                    dt_list.Columns.Add("Name");
                    dt_list.Columns.Add("Code");


                    ArrayList arr = GetFieldName();
                    if (arr.Count > 0)
                    {

                        for (int i = 0; i < arr.Count; i++)
                        {
                            dt_list.Rows.Add(new string[] { arr[i].ToString(), arr[i].ToString() });
                        }
                        //a.dt_list = dt;


                    }

                }
                #endregion Field
                TVCDesigner.MainForm frm = new TVCDesigner.MainForm(dt_list, dt_filter, dt_param);

                //frm.BringToFront();

                frm.Show();
                this.MinimizeBox = true;
            }
            catch (Exception ex)
            {
                lb_Err.Text = ex.Message;
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            LIST_QDDControl qddcontrol = new LIST_QDDControl();
            LIST_QDControl qdcontrol = new LIST_QDControl();


            if (txtqd_id.Text != "")
            {

                if (qdcontrol.IsExist(DB.Trim(), txtqd_id.Text.Trim()))
                {
                    if (MessageBox.Show("Do you want to delete this QD?", "Delete Message", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        sErr = qdcontrol.Delete_LIST_QD(DB.Trim(), txtqd_id.Text.Trim());
                        qddcontrol.Delete_LIST_QDD_By_QD_ID(txtqd_id.Text.Trim(), DB.Trim(), ref sErr);

                        _processStatus = "";
                    }
                    if (sErr == "")
                    {
                        DeleteTemplateToDB();
                    }
                    lb_Err.Text = sErr;
                }
                else
                    lb_Err.Text = "Query Designer Code is not exist.";
            }
            else
                lb_Err.Text = "Query Designer Code is not exist.";
        }



        private void btnNew_MouseMove(object sender, MouseEventArgs e)
        {
            if (sender is PictureBox)
                ((PictureBox)sender).BackColor = Color.LightBlue;
        }

        private void btnNew_MouseDown(object sender, MouseEventArgs e)
        {
            if (sender is PictureBox)
                ((PictureBox)sender).BackColor = Color.Blue;
        }

        private void btnNew_MouseLeave(object sender, EventArgs e)
        {
            if (sender is PictureBox)
                ((PictureBox)sender).BackColor = Color.Transparent;
        }

        private void txtdatabase_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                bt_database_Click(null, null);
        }

        private void txtdatasource_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                bt_datasource_Click(null, null);
        }

        private void btnQD_Click(object sender, EventArgs e)
        {
            btnView_Click(sender, e);
        }

        private void txtqd_id_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                btnQD_Click(null, null);
        }


        private void btnFirst_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.StartPage = 1;
        }

        private void btnPrev_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.StartPage--;
        }

        private void btnNext_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.StartPage++;
        }

        private void btnLast_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.StartPage = flexCelPreview1.TotalPages;
        }

        private void btnZoomOut_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.Zoom -= 0.1f;
        }

        private void btnZoomIn_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.Zoom += 0.1f;
        }

        private void btnPdf_Click(object sender, System.EventArgs e)
        {
            if (flexCelImgExport1.Workbook == null)
            {
                MessageBox.Show("There is no open file");
                return;
            }
            if (PdfSaveFileDialog.ShowDialog() != DialogResult.OK) return;

            using (FlexCelPdfExport PdfExport = new FlexCelPdfExport(flexCelImgExport1.Workbook, true))
            {
                if (!DoExportToPdf(PdfExport)) return;
            }

            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            Process.Start(PdfSaveFileDialog.FileName);
        }

        private bool DoExportToPdf(FlexCelPdfExport PdfExport)
        {
            PdfThread MyPdfThread = new PdfThread(PdfExport, PdfSaveFileDialog.FileName, true);
            Thread PdfExportThread = new Thread(new ThreadStart(MyPdfThread.ExportToPdf));
            PdfExportThread.Start();
            using (PdfProgressDialog Pg = new PdfProgressDialog())
            {
                Pg.ShowProgress(PdfExportThread, PdfExport);
                if (Pg.DialogResult != DialogResult.OK)
                {
                    PdfExport.Cancel();
                    PdfExportThread.Join(); //We could just leave the thread running until it dies, but there are 2 reasons for waiting until it finishes:
                    //1) We could dispose it before it ends. This is workaroundable.
                    //2) We might change its workbook object before it ends (by loading other file). This will surely bring issues.
                    return false;
                }

                if (MyPdfThread != null && MyPdfThread.MainException != null)
                {
                    throw MyPdfThread.MainException;
                }
            }
            return true;
        }
        private void btnRecalc_Click(object sender, System.EventArgs e)
        {
            if (flexCelImgExport1.Workbook == null)
            {
                MessageBox.Show("Please open a file before recalculating.");
                return;
            }
            flexCelImgExport1.Workbook.Recalc(true);
            flexCelPreview1.InvalidatePreview();

        }
        private void cbAntiAlias_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            switch (cbAntiAlias.SelectedIndex)
            {
                case 1:
                    flexCelPreview1.SmoothingMode = SmoothingMode.HighQuality;
                    break;
                case 2:
                    flexCelPreview1.SmoothingMode = SmoothingMode.HighSpeed;
                    break;
                default:
                    flexCelPreview1.SmoothingMode = SmoothingMode.AntiAlias;
                    break;
            }

            if (flexCelImgExport1.Workbook != null) flexCelPreview1.InvalidatePreview();
        }
        #region PdfThread
        class PdfThread
        {
            private FlexCelPdfExport PdfExport;
            private string FileName;
            private bool AllVisibleSheets;
            private Exception FMainException;

            internal PdfThread(FlexCelPdfExport aPdfExport, string aFileName, bool aAllVisibleSheets)
            {
                PdfExport = aPdfExport;
                FileName = aFileName;
                AllVisibleSheets = aAllVisibleSheets;
            }

            internal void ExportToPdf()
            {
                try
                {
                    if (AllVisibleSheets)
                    {
                        try
                        {
                            using (FileStream f = new FileStream(FileName, FileMode.Create, FileAccess.Write))
                            {
                                PdfExport.BeginExport(f);
                                PdfExport.PageLayout = TPageLayout.Outlines;
                                PdfExport.ExportAllVisibleSheets(false, System.IO.Path.GetFileNameWithoutExtension(FileName));
                                PdfExport.EndExport();
                            }
                        }
                        catch
                        {
                            try
                            {
                                File.Delete(FileName);
                            }
                            catch
                            {
                                //Not here.
                            }
                            throw;
                        }
                    }
                    else
                    {
                        PdfExport.PageLayout = TPageLayout.None;
                        PdfExport.Export(FileName);
                    }
                }
                catch (Exception ex)
                {
                    FMainException = ex;
                }
            }

            internal Exception MainException
            {
                get
                {
                    return FMainException;
                }
            }
        }
        #endregion
        private void edPage_Leave(object sender, System.EventArgs e)
        {
            ChangePages();
        }

        private void edPage_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                ChangePages();
            if (e.KeyChar == (char)27)
                UpdatePages();
        }
        private void UpdatePages()
        {
            edPage.Text = String.Format("{0} of {1}", flexCelPreview1.StartPage, flexCelPreview1.TotalPages);
        }
        private void flexCelPreview1_StartPageChanged(object sender, System.EventArgs e)
        {
            UpdatePages();
        }

        private void ChangePages()
        {
            string s = edPage.Text.Trim();
            int pos = 0;
            while (pos < s.Length && s[pos] >= '0' && s[pos] <= '9') pos++;
            if (pos > 0)
            {
                int page = flexCelPreview1.StartPage;
                try
                {
                    page = Convert.ToInt32(s.Substring(0, pos));
                }
                catch (Exception ex)
                {
                    lb_Err.Text = ex.Message;
                }

                flexCelPreview1.StartPage = page;
            }
            UpdatePages();
        }

        private void flexCelPreview1_ZoomChanged(object sender, System.EventArgs e)
        {
            UpdateZoom();
        }

        private void UpdateZoom()
        {
            edZoom.Text = String.Format("{0}%", (int)Math.Round(flexCelPreview1.Zoom * 100));
        }

        private void ChangeZoom()
        {
            string s = edZoom.Text.Trim();
            int pos = 0;
            while (pos < s.Length && s[pos] >= '0' && s[pos] <= '9') pos++;
            if (pos > 0)
            {
                int zoom = (int)Math.Round(flexCelPreview1.Zoom * 100);
                try
                {
                    zoom = Convert.ToInt32(s.Substring(0, pos));
                }
                catch (Exception ex)
                {
                    lb_Err.Text = ex.Message;
                }

                flexCelPreview1.Zoom = zoom / 100f;
            }
            UpdateZoom();
        }

        private void edZoom_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                ChangeZoom();
            if (e.KeyChar == (char)27)
                UpdateZoom();
        }

        private void edZoom_Enter(object sender, System.EventArgs e)
        {
            ChangeZoom();
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (_sqlBuilder.SelectedNodes.Count > 0 && txtqd_id.Text != "")
            {
                txtqd_id.Enabled = true;
                EnableForm();
                //txtdatabase.Enabled = true;
                txtdatasource.Enabled = true;
                //txtqd_id.Text = "";
                txtqd_id.Focus();
                _processStatus = "C";
                dgvSelectNodes.Enabled = true;
            }
        }




        private void cboLanguage_SelectedValueChanged(object sender, EventArgs e)
        {
            String language = "44"; //cboLanguage.SelectedValue.ToString();
            Configuration.clsConfigurarion.SetLanguages(this, "Query Designer", language);
        }

        private void twSchema_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            TreeNode tmpNode = twSchema.SelectedNode;

            if (tmpNode != null && dgvSelectNodes.AllowDrop == true && dgvSelectNodes.Enabled && tmpNode.Nodes.Count == 0)
            {
                bool flag = true;
                //string[] arrNode = tmpNode.Tag.ToString().Split(';');
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

        private void dgvSelectNodes_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (dgvSelectNodes.CurrentRow.DataBoundItem is QueryBuilder.Node)
            {
                QueryBuilder.Node x = (QueryBuilder.Node)dgvSelectNodes.CurrentRow.DataBoundItem;
                FrmSubField subFDlg = new FrmSubField();
                subFDlg.Description = x.Description;

                subFDlg.Index = x.Index;
                subFDlg.Length = x.Length;
                if (subFDlg.ShowDialog() == DialogResult.OK)
                {
                    x.Description = subFDlg.Description;
                    x.FTypeFull = (x.FType + "      ").Substring(0, 5) + subFDlg.Index.ToString("00") + subFDlg.Length.ToString("00");
                }
            }
        }

        private void btUserFunc_Click(object sender, EventArgs e)
        {
            Node tmp = null;
            if (dgvSelectNodes.CurrentRow.DataBoundItem is QueryBuilder.Node)
            {
                if (((Node)dgvSelectNodes.CurrentRow.DataBoundItem).Expresstion.Trim() != "")
                    tmp = dgvSelectNodes.CurrentRow.DataBoundItem as Node;
            }
            FrmUserFunc frmDlg = new FrmUserFunc(_sqlBuilder.SelectedNodes, tmp);
            if (frmDlg.ShowDialog() == DialogResult.OK)
            {
                if (!_sqlBuilder.SelectedNodes.Contains(frmDlg.Node))
                    _sqlBuilder.SelectedNodes.Add(frmDlg.Node);
                else
                {
                    foreach (Node x in _sqlBuilder.SelectedNodes)
                    {
                        if (x.Code == frmDlg.Node.Code)
                        {
                            x.Expresstion = frmDlg.Node.Expresstion;
                        }
                    }
                }
                //dgvSelectNodes.DataSource = _sqlBuilder.SelectedNodes;
            }
        }

        private void radMenuItem1_Click(object sender, EventArgs e)
        {
            //QDAddin addin = new QDAddin("A1",;
            //addin.Show();
        }

        private void txtdatasource_TextChanged(object sender, EventArgs e)
        {
            try
            {
                _sqlBuilder.Database = DB;
                BindingList<Node> list = SchemaDefinition.GetDecorateTableByCode(txtdatasource.Text.Trim(), _sqlBuilder.Database);
                twSchema = TreeViewLoader.LoadTree(ref twSchema, list, txtdatasource.Text.Trim(), "");
                BUS.LIST_QD_SCHEMAControl ctr = new LIST_QD_SCHEMAControl();
                DTO.LIST_QD_SCHEMAInfo inf = ctr.Get(DB, txtdatasource.Text, ref sErr);
                string key = inf.DEFAULT_CONN;
                _strConnectDes = Form_QD._config.GetConnection(ref key, "AP");
            }
            catch (Exception ex) { lb_Err.Text = ex.Message; }
        }

        private void btnTransferIn_Click(object sender, EventArgs e)
        {
            FrmTransferIn frm = new FrmTransferIn("QD");
            frm.ShowDialog();
        }

        private void btnTransferOut_Click(object sender, EventArgs e)
        {
            FrmTransferOut frm = new FrmTransferOut(DB, "QD");
            //frm.DTB = txt_database.Text;
            frm.QD_CODE = txtqd_id.Text;
            frm.ShowDialog();
        }

        private void btnParameter_Click(object sender, EventArgs e)
        {
            FrmParam frmDlg = new FrmParam();
            if (frmDlg.ShowDialog() == DialogResult.OK)
            {
                _sqlBuilder.Filters.Add(frmDlg.Filter);
                //dgvFilter.DataSource = _sqlBuilder.Filters;
            }
        }

        public void GetQueryBuilderFromFomular(string formular)
        {
            if (formular.Contains("USER TABLE"))
            {
                //Excel._Worksheet sheet = (Excel._Worksheet)_xlsApp.ActiveWorkbook.ActiveSheet;
                //string vParamsString = Regex.Match(formular, @"\" +  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ System.Convert.ToChar(34) + @"\,.+?\)").Value.ToString();

                //// fill to parameter Array
                //int i = 0, n = 0;

                //if (!(string.IsNullOrEmpty(vParamsString)))
                //{
                //    vParamsString =  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vParamsString.Substring(1);
                //    vParamsString = vParamsString.Substring(1, vParamsString.Length - 2);// Strings.Mid(vParamsString, 1,  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vParamsString.Length - 1); 
                //    vParamsString = vParamsString + ","; //  them dau , cho de xu ly
                //    n = Regex.Matches(vParamsString, ".*?,").Count; //  cac tham so
                //    if (n > 0)
                //    {
                //        string[] vParameter = new string[n]; // tham so dau tien la vi tri cua cong thuc
                //        foreach (System.Text.RegularExpressions.Match p in Regex.Matches(vParamsString, ".*?,"))
                //        {
                //            i = i + 1;
                //            if (i == 1)
                //            {
                //                _sqlBuilder.Pos = p.Value.ToString().Replace(",", string.Empty);

                //            }
                //            else
                //            {

                //                string address = p.Value.ToString().Replace(",", string.Empty);
                //                string value = "";
                //                try
                //                {
                //                    value = sheet.get_Range(address, Type.Missing).get_Value(Type.Missing).ToString();
                //                    _sqlBuilder.ParaValueList[i - 1] = value;
                //                }
                //                catch
                //                {
                //                }
                //                //vParameter[i - 1] = p.Value.ToString().Replace(",", string.Empty);
                //            }

                //        }
                //    }
                //}
                Parsing.Formular2SQLBuilder(formular, ref _sqlBuilder);

                //SetDataToForm();
            }
        }

        private void btnExpandAll_Click(object sender, EventArgs e)
        {
            ExpandGrid();
        }

        private void ExpandGrid()
        {
            //GridExpandAnimationType temp = this.dgvPreview.GroupExpandAnimationType;
            //this.dgvPreview.GroupExpandAnimationType = GridExpandAnimationType.None;
            //for (int i = 0; i < this.dgvPreview.Groups.Count; i++)
            //{
            //    this.Expand(this.dgvPreview.Groups[i]);
            //}
            //this.dgvPreview.GroupExpandAnimationType = temp;
        }

        private void btnCollapseAll_Click(object sender, EventArgs e)
        {
            CollapseGrid();
        }

        private void CollapseGrid()
        {
            //GridExpandAnimationType temp = this.dgvPreview.GroupExpandAnimationType;
            //this.dgvPreview.GroupExpandAnimationType = GridExpandAnimationType.None;
            //for (int i = 0; i < this.dgvPreview.Groups.Count; i++)
            //{
            //    this.Collapse(this.dgvPreview.Groups[i]);
            //}
            //this.dgvPreview.GroupExpandAnimationType = temp;
        }

        private void radMenuItem1_Click_1(object sender, EventArgs e)
        {
            FrmSystem frm = new FrmSystem();
            frm.THEME = THEME;
            if (frm.ShowDialog() == DialogResult.OK)
                LoadConfig("");
        }

        private void btnLicense_Click(object sender, EventArgs e)
        {
            FrmLicense frm = new FrmLicense();
            frm.THEME = THEME;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                ValidateLicense();
            }
        }
        bool _flagQDADD = false;
        private void ValidateLicense()
        {
            tsMain.Enabled = false;
            BUS.CommonControl ctr = new CommonControl();
            try
            {

                object data = ctr.executeScalar(@"SELECT SUN_DATA  FROM SSINSTAL WHERE INS_TB='LCS' and INS_KEY='QD'");

                if (data != null)//File.Exists(_pathLicense.Replace("file:\\", ""))
                {
                    //StreamReader reader = new StreamReader(_pathLicense.Replace("file:\\", ""));
                    //string result = reader.ReadLine();
                    string kq = RC2.DecryptString(data.ToString(), _key, _iv, _padMode, _opMode);
                    string[] tmp = kq.Split(';');
                    DTO.License license = new DTO.License();
                    license.CompanyName = tmp[0];
                    license.ExpiryDate = Convert.ToInt32(tmp[1]);
                    license.Modules = tmp[2];
                    license.NumUsers = Convert.ToInt32(tmp[3]);
                    license.SerialNumber = tmp[4];
                    license.Key = tmp[5];
                    //license.SerialCPU = tmp[6];
                    license.SerialCPU = ctr.executeScalar("SELECT   CONVERT(varchar(200), SERVERPROPERTY('servername'))").ToString(); //"BFEBFBFF000006FD";
                    //reader.Close();


                    string param = license.CompanyName + license.SerialNumber + license.NumUsers.ToString() + license.Modules + license.ExpiryDate.ToString() + license.SerialCPU;


                    string temp = RC2.EncryptString(param, _key, _iv, _padMode, _opMode);
                    string key = Convert.ToBase64String(new System.Security.Cryptography.SHA1CryptoServiceProvider().ComputeHash(Encoding.ASCII.GetBytes(temp)));
                    if (key == license.Key)
                    {
                        int now = Convert.ToInt32(DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00"));
                        //BUS.CommonControl ctr = new CommonControl();
                        object dt = ctr.executeScalar("select getdate()", _strConnect);//curdate()
                        if (dt != null && dt is DateTime)
                        {
                            now = Convert.ToInt32(((DateTime)dt).Year.ToString() + ((DateTime)dt).Month.ToString("00") + ((DateTime)dt).Day.ToString("00"));
                        }
                        if (now > license.ExpiryDate)
                        {
                            toolStrip1.Enabled = tsMain.Enabled = false;
                            lb_Err.Text = "Your license is expired!";
                        }
                        else
                        {
                            if (license.Modules.Length >= 4 && license.Modules.Substring(3, 1) == "Y")
                                _flagQDADD = true;
                            //else _flagQDADD = false;

                            if (license.Modules.Length >= 5 && license.Modules.Substring(4, 1) == "Y")
                                taskToolStripMenuItem.Visible = true;
                            //else taskToolStripMenuItem.Visible = true;
                            BUS.POSControl ctrPOS = new POSControl();
                            if (ctrPOS.GetCount(ref sErr) > license.NumUsers)
                            {
                                toolStrip1.Enabled = tsMain.Enabled = false;
                                lb_Err.Text = "Current number of users has exceeded limit!";
                            }
                            else
                                toolStrip1.Enabled = tsMain.Enabled = true;

                        }
                    }
                    else
                    {
                        toolStrip1.Enabled = tsMain.Enabled = false;
                        lb_Err.Text = "Application have not license!";
                    }

                }
                else
                {
                    toolStrip1.Enabled = tsMain.Enabled = false;
                    lb_Err.Text = "Application have not license!";
                }
            }
            catch (Exception ex)
            {
                toolStrip1.Enabled = tsMain.Enabled = false;
                lb_Err.Text = ex.Message;
            }

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

        private void tabItemQD_Selected()
        {
            _xlsFile = null;
            if (CheckError_TabGeneral())
            {
                tsMain.SelectedTab = tabItemGeneral;
                //radTabStrip1.Se
                //txtdatabase.Focus();
                tabItemGeneral.Select();
                //bt_Err.Enabled = true;
            }
            else
            {
                if (txt_sql.Text == "")
                {
                    //dgvSelectNodes.CellBeginEdit += new GridViewCellCancelEventHandler(dgvSelectNodes_CellBeginEdit);

                    _sqlBuilder.Database = DB.Trim();
                    if (txtdatasource.Text.Trim() != "" && (txtdatasource.Text.ToString().Trim() != _sqlBuilder.Table.Trim() || twSchema.Nodes.Count == 0))
                    {
                        if (txtdatasource.Text.ToString().Trim() != _sqlBuilder.Table.Trim())
                        {
                            _sqlBuilder.Filters.Clear();
                            _sqlBuilder.SelectedNodes.Clear();
                        }

                    }
                    //if (flag_view)
                    //{
                    //twSchema.Enabled = false;
                    //dgvFilter.Enabled = false;
                    //dgvSelectNodes.Enabled = false;
                    //_sqlBuilder = SQLBuilder.LoadSQLBuilderFromDataBase(txtqd_id.Text, DB, txtdatasource.Text);
                    //LoadSQLBuilderFromDataBase();
                    //}
                    _sqlBuilder.Table = txtdatasource.Text.Trim();

                }
                else
                {
                    //if (flag_view)
                    //{
                    //twSchema.Enabled = false;
                    //dgvFilter.Enabled = false;
                    //dgvSelectNodes.Enabled = false;
                    //_sqlBuilder = SQLBuilder.LoadSQLBuilderFromDataBase(txtqd_id.Text, DB, txtdatasource.Text);
                    //LoadSQLBuilderFromDataBase();
                    //}
                    //_sqlBuilder.Filters.Clear();
                    //_sqlBuilder.SelectedNodes.Clear();

                    String query = txt_sql.Text + " ";
                    String[] array_filter = query.Split('@');
                    if (array_filter.Length > 0)
                    {
                        for (int i = 1; i < array_filter.Length; i++)
                        {
                            string para = get_para(array_filter[i]);
                            para = "@" + para;

                            QueryBuilder.Filter tmp = new QueryBuilder.Filter(new Node(para, para));
                            if (IsExist(_sqlBuilder.Filters, para) == -1)
                                _sqlBuilder.Filters.Add(tmp);
                        }
                    }
                    //twSchema.Enabled = false;


                    //     dgvFilter.AllowDrop = false;
                }


                LoadSQLBuilderFromDataBase();
            }
        }

        private void tabItemPreview_Selected()
        {
            try
            {
                BUS.POSControl posCtr = new BUS.POSControl();
                DTO.POSInfo infPOS = new POSInfo(_user, DB, "Query Designer", "QD", DateTime.Now.ToString("yyyy-MM-dd hh:mm"));
                posCtr.InsertUpdate(infPOS);
                //CommoControl commo = new CommoControl();
                _sqlBuilder.StrConnectDes = _strConnectDes;
                DataTable dt = _sqlBuilder.BuildDataTable(txt_sql.Text);

                //dgvPreview.RootTable.Columns.Clear();
                //dgvPreview.DataSource = new DataTable();
                //dgvPreview.RootTable.Columns.BeginUpdate();
                dgvPreview.DataSource = dt;
                dgvPreview.RetrieveStructure();
                //foreach (QueryBuilder.Node node in _sqlBuilder.SelectedNodes)
                //{
                //    //if (dgvPreview.RootTable.Columns.IndexOf(node.MyCode) > 0)
                //    if (dgvPreview.RootTable.Columns.Contains(node.MyCode))
                //        dgvPreview.RootTable.Columns[node.MyCode].Caption = node.Description;

                //}
                for (int j = 0; j < _sqlBuilder.SelectedNodes.Count; j++)
                {
                    //if (dgvResult.RootTable.Columns.Contains(_sqlBuilder.SelectedNodes[j].MyCode))
                    //    dgvResult.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].Caption = _sqlBuilder.SelectedNodes[j].Description;

                    if (_sqlBuilder.SelectedNodes[j].Agregate != "")
                    {
                        if (dgvPreview.RootTable.Columns.Contains(_sqlBuilder.SelectedNodes[j].MyCode))
                        {
                            switch (_sqlBuilder.SelectedNodes[j].Agregate)
                            {
                                case "SUM":
                                    dgvPreview.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Sum;
                                    break;
                                case "COUNT":
                                    dgvPreview.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Count;
                                    break;
                                case "AVG":
                                    dgvPreview.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Average;
                                    break;
                                case "MAX":
                                    dgvPreview.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Max;
                                    break;
                                case "MIN":
                                    dgvPreview.RootTable.Columns[_sqlBuilder.SelectedNodes[j].MyCode].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Min;
                                    break;
                            }

                        }
                        else
                        {
                            if (dgvPreview.RootTable.Columns.Contains(_sqlBuilder.SelectedNodes[j].Description))
                            {
                                switch (_sqlBuilder.SelectedNodes[j].Agregate)
                                {
                                    case "SUM":
                                        dgvPreview.RootTable.Columns[_sqlBuilder.SelectedNodes[j].Description].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Sum;
                                        break;
                                    case "COUNT":
                                        dgvPreview.RootTable.Columns[_sqlBuilder.SelectedNodes[j].Description].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Count;
                                        break;
                                    case "AVG":
                                        dgvPreview.RootTable.Columns[_sqlBuilder.SelectedNodes[j].Description].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Average;
                                        break;
                                    case "MAX":
                                        dgvPreview.RootTable.Columns[_sqlBuilder.SelectedNodes[j].Description].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Max;
                                        break;
                                    case "MIN":
                                        dgvPreview.RootTable.Columns[_sqlBuilder.SelectedNodes[j].Description].AggregateFunction = Janus.Windows.GridEX.AggregateFunction.Min;
                                        break;
                                }

                            }
                        }
                    }
                }
                for (int i = 0; i < dgvPreview.RootTable.Columns.Count; i++)
                {
                    //dgvPreview.RootTable.Columns[i].AutoSizeMode = Janus.Windows.GridEX.ColumnAutoSizeMode.DisplayedCellsAndHeader;
                    if (dgvPreview.RootTable.Columns[i].DataTypeCode == TypeCode.Decimal || dgvPreview.RootTable.Columns[i].DataTypeCode == TypeCode.Double)
                    {
                        dgvPreview.RootTable.Columns[i].FormatString = "###,###.##";
                        dgvPreview.RootTable.Columns[i].TotalFormatString = "###,###.##";
                        //dgvPreview.RootTable.Columns[i].AggregateFunction = AggregateFunction.Sum;

                    }

                    // dgvResult.Columns[i].A
                }
                dgvPreview.RootTable.GroupTotals = GroupTotals.Always;
                dgvPreview.RootTable.TotalRow = InheritableBoolean.True;

                //dgvPreview.RootTable.Columns.EndUpdate();
                dgvPreview.AutoSizeColumns();
            }
            catch (System.Exception ex)
            {
                dgvPreview.DataSource = new DataTable();
                lb_Err.Text = ex.Message;
            }
        }

        private void tabItemSQLText_Selected()
        {
            if (CheckError_TabSQL())
            {
                tsMain.SelectedTab = tabItemGeneral;
                txt_sql.Focus();
                //bt_Err.Enabled = true;
            }
        }

        private void
            tabReportViewer_Selected()
        {
            BUS.POSControl posCtr = new BUS.POSControl();
            DTO.POSInfo infPOS = new POSInfo(_user, DB, "Query Designer", "QD", DateTime.Now.ToString("yyyy-MM-dd hh:mm"));
            posCtr.InsertUpdate(infPOS);
            //BUS.CommoControl commo = new CommoControl();
            //string connnectString = commo.CreateConnectString(Properties.Settings.Default.Server
            //             , Properties.Settings.Default.User
            //             , Properties.Settings.Default.Pass
            //             , Properties.Settings.Default.DBName);
            //if (_rpGen == null || _rpGen.IsClose())
            _rpGen = new ReportGenerator(_sqlBuilder, txtqd_id.Text, txt_sql.Text, _strConnectDes, __templatePath, __reportPath);
            LIST_TEMPLATEControl tmpCtr = new LIST_TEMPLATEControl();
            LIST_TEMPLATEInfo tempInfo = tmpCtr.Get(_dtb, txtqd_id.Text, ref sErr);
            if (tempInfo.Code != "")
            {
                _rpGen.STemp = tempInfo.Data;
                _rpGen.LengthTemp = tempInfo.Length;
            }
            //ExcelFile test = null;
            XlsFile xls = null;
            try
            {
                xls = new XlsFile();
                if (File.Exists(__templatePath + "\\NODATA.xls"))
                    xls.Open(__templatePath + "\\NODATA.xls");
                else
                {
                    if (File.Exists(_appPath + "\\NODATA.xls"))
                    {
                        File.Copy(_appPath + "\\NODATA.xls", __templatePath + "\\NODATA.xls");
                        xls.Open(__templatePath + "\\NODATA.xls");
                    }
                }

                _rpGen.Name = txtdesr.Text;
                if (CheckChange(_xlsFile, __templatePath, txtqd_id.Text))
                {
                    _xlsFile = _rpGen.CreateReport();

                    xls.Recalc();
                }
                cbAntiAlias.SelectedIndex = 0;
                //if (File.Exists(reportPath + txtqd_id.Text.Trim() + ".xls"))
                //{
                //    xls.Open(reportPath + txtqd_id.Text.Trim() + ".xls");

                //flexCelImgExport1.Workbook = test;
                flexCelImgExport1.Workbook = _xlsFile;
                flexCelPreview1.InvalidatePreview();
                //}
            }
            catch (Exception ex)
            {

                flexCelImgExport1.Workbook = xls;
                flexCelPreview1.InvalidatePreview();
                lb_Err.Text = ex.Message;
            }

        }

        private void tabItemGeneral_Selected()
        {
            _xlsFile = null;
        }



        private void Form_QD_KeyUp(object sender, KeyEventArgs e)
        {
            //TopMost = false;
        }




        private void twSchema_ItemDrag(object sender, ItemDragEventArgs e)
        {

            Node _node = Node.EmptyNode();
            if (e.Item != null)
            {
                _node = (Node)((TreeNode)e.Item).Tag;
                if (_node.FType != "S")
                    twSchema.DoDragDrop(_node, DragDropEffects.Copy);
            }
        }

        private void dgvFilter_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            QueryBuilder.Node node = ((QueryBuilder.Filter)dgvFilter.Rows[e.RowIndex].DataBoundItem).Node;

            if (node.MyCode.Length > 2 && node.MyCode.Substring(0, 2) != "__")
            {
                if (node.FType == "SDN" || node.FType == "D")
                {
                    frmDateFilterSelect frm = new frmDateFilterSelect(node.FType);
                    frm.FilterFrom = dgvFilter.Rows[e.RowIndex].Cells["FilterFrom"].Value.ToString();
                    frm.FilterTo = dgvFilter.Rows[e.RowIndex].Cells["FilterTo"].Value.ToString();
                    if (frm.ShowDialog() == DialogResult.OK)
                    {
                        dgvFilter.Rows[e.RowIndex].Cells["FilterFrom"].Value = frm.FilterFrom;
                        dgvFilter.Rows[e.RowIndex].Cells["FilterTo"].Value = frm.FilterTo;
                    }
                }
                else
                {
                    frmFilterSelect frm = new frmFilterSelect(_strConnectDes, _sqlBuilder, e.RowIndex);
                    frm.FilterFrom = dgvFilter.Rows[e.RowIndex].Cells["FilterFrom"].Value.ToString();
                    frm.FilterTo = dgvFilter.Rows[e.RowIndex].Cells["FilterTo"].Value.ToString();
                    if (frm.ShowDialog() == DialogResult.OK)
                    {
                        dgvFilter.Rows[e.RowIndex].Cells["FilterFrom"].Value = frm.FilterFrom;
                        dgvFilter.Rows[e.RowIndex].Cells["FilterTo"].Value = frm.FilterTo;
                    }
                }
            }
            else if (node.MyCode.Substring(0, 2) == "__")
            {
                if (node.FType == "SDN" || node.FType == "D")
                {
                    frmDateFilterSelect frm = new frmDateFilterSelect(node.FType);
                    frm.FilterFrom = dgvFilter.Rows[e.RowIndex].Cells["FilterFrom"].Value.ToString();
                    frm.FilterTo = dgvFilter.Rows[e.RowIndex].Cells["FilterTo"].Value.ToString();
                    if (frm.ShowDialog() == DialogResult.OK)
                    {
                        dgvFilter.Rows[e.RowIndex].Cells["FilterFrom"].Value = frm.FilterFrom;
                        dgvFilter.Rows[e.RowIndex].Cells["FilterTo"].Value = frm.FilterTo;
                    }
                }
            }
        }

        private void dgvFilter_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                try
                {
                    //SqlBuilder_Change();
                    _xlsFile = null;
                    QueryBuilder.Filter x = dgvFilter.Rows[e.RowIndex].DataBoundItem as QueryBuilder.Filter;
                    string value = dgvFilter.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null ? "" : dgvFilter.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    if (x != null)
                    {
                        if ((x.Node.FType == "D" || x.Node.FType == "SDN") && Regex.IsMatch(value, _datePaterm))
                        {
                            string yearnow = DateTime.Now.Year.ToString("0000");
                            string[] arr = value.Split('.');

                            if (arr.Length == 3)
                            {
                                if (arr[2].Length < 4)
                                    arr[2] = yearnow.Substring(0, 4 - arr[2].Length) + arr[2];
                                if (x.Node.FType == "D")
                                {
                                    value = arr[2] + "-" + arr[1] + "-" + arr[0];
                                }
                                else
                                    value = arr[2] + arr[1] + arr[0];
                            }
                            else
                            {
                                arr = value.Split('-');

                                if (arr.Length == 3)
                                {
                                    if (arr[2].Length < 4)
                                        arr[2] = yearnow.Substring(0, 4 - arr[2].Length) + arr[2];
                                    if (x.Node.FType == "D")
                                    {
                                        value = arr[2] + "-" + arr[1] + "-" + arr[0];
                                    }
                                    else
                                        value = arr[2] + arr[1] + arr[0];
                                }
                                else
                                {
                                    arr = value.Split('/');

                                    if (arr.Length == 3)
                                    {
                                        if (arr[2].Length < 4)
                                            arr[2] = yearnow.Substring(0, 4 - arr[2].Length) + arr[2];
                                        if (x.Node.FType == "D")
                                        {
                                            value = arr[2] + "-" + arr[1] + "-" + arr[0];
                                        }
                                        else
                                            value = arr[2] + arr[1] + arr[0];
                                    }
                                }
                                dgvFilter.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = value;
                            }
                        }
                        else
                        {
                            if (dgvFilter.Columns[e.ColumnIndex].DataPropertyName == "FilterFrom")
                                x.FilterFrom = x.ValueFrom = value;
                            else if (dgvFilter.Columns[e.ColumnIndex].DataPropertyName == "FilterTo")
                                x.FilterTo = x.ValueTo = value;
                        }
                    }
                }
                catch { }
            }
        }

        private void dgvFilter_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void dgvSelectNodes_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.Row.DataBoundItem is QueryBuilder.Node)
            //{
            //    QueryBuilder.Node x = (QueryBuilder.Node)e.Row.DataBoundItem;
            //    FrmSubField subFDlg = new FrmSubField();
            //    subFDlg.Description = x.Description;
            //    int from = Convert.ToInt32(x.FTypeFull.Substring(5, 2));
            //    int length = Convert.ToInt32(x.FTypeFull.Substring(7, 2));
            //    subFDlg.Index = from;
            //    subFDlg.Length = length;
            //    if (subFDlg.ShowDialog() == DialogResult.OK)
            //    {
            //        x.Description = subFDlg.Description;
            //        x.FTypeFull = x.FType + subFDlg.Index.ToString("00") + subFDlg.Length.ToString("00");
            //    }
            //}
        }

        private void dgvSelectNodes_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

        }



        private void btnConnect_Click(object sender, EventArgs e)
        {

        }

        private void tsMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tsMain.SelectedTab == tabItemGeneral)
                tabItemGeneral_Selected();
            else if (tsMain.SelectedTab == tabItemPreview)
                tabItemPreview_Selected();
            else if (tsMain.SelectedTab == tabItemQD)
                tabItemQD_Selected();
            else if (tsMain.SelectedTab == tabItemSQLText)
                tabItemSQLText_Selected();
            else if (tsMain.SelectedTab == tabReportViewer)
                tabReportViewer_Selected();
            else if (tsMain.SelectedTab == tabChart)
                tabChart_Selected();
        }

        private void tabChart_Selected()
        {
            BUS.POSControl posCtr = new BUS.POSControl();
            DTO.POSInfo infPOS = new POSInfo(_user, DB, "Query Designer", "QD", DateTime.Now.ToString("yyyy-MM-dd hh:mm"));
            posCtr.InsertUpdate(infPOS);
            //_xlsFile = null;
            _xlsFile = ShowWeb();



            //return;
            ////////////////////////////////////////////////////

            ShowChart(_xlsFile);
        }

        private void ShowChart(ExcelFile test)
        {
            try
            {
                //if (_DataXML == "")
                _DataXML = GetChartData(test, _DataXML, _propertyChart);
                if (_DataXML != "")
                {
                    _propertyChart.ReadProperty(_DataXML.Substring(_DataXML.IndexOf("<chart ") + 6, _DataXML.IndexOf(">") - (_DataXML.IndexOf("<chart ") + 6)));
                    ShowChart(_propertyChart, _DataXML);
                }
                else
                    webGadget.DocumentText = "";

            }
            catch (Exception ex) { lb_Err.Text = ex.Message; }
        }

        private ExcelFile ShowWeb()
        {
            //XlsFile xls = new XlsFile();
            //BUS.CommoControl commo = new CommoControl();
            //if (_rpGen == null || _rpGen.IsClose())
            _rpGen = new ReportGenerator(_sqlBuilder, txtqd_id.Text, txt_sql.Text, _strConnectDes, __templatePath, __reportPath);
            LIST_TEMPLATEControl tmpCtr = new LIST_TEMPLATEControl();
            LIST_TEMPLATEInfo tempInfo = tmpCtr.Get(_dtb, txtqd_id.Text, ref sErr);
            if (tempInfo.Code != "")
            {
                _rpGen.STemp = tempInfo.Data;
                _rpGen.LengthTemp = tempInfo.Length;
            }
            string url = "";
            try
            {
                _rpGen.Name = txtdesr.Text;
                if (_rpGen.IsClose() || _filehtml == "")
                {
                    if (CheckChange(_xlsFile, __templatePath, txtqd_id.Text))
                        _xlsFile = _rpGen.CreateReport();
                    else _rpGen.XlsFile = _xlsFile;
                    _filehtml = _rpGen.ExportHTMLToPath(__reportPath);

                    url = "file:///" + _filehtml.Replace("\\", "/");
                    wbChart.Url = new Uri(url);
                }
            }
            catch (Exception ex) { lb_Err.Text = ex.Message; wbChart.Url = new Uri("about:"); }
            return _xlsFile;
        }

        private static string PasreTextHTML(Match x)
        {
            string tmp = x.Value;
            string function = Regex.Match(tmp, "\"=TT_XLB_EB(.+)\"").Value.Substring(1, Regex.Match(tmp, "\"=TT_XLB_EB(.+)\"").Value.Length - 2);
            string value = Regex.Match(tmp, ">.*</td>").Value.Substring(1, Regex.Match(tmp, ">.*</td>").Value.Length - 6);
            return "><a style='text-decoration: none' href='tvcqd:" + function + "'>" + value + "</a></td>";
        }

        private string GetChartData(ExcelFile test, string DataXML, clsChartProperty property)
        {
            if (property.SheetChart == "" || property.DataRange == "")
                return "";
            //XlsFile xls = new XlsFile();
            //BUS.CommoControl commo = new CommoControl();
            //string connnectString = commo.CreateConnectString(Properties.Settings.Default.Server
            //             , Properties.Settings.Default.User
            //             , Properties.Settings.Default.Pass
            //             , Properties.Settings.Default.DBName);
            //if (_rpGen == null || _rpGen.IsClose())
            _rpGen = new ReportGenerator(_sqlBuilder, txtqd_id.Text, txt_sql.Text, _strConnectDes, __templatePath, __reportPath);
            //ExcelFile test = null;
            try
            {
                _rpGen.Name = txtdesr.Text;
                test = _rpGen.CreateReport();
                int XF = 0;

                int currentSheet = test.ActiveSheet;
                int indexSheet = test.GetSheetIndex(property.SheetChart);
                test.ActiveSheet = indexSheet;
                TXlsNamedRange range = test.GetNamedRange(property.DataRange, indexSheet); string attribute = "";
                DataXML = "<chart>";
                if (property.CaptionRange != null && property.CaptionRange != "")
                {
                    TXlsNamedRange rangeCaption = test.GetNamedRange(property.CaptionRange, indexSheet);
                    if (rangeCaption != null && rangeCaption.IsOneCell)
                    {
                        string caption = test.GetCellValue(indexSheet, rangeCaption.Top, rangeCaption.Left, ref XF) == null ? "" : test.GetCellValue(indexSheet, rangeCaption.Top, rangeCaption.Left, ref XF).ToString();
                        if (caption.Trim() != "")
                            attribute += " caption='" + XmlEncoder.Decode(caption) + "'";
                    }
                }
                if (property.SubCaptionRange != null && property.SubCaptionRange != "")
                {
                    TXlsNamedRange rangeSubCaption = test.GetNamedRange(property.SubCaptionRange, indexSheet);
                    if (rangeSubCaption != null && rangeSubCaption.IsOneCell)
                    {
                        string subCaption = test.GetCellValue(indexSheet, rangeSubCaption.Top, rangeSubCaption.Left, ref XF) == null ? "" : test.GetCellValue(indexSheet, rangeSubCaption.Top, rangeSubCaption.Left, ref XF).ToString();
                        if (subCaption.Trim() != "")
                            attribute += " subCaption='" + XmlEncoder.Decode(subCaption) + "'";
                    }
                }
                if (range.ColCount == 2)
                {
                    string datalable = test.GetCellValue(indexSheet, range.Top, range.Left, ref XF) == null ? "" : test.GetCellValue(indexSheet, range.Top, range.Left, ref XF).ToString();
                    if (datalable != "")
                        attribute += " xAxisName='" + XmlEncoder.Decode(datalable) + "'";
                    string datavalue = test.GetCellValue(indexSheet, range.Top, range.Left + 1, ref XF) == null ? "" : test.GetCellValue(indexSheet, range.Top, range.Left + 1, ref XF).ToString();
                    if (datavalue != "")
                        attribute += " yAxisName='" + XmlEncoder.Decode(datavalue) + "'";
                    //attribute = datalable + datavalue;


                    for (int i = 1; i < range.RowCount - 1; i++)
                    {
                        //<set label="Item A" value="4" />

                        string comment = "";
                        comment = test.GetComment(range.Top + i, range.Left + 1).ToString();
                        //System.Web.Configuration.Converter convert = new System.Web.Configuration.Converter();
                        comment = XmlEncoder.Encode(comment);
                        //comment = comment.Replace("<", "%3C")
                        //                .Replace(">", "%3E")
                        //                .Replace("/", "%2F")
                        //                .Replace(" ", "%20")
                        //                .Replace("=", "%3D")
                        //                .Replace(":", "%3A")
                        //                .Replace(";", "%3B")
                        //                .Replace("(", "%28")
                        //                .Replace(")", "%29")
                        //                .Replace("\"", "%22")
                        //                .Replace("\\", "%5C");
                        if (comment.Contains("TT_XLB_EB"))
                            comment = " link='tvcqd:" + comment + "'";

                        string name = test.GetCellValue(indexSheet, range.Top + i, range.Left, ref XF) == null ? "" : test.GetCellValue(indexSheet, range.Top + i, range.Left, ref XF).ToString();
                        string value = test.GetCellValue(indexSheet, range.Top + i, range.Left + 1, ref XF) == null ? "" : test.GetCellValue(indexSheet, range.Top + i, range.Left + 1, ref XF).ToString();
                        string temp = "<set label ='" + XmlEncoder.Encode(name) + "'" +
                                            " value='" + XmlEncoder.Encode(value) + "' " + comment + "/>";
                        DataXML += temp;

                    }
                }
                else if (range.ColCount > 2)
                {
                    #region DataLable
                    DataXML += "<categories>";


                    for (int i = 1; i < range.RowCount - 1; i++)
                    {
                        //<category label='Austria' /> 
                        string value = test.GetCellValue(indexSheet, range.Top + i, range.Left, ref XF) == null ? "" : test.GetCellValue(indexSheet, range.Top + i, range.Left, ref XF).ToString();
                        string temp = "<category label ='" + value + "'/>";
                        DataXML += temp;

                    }
                    DataXML += "</categories>";
                    #endregion

                    for (int i = 1; i < range.ColCount; i++)
                    {
                        //<dataset seriesName='1997' color='F6BD0F' showValues='0'>
                        string seriesName = test.GetCellValue(indexSheet, range.Top, range.Left + i, ref XF) == null ? "" : test.GetCellValue(indexSheet, range.Top, range.Left + i, ref XF).ToString();
                        DataXML += "<dataset seriesName='" + XmlEncoder.Encode(seriesName) + "'>";

                        for (int j = 1; j < range.RowCount - 1; j++)
                        {
                            //<set label="Item A" value="4" />

                            string comment = "";
                            comment = test.GetComment(range.Top + j, range.Left + i).ToString();
                            comment = XmlEncoder.Encode(comment);
                            //comment.Replace("<", "%3C")
                            //            .Replace(">", "%3E")
                            //            .Replace("/", "%2F")
                            //            .Replace(" ", "%20")
                            //            .Replace("=", "%3D")
                            //            .Replace(":", "%3A")
                            //            .Replace(";", "%3B")
                            //            .Replace("(", "%28")
                            //            .Replace(")", "%29")
                            //            .Replace("\"", "%22")
                            //            .Replace("\\", "%5C");
                            comment = XmlEncoder.Decode(comment);
                            if (comment.Contains("TT_XLB_EB"))
                                comment = " link='tvcqd:" + comment + "'";

                            string value = test.GetCellValue(indexSheet, range.Top + j, range.Left + i, ref XF) == null ? "" : test.GetCellValue(indexSheet, range.Top + j, range.Left + i, ref XF).ToString();
                            string temp = "<set value='" + XmlEncoder.Decode(value) + "' " + comment + "/>";
                            DataXML += temp;

                        }
                        DataXML += "</dataset>";
                    }
                }
                else
                {
                    test.ActiveSheet = currentSheet;
                    return "";
                }
                test.ActiveSheet = currentSheet;
                DataXML = DataXML.Insert(DataXML.IndexOf("<chart>") + 6, attribute);
                DataXML += "</chart>";
            }
            catch (Exception ex) { lb_Err.Text = ex.Message; }
            return DataXML;
        }

        private void ShowChart(clsChartProperty propertyChart, string DataXML)
        {
            try
            {
                string attribute = "";
                attribute = propertyChart.GetPropertyForChart(attribute);
                DataXML = "<chart" + DataXML.Substring(DataXML.IndexOf(">"));
                DataXML = DataXML.Insert(DataXML.IndexOf("<chart>") + 6, attribute);
                StreamReader reader = new StreamReader(_appPath + "\\charts.htm");
                string content = reader.ReadToEnd();
                reader.Close();
                string PathScript = _appPath + "\\FusionCharts\\FusionCharts.js";
                PathScript = PathScript.Replace("\\", "/");
                string PathChart = _appPath + "\\FusionCharts\\" + propertyChart.ChartName + ".swf";
                PathChart = PathChart.Replace("\\", "/");

                //DataXML = DataXML.Replace("<", "%3C")
                //    .Replace(">", "%3E")
                //    .Replace("/", "%2F")
                //    .Replace(" ", "%20")
                //    .Replace("=", "%3D")
                //    .Replace(":", "%3A")
                //    .Replace(";", "%3B")
                //    .Replace("(", "%28")
                //    .Replace(")", "%29");
                content = content.Replace("<#PathScript>", PathScript)
                    .Replace("<#PathChart>", PathChart)
                    .Replace("<#DataXML>", DataXML)
                    .Replace("<#Width>", "100%")// (webBrowser1.Width - 50).ToString()
                    .Replace("<#Height>", "100%");//(webBrowser1.Height - 50).ToString()
                webGadget.DocumentText = content;
            }
            catch { }
        }

        private void dgvFilter_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(Node)))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void twSchema_NodeMouseHover(object sender, TreeNodeMouseHoverEventArgs e)
        {
            _currentNode = e.Node;
        }

        private void dgvSelectNodes_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(Node)))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
            e.Effect = DragDropEffects.Copy | DragDropEffects.All;
        }

        private void dgvSelectNodes_DragDrop(object sender, DragEventArgs e)
        {
            _sqlBuilder.SelectedNodes.Add((Node)e.Data.GetData(typeof(Node)));
            _xlsFile = null;
            //SqlBuilder_Change();
        }

        private void dgvFilter_DragDrop(object sender, DragEventArgs e)
        {
            _sqlBuilder.Filters.Add(new QueryBuilder.Filter((Node)e.Data.GetData(typeof(Node))));
            _xlsFile = null;
            //SqlBuilder_Change();
        }

        private void txtdatabase_TextChanged(object sender, EventArgs e)
        {

        }



        private void twSchema_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TreeNode tmpNode = twSchema.SelectedNode;

                if (tmpNode != null && dgvSelectNodes.AllowDrop == true)
                {
                    bool flag = true;
                    //string[] arrNode = tmpNode.Tag.ToString().Split(';');
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

        private void txtCommand_KeyUp(object sender, KeyEventArgs e)
        {
            if (DB != "")
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (txtCommand.Text == "QDADD" && _flagQDADD)
                    {
                        try
                        {
                            frmQDADD frm = new frmQDADD(DB, _user);
                            frm._config = _config;
                            frm.Show();
                        }
                        catch (Exception ex)
                        {
                            lb_Err.Text = ex.Message;
                        }
                    }
                }
            }
            else lb_Err.Text = "Required Database!";
        }

        private void btnChartPropety_Click(object sender, EventArgs e)
        {

            frmChartPro frm = new frmChartPro(_propertyChart);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                _propertyChart = frm.ReturnProperty;
                //string result = _propertyChart.GetProperty();
                //_propertyChart.ReadProperty(result);
                ShowChart(_xlsFile);
            }
        }
        public static string StringToBase64(string str)
        {
            byte[] b = System.Text.Encoding.UTF8.GetBytes(str);
            string b64 = Convert.ToBase64String(b);
            return b64;
        }


        public static string Base64ToString(string b64)
        {
            byte[] b = Convert.FromBase64String(b64);
            return (System.Text.Encoding.UTF8.GetString(b));
        }

        private void wbChart_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            string schema = "";
            string function = "";
            try
            {
                //ContextMenuStrip ex = wbChart.ContextMenuStrip;
                schema = e.Url.Scheme;
                string myurl = e.Url.OriginalString.Replace("res://ieframe.dll/unknownprotocol.htm#", "");
                if (myurl.Length > 9 && myurl.Substring(0, 9) == "tavico://")
                {
                    string[] arry = myurl.Replace("tavico://", "").Split('?');
                    if (arry.Length == 2)
                    {
                        if (arry[0].ToUpper().Contains("ANALYSERUI"))
                        {
                            try
                            {
                                //function = XmlEncoder.Decode(function);
                                QDAddinDrillDown frm = new QDAddinDrillDown("A1", null, Base64ToString(arry[1].Replace("tag=", "")), _strConnectDes);
                                frm.Config = _config;
                                frm.Show(this);
                            }
                            catch (Exception ex)
                            {
                                lb_Err.Text = ex.Message;
                            }
                        }
                        e.Cancel = true;
                    }
                    e.Cancel = true;
                }


                function = e.Url.AbsolutePath.Replace("#tvcqd:", "");
                //Replace("%3C", "<")
                //                            .Replace("%3E", ">")
                //                            .Replace("%2F", "/")
                //                            .Replace("%20", " ")
                //                            .Replace("%3D", "=")
                //                            .Replace("%3A", ":")
                //                            .Replace("%3B", ";")
                //                            .Replace("%28", "(")
                //                            .Replace("%29", ")")
                //                            .Replace("%22", "\"")
                //                            .Replace("%5C", "\\")
                //                            .Replace("#tvcqd:", "")
                //                            .Replace("%7B", "{")
                //                            .Replace("%7D", "}");

                lb_Err.Text = schema;
                //MessageBox.Show("aa");
            }
            catch (Exception ex)
            {
                lb_Err.Text = ex.Message;
                e.Cancel = true;
            }




        }
        private bool DoSetup(PrintDocument doc)
        {
            printDialog1.Document = doc;
            printDialog1.PrinterSettings.DefaultPageSettings.PaperSize = doc.DefaultPageSettings.PaperSize;
            //printDialog1.PrinterSettings.
            bool Result = printDialog1.ShowDialog() == DialogResult.OK;
            //printDialog1.PrinterSettings.PaperSizes = doc.PrinterSettings.pga
            //Landscape.Checked = flexCelPrintDocument1.DefaultPageSettings.Landscape;
            return Result;
        }
        private void btnPrint_Click(object sender, EventArgs e)
        {

            PrintPreviewDialog preview = new PrintPreviewDialog();
            if (!DoSetup(gridEXPrintDocument1)) return;
            preview.Document = gridEXPrintDocument1;
            preview.Show();
            //frmPrintPreview frm = new frmPrintPreview();
            //frm.Show(gridEXPrintDocument1, this);
            //try
            //{
            //    gridEXPrintDocument1.PrepareDocument();
            //    gridEXPrintDocument1.Print();
            //}
            //catch (Exception ex)
            //{
            //    lb_Err.Text = ex.Message;
            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileStream sr = new FileStream(saveFileDialog1.FileName, FileMode.OpenOrCreate, FileAccess.Write);
                gridEXExporter1.Export(sr);
                sr.Close();
                if (MessageBox.Show("Do you want to open this document?", "Message", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    Process.Start(saveFileDialog1.FileName);
                }
            }
        }

        private void expandedAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dgvPreview.ExpandGroups();
        }

        private void collapsedAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dgvPreview.CollapseGroups();
        }
        frmPOD _frmPOD;
        frmPOG _frmPOG;
        frmPOP _frmPOP;
        frmDA _frmDA;
        private void operatorDefinitionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_frmPOD == null || !this.Controls.Contains(_frmPOD))
            {
                _frmPOD = new frmPOD();
                _frmPOD.Show(this);
            }
            else
            {
                _frmPOD.Focus();
                _frmPOD.BringToFront();
            }
        }

        private void operatorGroupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_frmPOG == null || !this.Controls.Contains(_frmPOG))
            {
                _frmPOG = new frmPOG();
                _frmPOG.Show(this);
            }
            else
            {
                _frmPOG.Focus();
                _frmPOG.BringToFront();
            }
        }

        private void permissionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_frmPOP == null || !this.Controls.Contains(_frmPOP))
            {
                _frmPOP = new frmPOP();
                _frmPOP.Show(this);
            }
            else
            {
                _frmPOP.Focus();
                _frmPOP.BringToFront();
            }
        }

        private void dataAccessGroupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_frmDA == null || !this.Controls.Contains(_frmDA))
            {
                _frmDA = new frmDA();
                _frmDA.Show(this);
            }
            else
            {
                _frmDA.Focus();
                _frmDA.BringToFront();
            }
        }

        private void txtANAL_Q2_TextChanged(object sender, EventArgs e)
        {
            LIST_DAControl ctr = new LIST_DAControl();
            lbgroup.Text = ctr.Get(txtANAL_Q2.Text, ref sErr).NAME;
        }

        private void bt_group_Click(object sender, EventArgs e)
        {
            frmDAView frm = new frmDAView();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                txtANAL_Q2.Text = frm.Code;
            }
        }

        private void btnPrintReport_Click(object sender, EventArgs e)
        {
            try
            {
                if (!LoadPreferences()) return;
                if (!DoSetup(flexCelPrintDocument1)) return;
                flexCelPrintDocument1.Print();
            }
            catch (Exception ex)
            {
                lb_Err.Text = ex.Message;
            }
            //try
            //{
            //    frmFlexPrintPreview frmPrint = new frmFlexPrintPreview(flexCelImgExport1.Workbook);
            //    frmPrint.Show(this);
            //}
            //catch (Exception ex)
            //{
            //    lb_Err.Text = ex.Message;
            //}
        }
        FlexCelPrintDocument flexCelPrintDocument1 = new FlexCelPrintDocument();
        private bool LoadPreferences()
        {
            try
            {
                flexCelPrintDocument1.Workbook = flexCelImgExport1.Workbook;
                ExcelFile Xls = flexCelPrintDocument1.Workbook;
                Xls.PrintHeadings = false;
                Xls.PrintGridLines = false;
                //Xls.PrintPaperSize = TPaperSize.

                flexCelPrintDocument1.DefaultPageSettings.PaperSize = new PaperSize(Xls.PrintPaperDimensions.PaperName, Convert.ToInt32(Xls.PrintPaperDimensions.Width), Convert.ToInt32(Xls.PrintPaperDimensions.Height));
                //flexCelPrintDocument1.PrintPa
                flexCelPrintDocument1.DefaultPageSettings.Landscape = (Xls.PrintOptions & TPrintOptions.Orientation) == 0;
                return true;
            }
            catch (Exception ex) { throw ex; }
        }

        private void changePasswordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmChangePass frm = new frmChangePass(_user);
            frm.ShowDialog();
        }

        private void dgvPreview_DoubleClick(object sender, EventArgs e)
        {
            GridArea clickArea = dgvPreview.HitTest();
            switch (clickArea)
            {
                case GridArea.GroupByBox:
                case GridArea.GroupByBoxInfoText:
                    this.ShowGroupByDialog();
                    break;
                //case GridArea.Cell:
                //case GridArea.PreviewRow:
                //case GridArea.CardCaption:
                //    this.Edit();

                //    break;
            }
        }
        public void ShowGroupByDialog()
        {
            //frmGroupBy frm = new frmGroupBy();
            //frm.ShowDialog(dgvPreview, this.ParentForm);
            //frm.Dispose();
        }
        public void ShowFieldsDialog()
        {
            frmShowFields frm = new frmShowFields();
            frm.ShowDialog(this.dgvPreview, this.ParentForm);
            frm.Dispose();
        }
        public void ShowFormatViewDialog()
        {
            frmFormatView frm = new frmFormatView();
            frm.ShowDialog(this.dgvPreview, this.ParentForm);
            frm.Dispose();

        }
        public void ShowFormatConditionsDialog()
        {
            //frmFormatConditions frm = new frmFormatConditions();
            //frm.ShowDialog(this.dgvPreview, this.ParentForm);
            //frm.Dispose();
        }
        public void ShowFilterDialog()
        {
            //frmFilter frm = new frmFilter();
            //frm.ShowDialog(this.dgvPreview, this.ParentForm);
            //frm.Dispose();
        }
        private void customGroupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.ShowGroupByDialog();
        }

        private frmFieldChooser fieldChoser;
        private void OnFieldChooserCommand()
        {
            if (fieldChoser == null || fieldChoser.IsDisposed)
            {
                fieldChoser = new frmFieldChooser();
                fieldChoser.Show(this.dgvPreview, this);
            }
            else
            {
                fieldChoser.Close();
                fieldChoser.Dispose();
                fieldChoser = null;
            }
        }
        private void showFieldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.ShowFieldsDialog();
        }

        private void formatViewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.ShowFormatViewDialog();
        }

        private void formatConditionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.ShowFormatConditionsDialog();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.ShowFilterDialog();
        }



        private void filterBindingSource_DataSourceChanged(object sender, EventArgs e)
        {
            //SqlBuilder_Change();
        }

        private void dgvSelectNodes_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //SqlBuilder_Change();
        }

        private void dgvSelectNodes_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            _xlsFile = null;
            //SqlBuilder_Change();
        }
        int _zoom = 100;
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (_zoom > 10)
            {
                _zoom -= 10;
                webGadget.Document.Body.SetAttribute("zoom", _zoom + "%");
                //webGadget.Document.Body.SetAttribute("style", "zoom:5;");

            }
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {

            _zoom += 10;
            webGadget.Document.Body.SetAttribute("zoom", _zoom + "%");

        }

        private void importDefinitionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmImportDefinition frm = new frmImportDefinition();
            frm.DTB = DB;
            frm._config = _config;
            frm.Show(this);
        }

        private void importToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            frmImport frm = new frmImport(DB);
            frm.DTB = DB;
            frm.Show(this);
        }

        private void txt_database_Click(object sender, EventArgs e)
        {

        }

        private void txtdatabase_Validating(object sender, CancelEventArgs e)
        {

        }

        private void txtdatabase_Validated(object sender, EventArgs e)
        {
            BUS.DBAControl dbaCtr = new DBAControl();
            DTO.DBAInfo dbaInf = dbaCtr.Get(DB, ref sErr);
            //txt_database.Text = dbaInf.DESCRIPTION;
            ResetForm();
        }

        private void txtqd_id_Validated(object sender, EventArgs e)
        {
            BUS.LIST_QDControl ctr = new LIST_QDControl();
            DTO.LIST_QDInfo inf = ctr.Get_LIST_QD(DB, txtqd_id.Text, ref sErr);
            if (inf.QD_ID != "")
            {
                LoadQD(inf);
                _processStatus = "V";
            }
        }

        private void changeDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmChangeDB frm = new frmChangeDB(DB);
            frm.User = _user;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                //if (DB != frm.)
                //{
                //    DB = DB;
                //txtdatabase.Focus();
                ResetForm();
                _sqlBuilder.Database = DB;
                Text = "Query Desinger for WinForm - " + _user + "@" + DB;
                //}
            }
        }

        private void connectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmSystem frm = new FrmSystem();

            frm.THEME = THEME;
            if (frm.ShowDialog() == DialogResult.OK)
                LoadConfig("");
        }

        private void txtdatasource_Validated(object sender, EventArgs e)
        {
            //try
            //{
            //    _sqlBuilder.Database = DB;
            //    BindingList<Node> list = SchemaDefinition.GetDecorateTableByCode(txtdatasource.Text.Trim(), _sqlBuilder.Database);
            //    twSchema = TreeViewLoader.LoadTree(ref twSchema, list, txtdatasource.Text.Trim(), "");
            //    BUS.LIST_QD_SCHEMAControl ctr = new LIST_QD_SCHEMAControl();
            //    DTO.LIST_QD_SCHEMAInfo inf = ctr.Get(DB, txtdatasource.Text, ref sErr);
            //    string key = inf.DEFAULT_CONN;
            //    _strConnectDes = Form_QD._config.GetConnection(ref key, "AP");
            //}
            //catch (Exception ex) { lb_Err.Text = ex.Message; }
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            frmDBDef frm = new frmDBDef();
            frm.ShowDialog();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void taskToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmTask frm = new frmTask(DB, _user);
            frm.Show();
        }

        private void dgvPreview_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.C && e.Modifiers == Keys.Control)
            {
                if (dgvPreview.Row >= 0)
                {
                    Clipboard.SetDataObject(dgvPreview.GetRow(dgvPreview.Row).Cells[dgvPreview.Col].Value.ToString());
                }
            }
        }

        private void opeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmPOS frm = new frmPOS(_user);
            frm.ShowDialog();
        }

        private void Form_QD_FormClosed(object sender, FormClosedEventArgs e)
        {
            BUS.POSControl posCtr = new POSControl();
            posCtr.Delete(_user);
        }

        private void btnTmp_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel file(*.xls,*.xlsx)|*.xls*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string filename = ofd.FileName;
                txtTmp.Text = filename;
            }
        }

        private void btnTmpClear_Click(object sender, EventArgs e)
        {

            LIST_TEMPLATEControl ctr = new LIST_TEMPLATEControl();
            if (ctr.IsExist(_dtb, txtqd_id.Text)
                && MessageBox.Show("Do you want to delete this Template in database?", "Delete Message", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                DeleteTemplateToDB();
            }
            txtTmp.Text = "";
        }

        private void DeleteTemplateToDB()
        {
            LIST_TEMPLATEControl ctr = new LIST_TEMPLATEControl();
            sErr = ctr.Delete(_dtb, txtqd_id.Text);

        }

        private void UpdateTemplateToDB()
        {
            if (File.Exists(txtTmp.Text))
            {
                LIST_TEMPLATEControl ctr = new LIST_TEMPLATEControl();
                LIST_TEMPLATEInfo info = new LIST_TEMPLATEInfo();
                info.DTB = _dtb;
                info.Code = txtqd_id.Text;
                using (FileStream fs = new FileStream(txtTmp.Text, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    info.Data = new byte[fs.Length];
                    info.Length = (int)fs.Length;
                    fs.Read(info.Data, 0, (int)fs.Length);
                }
                sErr = ctr.InsertUpdate(info);
                if (sErr == "")
                    txtTmp.Text = info.Code;
            }
        }






    }

}
