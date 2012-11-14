using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;
using AddinExpress.MSO;
using dCube;
using System.IO;
using BUS;
using QueryBuilder;
using System.Data;

namespace dCubeXL
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("6214C6CA-20D4-4219-9E14-535B88699FD0"), ProgId("dCubeXL.AddinModule")]
    public class AddinModule : AddinExpress.MSO.ADXAddinModule
    {

        QDConfig _config = new QDConfig();
        frmLoginEx frmLog;
        QDAddIn frm;
        QDAddinDrillDown frmdrill;
        FrmSystem frmConnect;
        object type = Type.Missing;
        Excel.Range _xlsCell;
        public static string _strConnect = "";
        public static string _strConnectDes = "";
        string __templatePath = "";
        string __reportPath = "";

        public static string __documentDirectory = string.Empty;
        string _sErr = "";
        string _address = "A1";
        string _conn_ID = "";
        private ADXRibbonTab adxRibbonTab1;
        private ADXRibbonGroup adxRibbonGroup1;
        private ADXRibbonButton btnRSetting;
        private ADXRibbonButton btnRDesign;
        private ADXRibbonButton btnRComment;
        private ADXRibbonButton btnRAnalysis;
        private ADXCommandBarButton btnLogin;
        private ADXRibbonButton adxLogin;
        string _appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Replace("file:\\", "");
        private ADXRibbonButton btnView;
        private ADXRibbonButton adxRibbonButton1;
        private ADXRibbonButton adxRibbonButton2;
        private ADXRibbonButton adxRibbonButton3;
        private ADXCommandBarButton btnViewCol;
        string _user = "";
        private void InitDocument()
        {
            //string filename = _appPath + "\\Configuration\\xmlConnect.xml";
            __documentDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\" + Form_QD.DocumentFolder;
            if (!Directory.Exists(__documentDirectory))
            {
                Directory.CreateDirectory(__documentDirectory);
            }
            string configureDirectory = __documentDirectory + "\\Configuration";
            if (!Directory.Exists(configureDirectory))
            {
                Directory.CreateDirectory(configureDirectory);
            }
            //string connectionFile = configureDirectory + "\\xmlConnect.xml";
            //if (!File.Exists(connectionFile))
            //{
            //    File.Copy(filename, connectionFile);
            //}
            string logFolder = configureDirectory + "\\Log";
            if (!Directory.Exists(logFolder))
            {
                Directory.CreateDirectory(logFolder);
            }
            //ValidateLicense(configureDirectory + "\\license.bin");
        }


        private void LoadConfig(string strAP)
        {
            if (File.Exists(__documentDirectory + "\\Configuration\\QDConfig.tvc"))
            {
                _config.LoadConfig(__documentDirectory + "\\Configuration\\QDConfig.tvc");
                string key = "";
                _strConnect = _config.GetConnection(ref key, "QD");
                QueryBuilder.SQLBuilder.SetConnection(_strConnect);
                CommonControl.SetConnection(_strConnect);
                //MessageBox.Show(_strConnect);
                _strConnectDes = _config.GetConnection(ref strAP, "AP");
                _conn_ID = strAP;



                if (_config.DIR.Count > 0)
                {
                    __templatePath = _config.DIR[0]["TMP"].ToString();
                    __reportPath = _config.DIR[0]["RPT"].ToString();
                }
                if (_config.SYS.Count > 0)

                    ReportGenerator.User2007 = (bool)_config.SYS[0][_config.SYS.USE2007Column];
            }
        }

        public AddinModule()
        {
            Application.EnableVisualStyles();
            InitializeComponent();

            try
            {
                InitDocument();
                LoadConfig("");
            }
            catch (Exception ex)
            {
                BUS.CommonControl.AddLog("ErroLog", __documentDirectory + "\\Log", "Addin : " + ex.Message + "\n\t" + ex.Source + "\n\t" + ex.StackTrace);
            }

            // Please add any initialization code to the AddinInitialize event handler
        }

        private ADXCommandBar QDCommandBar;
        private ADXCommandBarButton btnSetting;
        private ImageList ilMain;
        private ADXCommandBarButton btnDesign;
        private ADXCommandBarButton btnComment;
        private ADXCommandBarButton btnAnalysis;
        private ADXExcelAppEvents adxExcelEvents;

        #region Component Designer generated code
        /// <summary>
        /// Required by designer
        /// </summary>
        private System.ComponentModel.IContainer components;

        /// <summary>
        /// Required by designer support - do not modify
        /// the following method
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinModule));
            this.QDCommandBar = new AddinExpress.MSO.ADXCommandBar(this.components);
            this.btnLogin = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.ilMain = new System.Windows.Forms.ImageList(this.components);
            this.btnSetting = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.btnDesign = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.btnComment = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.btnAnalysis = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxExcelEvents = new AddinExpress.MSO.ADXExcelAppEvents(this.components);
            this.adxRibbonTab1 = new AddinExpress.MSO.ADXRibbonTab(this.components);
            this.adxRibbonGroup1 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.adxLogin = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.btnRSetting = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.btnRDesign = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.btnRComment = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.btnRAnalysis = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonButton1 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonButton2 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonButton3 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.btnView = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.btnViewCol = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            // 
            // QDCommandBar
            // 
            this.QDCommandBar.CommandBarName = "TVC-QD";
            this.QDCommandBar.CommandBarTag = "1ca346fe-f8af-48ac-bb53-e85444232d09";
            this.QDCommandBar.Controls.Add(this.btnLogin);
            this.QDCommandBar.Controls.Add(this.btnSetting);
            this.QDCommandBar.Controls.Add(this.btnDesign);
            this.QDCommandBar.Controls.Add(this.btnComment);
            this.QDCommandBar.Controls.Add(this.btnAnalysis);
            this.QDCommandBar.Controls.Add(this.btnViewCol);
            this.QDCommandBar.Description = "Tavicosoft Addin";
            this.QDCommandBar.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;
            this.QDCommandBar.UpdateCounter = 18;
            // 
            // btnLogin
            // 
            this.btnLogin.Caption = "Login";
            this.btnLogin.ControlTag = "3e334944-80b5-4d31-beab-e482e1ea4463";
            this.btnLogin.Image = 0;
            this.btnLogin.ImageList = this.ilMain;
            this.btnLogin.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnLogin.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.btnLogin.UpdateCounter = 8;
            this.btnLogin.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.btnLogin_Click);
            // 
            // ilMain
            // 
            this.ilMain.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ilMain.ImageStream")));
            this.ilMain.TransparentColor = System.Drawing.Color.Transparent;
            this.ilMain.Images.SetKeyName(0, "1334119670_login.png");
            this.ilMain.Images.SetKeyName(1, "1334119724_setting.png");
            this.ilMain.Images.SetKeyName(2, "1334120016_pie_chart.png");
            this.ilMain.Images.SetKeyName(3, "comment.png");
            this.ilMain.Images.SetKeyName(4, "design.png");
            this.ilMain.Images.SetKeyName(5, "columns.png");
            // 
            // btnSetting
            // 
            this.btnSetting.Caption = "Setting";
            this.btnSetting.ControlTag = "045b2d4c-d402-4971-951d-e680641c70e7";
            this.btnSetting.Image = 1;
            this.btnSetting.ImageList = this.ilMain;
            this.btnSetting.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnSetting.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndWrapCaption;
            this.btnSetting.UpdateCounter = 21;
            this.btnSetting.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.btnSetting_Click);
            // 
            // btnDesign
            // 
            this.btnDesign.Caption = "Design";
            this.btnDesign.ControlTag = "65808cf7-fbba-400e-954c-7ae93e214ae9";
            this.btnDesign.Image = 4;
            this.btnDesign.ImageList = this.ilMain;
            this.btnDesign.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnDesign.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndWrapCaption;
            this.btnDesign.UpdateCounter = 12;
            this.btnDesign.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.btnDesign_Click);
            // 
            // btnComment
            // 
            this.btnComment.Caption = "Comment";
            this.btnComment.ControlTag = "628d01ab-07cd-4377-b5d7-851d931ec005";
            this.btnComment.Image = 3;
            this.btnComment.ImageList = this.ilMain;
            this.btnComment.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnComment.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.btnComment.UpdateCounter = 8;
            this.btnComment.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.btnComment_Click);
            // 
            // btnAnalysis
            // 
            this.btnAnalysis.Caption = "Analysis";
            this.btnAnalysis.ControlTag = "a44bc5bb-d924-44ab-8827-7974ce11ee49";
            this.btnAnalysis.Image = 2;
            this.btnAnalysis.ImageList = this.ilMain;
            this.btnAnalysis.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnAnalysis.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.btnAnalysis.UpdateCounter = 9;
            this.btnAnalysis.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.btnAnalysis_Click);
            // 
            // adxExcelEvents
            // 
            this.adxExcelEvents.SheetSelectionChange += new AddinExpress.MSO.ADXExcelSheet_EventHandler(this.adxExcelEvents_SheetSelectionChange);
            this.adxExcelEvents.SheetBeforeDoubleClick += new AddinExpress.MSO.ADXExcelSheetBefore_EventHandler(this.adxExcelEvents_SheetBeforeDoubleClick);
            // 
            // adxRibbonTab1
            // 
            this.adxRibbonTab1.Caption = "TVC-QD";
            this.adxRibbonTab1.Controls.Add(this.adxRibbonGroup1);
            this.adxRibbonTab1.Id = "adxRibbonTab_72d9a8860ea24ea08ec7ced4085f0b35";
            this.adxRibbonTab1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonGroup1
            // 
            this.adxRibbonGroup1.Caption = "Command";
            this.adxRibbonGroup1.Controls.Add(this.adxLogin);
            this.adxRibbonGroup1.Controls.Add(this.btnRSetting);
            this.adxRibbonGroup1.Controls.Add(this.btnRDesign);
            this.adxRibbonGroup1.Controls.Add(this.btnRComment);
            this.adxRibbonGroup1.Controls.Add(this.btnRAnalysis);
            this.adxRibbonGroup1.Controls.Add(this.btnView);
            this.adxRibbonGroup1.Id = "adxRibbonGroup_81754909ad5e45eb9ea84f07866d115e";
            this.adxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxLogin
            // 
            this.adxLogin.Caption = "Login";
            this.adxLogin.Id = "adxRibbonButton_0505baae132e43b499c301bb7bf26d13";
            this.adxLogin.Image = 0;
            this.adxLogin.ImageList = this.ilMain;
            this.adxLogin.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxLogin.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxLogin.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxLogin.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.btnRLogin_OnClick);
            // 
            // btnRSetting
            // 
            this.btnRSetting.Caption = "Setting";
            this.btnRSetting.Id = "adxRibbonButton_78c501bd03b741d88605e8a2439aaa50";
            this.btnRSetting.Image = 1;
            this.btnRSetting.ImageList = this.ilMain;
            this.btnRSetting.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnRSetting.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.btnRSetting.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.btnRSetting.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.btnRSetting_OnClick);
            // 
            // btnRDesign
            // 
            this.btnRDesign.Caption = "Design";
            this.btnRDesign.Id = "adxRibbonButton_adea41b72bc74ee2afa29109733273ee";
            this.btnRDesign.Image = 4;
            this.btnRDesign.ImageList = this.ilMain;
            this.btnRDesign.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnRDesign.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.btnRDesign.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.btnRDesign.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.btnRDesign_OnClick);
            // 
            // btnRComment
            // 
            this.btnRComment.Caption = "Comment";
            this.btnRComment.Id = "adxRibbonButton_38eb5c511e4544c5adb5875c2a99edf0";
            this.btnRComment.Image = 3;
            this.btnRComment.ImageList = this.ilMain;
            this.btnRComment.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnRComment.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.btnRComment.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.btnRComment.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.btnRComment_OnClick);
            // 
            // btnRAnalysis
            // 
            this.btnRAnalysis.Caption = "Analysis";
            this.btnRAnalysis.Id = "adxRibbonButton_d1d6b2c5c61b4c97aaae713ea9547164";
            this.btnRAnalysis.Image = 2;
            this.btnRAnalysis.ImageList = this.ilMain;
            this.btnRAnalysis.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnRAnalysis.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.btnRAnalysis.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.btnRAnalysis.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.btnRAnalysis_OnClick);
            // 
            // adxRibbonButton1
            // 
            this.adxRibbonButton1.Caption = "Analysis";
            this.adxRibbonButton1.Id = "adxRibbonButton_d1d6b2c5c61b4c97aaae713ea9547164";
            this.adxRibbonButton1.Image = 2;
            this.adxRibbonButton1.ImageList = this.ilMain;
            this.adxRibbonButton1.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton1.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            // 
            // adxRibbonButton2
            // 
            this.adxRibbonButton2.Caption = "Analysis";
            this.adxRibbonButton2.Id = "adxRibbonButton_d1d6b2c5c61b4c97aaae713ea9547164";
            this.adxRibbonButton2.Image = 2;
            this.adxRibbonButton2.ImageList = this.ilMain;
            this.adxRibbonButton2.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton2.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton2.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            // 
            // adxRibbonButton3
            // 
            this.adxRibbonButton3.Caption = "Analysis";
            this.adxRibbonButton3.Id = "adxRibbonButton_d1d6b2c5c61b4c97aaae713ea9547164";
            this.adxRibbonButton3.Image = 2;
            this.adxRibbonButton3.ImageList = this.ilMain;
            this.adxRibbonButton3.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton3.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton3.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            // 
            // btnView
            // 
            this.btnView.Caption = "Columns";
            this.btnView.Id = "adxRibbonButton_4738252f541f462daea013d0309a0eb1";
            this.btnView.Image = 5;
            this.btnView.ImageList = this.ilMain;
            this.btnView.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnView.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.btnView.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.btnView.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.btnView_OnClick);
            // 
            // btnViewCol
            // 
            this.btnViewCol.Caption = "Fields";
            this.btnViewCol.ControlTag = "5a27efb9-c965-4baa-b0fb-fcb546fcf171";
            this.btnViewCol.Image = 5;
            this.btnViewCol.ImageList = this.ilMain;
            this.btnViewCol.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.btnViewCol.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.btnViewCol.UpdateCounter = 8;
            this.btnViewCol.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.btnViewCol_Click);
            // 
            // AddinModule
            // 
            this.AddinName = "OfficeAddin";
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;

        }
        #endregion

        #region Add-in Express automatic code

        // Required by Add-in Express - do not modify
        // the methods within this region

        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }

        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }

        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }

        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public static new AddinModule CurrentInstance
        {
            get
            {
                return AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule;
            }
        }

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }

        private void ShowSetting()
        {
            frmConnect = new FrmSystem();
            if (frmConnect.ShowDialog() == DialogResult.OK)
                LoadConfig("");
        }
        public void ShowFields()
        {
            Excel._Worksheet sheet = (Excel._Worksheet)ExcelApp.ActiveSheet;
            _xlsCell = (Excel.Range)ExcelApp.ActiveCell;
            string _address = _xlsCell.get_AddressLocal(1, 1, Excel.XlReferenceStyle.xlA1, 0, 0).ToString();
            _address = _address.Replace("$", "");
            string formular = _xlsCell.Formula.ToString();
            if (formular.Contains("TVC_QUERY") && formular.Contains("USER TABLE"))
            {
                string tablename = "data";
                try
                {
                    Excel.Range rangeTableName = ExcelApp.get_Range("A" + _xlsCell.Row);
                    tablename = rangeTableName.Value.ToString();
                }
                catch { }
                string tmp = formular.Replace("USER TABLE(", "");
                formular = tmp.Substring(0, tmp.Length - 1);
                SQLBuilder _sqlBuilder = new SQLBuilder(processingMode.Details);

                if (!formular.Contains("TVC_QUERY"))
                    Parsing.Formular2SQLBuilder(formular, ref _sqlBuilder);
                else
                {
                    Parsing.TVCFormular2SQLBuilder(formular, ref _sqlBuilder);
                }

                DataTable dt_list = new DataTable();
                if (_sqlBuilder.SelectedNodes.Count > 0)
                {
                    //CommoControl commo = new CommoControl();
                    //string connnectString = commo.CreateConnectString(Properties.Settings.Default.Server
                    //         , Properties.Settings.Default.User
                    //         , Properties.Settings.Default.Pass
                    //         , Properties.Settings.Default.DBName);

                    //a.THEME = this.THEME;
                    dt_list.TableName = tablename;
                    dt_list.Columns.Add("Name");
                    dt_list.Columns.Add("Code");

                    for (int i = 0; i < _sqlBuilder.SelectedNodes.Count; i++)
                    {
                        Node colum = _sqlBuilder.SelectedNodes[i];
                        string desc = colum.Description;
                        int dem = 0;
                        for (int j = i - 1; j >= 0; j--)
                        {
                            Node node = _sqlBuilder.SelectedNodes[j];
                            if (node.Description == colum.Description)
                            {
                                dem++;
                            }
                        }
                        if (dem > 0)
                        {
                            desc = colum.Description + dem;
                        }
                        dt_list.Rows.Add(new string[] { colum.Description, desc });
                    }
                    TVCDesigner.MainForm frm = new TVCDesigner.MainForm(dt_list, null, null);
                    frm.Show(new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode));
                }


            }
        }
        private void ShowDesign()
        {
            if (_user == "")
            {
                MessageBox.Show("Please login...");
                return;
            }
            Excel._Worksheet sheet = (Excel._Worksheet)ExcelApp.ActiveSheet;
            _xlsCell = (Excel.Range)ExcelApp.ActiveCell;
            string _address = _xlsCell.get_AddressLocal(1, 1, Excel.XlReferenceStyle.xlA1, 0, 0).ToString();
            _address = _address.Replace("$", "");
            string formular = _xlsCell.Formula.ToString();
            if (frm == null)
            {
                frm = new QDAddIn(_address, ExcelApp, formular, _strConnect, _strConnectDes, _user);
                frm.User = _user;
                frm.Config = _config;
                frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                //frm.Pos = _address;
                //frm.TopMost = true;
                //IWin32Window wincurrent = new WindowWrapper((IntPtr)ExcelApp.);
                frm.Show(new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode));
            }
            //else if (frm.DialogResult == System.Windows.Forms.DialogResult.Yes)
            //{
            //    frm.BringToFront();
            //    frm.GetQueryBuilderFromFomular(formular);
            //}
            else
            {
                frm.Close();

                frm = new QDAddIn(_address, ExcelApp, formular, _strConnect, _strConnectDes, _user);
                frm.User = _user;
                frm.Config = _config;
                frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                //frm.Pos = _address;
                //frm.TopMost = true;
                //IWin32Window wincurrent = new WindowWrapper((IntPtr)ExcelApp.Hwnd);
                frm.Show(new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode));
            }
        }
        private void ShowComment()
        {
            if (_user == "")
            {
                MessageBox.Show("Please login...");
                return;
            }
            try
            {
                Excel._Worksheet sheet = (Excel._Worksheet)ExcelApp.ActiveSheet;
                _xlsCell = (Excel.Range)ExcelApp.ActiveCell;
                string _address = _xlsCell.get_AddressLocal(1, 1, Excel.XlReferenceStyle.xlA1, 0, 0).ToString().Replace("$", "");
                if (_xlsCell.Comment != null)
                {
                    string formular = _xlsCell.Comment.Text(Type.Missing, Type.Missing, Type.Missing);
                    if (frm == null)
                    {
                        frm = new QDAddIn(_address, ExcelApp, formular, _strConnect, _strConnect, _user);
                        //frm.User = _user;
                        frm.Config = _config;
                        frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                        //frm.Pos = _address;
                        //frm.TopMost = true;
                        frm.Show(new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode));
                    }
                    //else if (frm.DialogResult == System.Windows.Forms.DialogResult.Yes)
                    //{
                    //    frm.BringToFront();
                    //    frm.GetQueryBuilderFromFomular(formular);
                    //}
                    else
                    {
                        frm.Close();
                        frm = new QDAddIn(_address, ExcelApp, formular, _strConnect, _strConnectDes, _user);
                        frm.User = _user;
                        frm.Config = _config;
                        frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                        //frm.Pos = _address;
                        //frm.TopMost = true;
                        frm.Show(new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode));//new WindowWrapper((IntPtr)ExcelApp.Hwnd)
                    }
                }
                else { MessageBox.Show("Cell selected is incorrect!"); }

            }
            catch (Exception ex)
            {
                BUS.CommonControl.AddLog("ErroLog", __documentDirectory + "\\Log", "Addin : " + ex.Message + "\n\t" + ex.Source + "\n\t" + ex.StackTrace);
            }
        }
        private void ShowAnalysis()
        {
            Excel._Worksheet sheet = ExcelApp.ActiveSheet as Excel.Worksheet;
            _xlsCell = (Excel.Range)ExcelApp.ActiveCell;
            string _address = _xlsCell.get_AddressLocal(1, 1, Excel.XlReferenceStyle.xlA1, 0, 0).ToString().Replace("$", "");
            if (_xlsCell.Comment != null)
            {
                string formular = _xlsCell.Comment.Text(Type.Missing, Type.Missing, Type.Missing);
                if (frmdrill == null)
                {
                    frmdrill = new QDAddinDrillDown(_address, ExcelApp, formular, _strConnectDes, _user);
                    frmdrill.User = _user;
                    frmdrill.Config = _config;
                    frmdrill.FormClosed += new FormClosedEventHandler(frmdrill_FormClosed);
                    //frm.Pos = _address;
                    //frm.TopMost = true;
                    frmdrill.Show(new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode));
                }
                //else if (frmdrill.DialogResult == System.Windows.Forms.DialogResult.Yes)
                //{
                //    frmdrill.BringToFront();
                //    frmdrill.GetQueryBuilderFromFomular(formular);
                //}
                else
                {
                    frmdrill.Close();
                    frmdrill = new QDAddinDrillDown(_address, ExcelApp, formular, _strConnectDes, _user);
                    frmdrill.User = _user;
                    frmdrill.Config = _config;
                    frmdrill.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frmdrill_FormClosed);
                    //frm.Pos = _address;
                    //frm.TopMost = true;
                    frmdrill.Show(new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode));//new WindowWrapper((IntPtr)ExcelApp.Hwnd)
                }
            }
            else { MessageBox.Show("Cell selected is incorrect!"); }
        }

        private void adxExcelEvents_SheetSelectionChange(object sender, object sheet, object range)
        {
            Excel.Range Target = range as Excel.Range;
            if (frm != null && frm.Status == "I")
            {
                _address = Target.get_AddressLocal(Target.Row, Target.Column, Excel.XlReferenceStyle.xlA1, 0, 0);
                string address = _address.Replace("$", "");
                string value = "";
                try { value = Target.Value.ToString(); }//(type)
                catch { }
                frm.SetValueFocus(address, value);
            }
        }

        private void btnSetting_Click(object sender)
        {
            ShowSetting();
        }


        private void btnDesign_Click(object sender)
        {
            ShowDesign();
        }


        private void btnComment_Click(object sender)
        {
            ShowComment();
        }


        public class WindowWrapper : System.Windows.Forms.IWin32Window
        {
            public WindowWrapper(IntPtr handle)
            {
                _hwnd = handle;
            }

            public IntPtr Handle
            {
                get { return _hwnd; }
            }

            private IntPtr _hwnd;
        }
        void frm_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                Excel._Worksheet a = ExcelApp.ActiveSheet as Excel.Worksheet;
                if (frm.Status == "C")
                {
                    try
                    {
                        a.get_Range(frm.Pos, type).ClearComments();
                    }
                    catch { }
                    a.get_Range(frm.Pos, type).AddComment(frm.TTFormular);
                }
                else if (frm.Status == "T")
                {
                    a.get_Range(frm.Pos, type).Value = String.Format("<#{0}>", frm.TTFormular);
                }
                else if (frm.Status == "L")
                {
                    try
                    {
                        //DataTable dt = frm.DataReturn;

                        System.Data.DataTable dt = frm.DataReturn;
                        //Excel.DataTable dtEx;
                        //Excel.Workbook _wbook = (Excel.Workbook)Application.ActiveWorkbook;
                        //_wbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                        Excel.Worksheet _wsheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                        Excel.Range currentRange = _wsheet.get_Range(_address, Type.Missing);

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            Excel.Range _range = (Excel.Range)_wsheet.Cells[currentRange.Row, i + currentRange.Column];
                            _range.Font.Bold = true;
                            _range.Value = dt.Columns[i].ColumnName;
                        }
                        for (int i = 0; i < dt.Rows.Count; i++)
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                Excel.Range _range = (Excel.Range)_wsheet.Cells[i + currentRange.Row + 1, j + currentRange.Column];
                                _range.Value = dt.Rows[i][j];
                            }
                        //string add = _wsheet.Name + "!R1C1:R" + (dt.Rows.Count + 1) + "C" + dt.Columns.Count;

                        //_wbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                        //Excel.Worksheet _wpivotsheet = (Excel.Worksheet)Application.ActiveWorkbook.ActiveSheet;
                        //string des = _wpivotsheet.Name + "!R3C1";
                        //_wbook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, add).CreatePivotTable(des, "PivotTable1", Type.Missing, Excel.XlPivotTableVersionList.xlPivotTableVersion10);
                    }
                    catch (Exception ex) { BUS.CommonControl.AddLog("ErroLog", __documentDirectory + "\\Log", "[Addin] [" + DateTime.Now.ToString() + "] : " + ex.Message + "\n\t" + ex.Source + "\n\t" + ex.StackTrace); }
                    //a.get_Range(frm.Pos, type).AddComment(frm.TTFormular);
                }
                else if (frm.Status == "F")
                    a.get_Range(frm.Pos, type).Value = "=" + frm.TTFormular;
                else if (frm.Status == "U")
                    a.get_Range(frm.Pos, type).Value = frm.TTFormular;
            }
        }

        private void btnAnalysis_Click(object sender)
        {
            ShowAnalysis();
        }


        void frmdrill_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (frmdrill.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                System.Data.DataTable dt = frmdrill._dataTable;
                Excel.DataTable dtEx;
                Excel.Workbook _wbook = (Excel.Workbook)ExcelApp.ActiveWorkbook;
                _wbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                Excel.Worksheet _wsheet = (Excel.Worksheet)ExcelApp.ActiveSheet;
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    Excel.Range _range = (Excel.Range)_wsheet.Cells[1, i + 1];
                    _range.Font.Bold = true;
                    _range.Value = dt.Columns[i].ColumnName;
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        Excel.Range _range = (Excel.Range)_wsheet.Cells[i + 2, j + 1];
                        _range.Value = dt.Rows[i][j];
                    }
                string add = _wsheet.Name + "!R1C1:R" + (dt.Rows.Count + 1) + "C" + dt.Columns.Count;

                _wbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                Excel.Worksheet _wpivotsheet = (Excel.Worksheet)ExcelApp.ActiveSheet;
                string des = _wpivotsheet.Name + "!R3C1";
                _wbook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, add).CreatePivotTable(des, "PivotTable1", Type.Missing);//, Excel. Excel.XlPivotTableVersionList.xlPivotTableVersion10);
            }
        }

        private void adxExcelEvents_SheetBeforeDoubleClick(object sender, ADXExcelSheetBeforeEventArgs e)
        {
            Excel.Range Target = e.Range as Excel.Range;
            _address = Target.get_AddressLocal(Target.Row, Target.Column, Excel.XlReferenceStyle.xlA1, 0, 0).Replace("$", "");
            if (Target.Formula != null)
            {
                string formular = Target.Formula.ToString();

                if (formular.Contains("TT_XLB_EB") || formular.Contains("USER TABLE"))
                {

                    //Target.set_Value(Type.Missing, formular);
                    //Application.Undo();
                    if (frm != null)
                    {
                        frm.Close();
                        frm = new QDAddIn(_address, ExcelApp, formular, _strConnect, _strConnectDes, _user);
                        frm.Config = _config;
                        frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                        //frm.Pos = _address;
                        //if (value.Contains("TT_XLB_ED"))
                        //    frm.GetQueryBuilderFromFomular(value);
                        //frm.TopMost = true;
                        frm.Show(new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode));//new WindowWrapper((IntPtr)ExcelApp.Hwnd)
                    }
                    else
                    {
                        frm = new QDAddIn(_address, ExcelApp, formular, _strConnect, _strConnectDes, _user);
                        frm.Config = _config;
                        frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                        //frm.Pos = _address;
                        //if (value.Contains("TT_XLB_ED"))
                        //    frm.GetQueryBuilderFromFomular(value);
                        //frm.TopMost = true;
                        //
                        frm.Show(new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode));//new WindowWrapper((IntPtr)ExcelApp.Hwnd)
                    }
                    frm.Focus();
                    e.Cancel = true;

                }
            }
            else
            {
                if (Target.Text != null)
                {
                    string formular = Target.Text.ToString();

                    if (formular.Contains("TT_XLB_EB") || formular.Contains("USER TABLE"))
                    {

                        //Target.set_Value(Type.Missing, formular);
                        //Application.Undo();
                        if (frm != null)
                        {
                            frm.Close();
                            frm = new QDAddIn(_address, ExcelApp, formular, _strConnect, _strConnectDes, _user);
                            frm.Config = _config;
                            frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                            //frm.Pos = _address;
                            //if (value.Contains("TT_XLB_ED"))
                            //    frm.GetQueryBuilderFromFomular(value);
                            //frm.TopMost = true;
                            frm.Show();//new WindowWrapper((IntPtr)ExcelApp.Hwnd)
                        }
                        else
                        {
                            frm = new QDAddIn(_address, ExcelApp, formular, _strConnect, _strConnectDes, _user);
                            frm.Config = _config;
                            frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                            //frm.Pos = _address;
                            //if (value.Contains("TT_XLB_ED"))
                            //    frm.GetQueryBuilderFromFomular(value);
                            //frm.TopMost = true;
                            //

                            frm.Show(new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode));//
                        }
                        frm.Focus();
                        e.Cancel = true;

                    }
                }
            }
        }

        private void btnRSetting_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            ShowSetting();
        }

        private void btnRDesign_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            ShowDesign();
        }

        private void btnRComment_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            ShowComment();
        }

        private void btnRAnalysis_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            ShowAnalysis();
        }

        private void btnLogin_Click(object sender)
        {
            ShowLogin();
        }

        private void ShowLogin()
        {
            if (frmLog == null)
            {
                frmLog = new frmLoginEx();


                IWin32Window wincurrent = new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode);
                if (frmLog.ShowDialog(wincurrent) == DialogResult.OK)
                {
                    _user = frmLog.User;
                }
            }
            else
            {
                frmLog.Close();
                frmLog = new frmLoginEx();
                IWin32Window wincurrent = new WindowWrapper((IntPtr)ExcelApp.DDEAppReturnCode);
                if (frmLog.ShowDialog(wincurrent) == DialogResult.OK)
                {
                    _user = frmLog.User;
                }
            }
        }

        private void btnRLogin_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            ShowLogin();
        }

        private void btnView_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            ShowFields();
        }

        private void btnViewCol_Click(object sender)
        {
            ShowFields();
        }



    }
}

