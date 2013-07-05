using System;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using QueryDesigner;
using System.IO;
using BUS;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        QDConfig _config = new QDConfig();
        QueryDesigner.QDAddIn frm;
        QueryDesigner.QDAddinDrillDown frmdrill;
        QueryDesigner.FrmSystem frmConnect;
        object type = Type.Missing;
        Excel.Range _xlsCell;
        public static string _strConnect = "";
        public static string _strConnectDes = "";
        string __templatePath = "";
        string __reportPath = "";

        //private Excel.Menu _menu;
        static Office.CommandBar _oCommand;

        static Office.CommandBarPopup _oPop;
        static Office.CommandBarButton _oBtn;
        static Office.CommandBarButton _oBtnComment;
        static Office.CommandBarButton _oBtnAnalysis;
        static Office.CommandBarButton _oBtnConnect;
        //static Office.CommandBarControl _oBtnControl;
        public static string __documentDirectory = string.Empty;
        string _sErr = "";
        string _address = "A1";
        string _conn_ID = "";
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            #region VSTO generated code InitDocument();
            try
            {
                InitDocument();
                LoadConfig("");
            }
            catch (Exception ex)
            {
                BUS.CommonControl.AddLog("ErroLog", __documentDirectory + "\\Log", "Addin : " + ex.Message + "\n\t" + ex.Source + "\n\t" + ex.StackTrace);
            }
            try
            {
                Office.CommandBar imenu = (Office.CommandBar)Application.CommandBars["TVC-QD"];


                //imenu.Delete(); 
                if (imenu == null || imenu.Controls.Count == 0)
                    AddMenuItem(imenu);
            }
            catch (Exception ex)
            {
                //System.Windows.Forms.MessageBox.Show(ex.Message);
                BUS.CommonControl.AddLog("ErroLog", __documentDirectory + "\\Log", "Addin : " + ex.Message + "\n\t" + ex.Source + "\n\t" + ex.StackTrace);
                AddMenuItem(null);
            }

            Application.SheetSelectionChange += new Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler(Application_SheetSelectionChange);
            Application.SheetBeforeDoubleClick += new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeDoubleClickEventHandler(Application_SheetBeforeDoubleClick);

            #endregion

        }

        private void InitDocument()
        {
            string filename = Application.StartupPath + "\\Configuration\\xmlConnect.xml";
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
            string connectionFile = configureDirectory + "\\xmlConnect.xml";
            if (!File.Exists(connectionFile))
            {
                File.Copy(filename, connectionFile);
            }
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
                _strConnectDes = _config.GetConnection(ref strAP, "AP");
                _conn_ID = strAP;



                if (_config.DIR.Rows.Count > 0)
                {
                    __templatePath = _config.DIR.Rows[0]["TMP"].ToString();
                    __reportPath = _config.DIR.Rows[0]["RPT"].ToString();
                }
                if (_config.SYS.Rows.Count > 0)

                    ReportGenerator.User2007 = (bool)_config.SYS.Rows[0][_config.SYS.USE2007Column];
            }
        }
        private void AddMenuItem(Office.CommandBar imenu)
        {
            try
            {
                Application.ScreenUpdating = true;
                //Application.rev
                if (imenu == null)
                {
                    _oCommand = Application.CommandBars.Add("TVC-QD", Office.MsoBarPosition.msoBarTop, Type.Missing, Type.Missing);
                    //_oCommand.Name = "TVC-QD";
                    _oCommand.Visible = true;
                }
                else
                    _oCommand = imenu;

                _oPop = (Office.CommandBarPopup)_oCommand.Controls.Add(Office.MsoControlType.msoControlPopup, Type.Missing, Type.Missing, 1, true);
                _oPop.Caption = "TVC-QD";
                _oPop.Enabled = true;


                _oBtn = (Office.CommandBarButton)_oPop.CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, 1, true);
                _oBtn.DescriptionText = _oBtn.Caption = "Balance";
                _oBtn.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(_oBtn_Click);
                _oBtn.Picture = PictureDispConverter.ToIPictureDisp(Properties.Resources.blance);

                _oBtnComment = (Office.CommandBarButton)_oPop.CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, 1, true);
                _oBtnComment.Caption = "Comment";
                _oBtnComment.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(_oBtnComment_Click);
                _oBtnComment.Picture = PictureDispConverter.ToIPictureDisp(Properties.Resources.comment);

                _oBtnAnalysis = (Office.CommandBarButton)_oPop.CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, 1, true);
                _oBtnAnalysis.Caption = "Analysis";
                _oBtnAnalysis.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(_oBtnAnalysis_Click);
                _oBtnAnalysis.Picture = PictureDispConverter.ToIPictureDisp(Properties.Resources.analysis);

                _oBtnConnect = (Office.CommandBarButton)_oPop.CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, 1, true);
                _oBtnConnect.Caption = "Connection";
                _oBtnConnect.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(_oBtnConnect_Click);
                _oBtnConnect.Picture = PictureDispConverter.ToIPictureDisp(Properties.Resources.connect);
            }
            catch (Exception ex)
            {
                BUS.CommonControl.AddLog("ErroLog", __documentDirectory + "\\Log", "Addin : " + ex.Message + "\n\t" + ex.Source + "\n\t" + ex.StackTrace);
            }
        }

        void _oBtnConnect_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            frmConnect = new QueryDesigner.FrmSystem();
            if (frmConnect.ShowDialog() == DialogResult.OK)
                LoadConfig("");
        }

        void _oBtnAnalysis_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Excel._Worksheet sheet = (Excel._Worksheet)Application.ActiveWorkbook.ActiveSheet;
            _xlsCell = (Excel.Range)Application.ActiveCell;
            string _address = _xlsCell.get_AddressLocal(1, 1, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, 0, 0).ToString().Replace("$", "");
            string formular = _xlsCell.Comment.Text(Type.Missing, Type.Missing, Type.Missing);
            if (frmdrill == null)
            {
                frmdrill = new QDAddinDrillDown(_address, Application, formular, _strConnectDes);
                frmdrill.Config = _config;
                frmdrill.FormClosed += new FormClosedEventHandler(frmdrill_FormClosed);
                //frm.Pos = _address;
                //frm.TopMost = true;
                frmdrill.Show(new WindowWrapper((IntPtr)Application.Hwnd));
            }
            //else if (frmdrill.DialogResult == System.Windows.Forms.DialogResult.Yes)
            //{
            //    frmdrill.BringToFront();
            //    frmdrill.GetQueryBuilderFromFomular(formular);
            //}
            else
            {
                frmdrill.Close();
                frmdrill = new QDAddinDrillDown(_address, Application, formular, _strConnectDes);
                frmdrill.Config = _config;
                frmdrill.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frmdrill_FormClosed);
                //frm.Pos = _address;
                //frm.TopMost = true;
                frmdrill.Show(new WindowWrapper((IntPtr)Application.Hwnd));
            }
        }

        void frmdrill_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (frmdrill.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                System.Data.DataTable dt = frmdrill._dataTable;
                Microsoft.Office.Interop.Excel.DataTable dtEx;
                Excel.Workbook _wbook = (Excel.Workbook)Application.ActiveWorkbook;
                _wbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                Excel.Worksheet _wsheet = (Excel.Worksheet)Application.ActiveWorkbook.ActiveSheet;
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    Excel.Range _range = (Excel.Range)_wsheet.Cells[1, i + 1];
                    _range.Font.Bold = true;
                    _range.set_Value(Type.Missing, dt.Columns[i].ColumnName);
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        Excel.Range _range = (Excel.Range)_wsheet.Cells[i + 2, j + 1];
                        _range.set_Value(Type.Missing, dt.Rows[i][j]);
                    }
                string add = _wsheet.Name + "!R1C1:R" + (dt.Rows.Count + 1) + "C" + dt.Columns.Count;

                _wbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                Excel.Worksheet _wpivotsheet = (Excel.Worksheet)Application.ActiveWorkbook.ActiveSheet;
                string des = _wpivotsheet.Name + "!R3C1";
                _wbook.PivotCaches().Add(Microsoft.Office.Interop.Excel.XlPivotTableSourceType.xlDatabase, add).CreatePivotTable(des, "PivotTable1", Type.Missing, Microsoft.Office.Interop.Excel.XlPivotTableVersionList.xlPivotTableVersion10);
            }
        }

        void _oBtnComment_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Excel._Worksheet sheet = (Excel._Worksheet)Application.ActiveWorkbook.ActiveSheet;
            _xlsCell = (Excel.Range)Application.ActiveCell;
            string _address = _xlsCell.get_AddressLocal(1, 1, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, 0, 0).ToString().Replace("$", "");
            string formular = _xlsCell.Comment.Text(Type.Missing, Type.Missing, Type.Missing);
            if (frm == null)
            {
                frm = new QDAddIn(_address, Application, formular, _strConnect, _strConnect);
                frm.Config = _config;
                frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                //frm.Pos = _address;
                //frm.TopMost = true;
                frm.Show(new WindowWrapper((IntPtr)Application.Hwnd));
            }
            //else if (frm.DialogResult == System.Windows.Forms.DialogResult.Yes)
            //{
            //    frm.BringToFront();
            //    frm.GetQueryBuilderFromFomular(formular);
            //}
            else
            {
                frm.Close();
                frm = new QDAddIn(_address, Application, formular, _strConnect, _strConnectDes);
                frm.Config = _config;
                frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                //frm.Pos = _address;
                //frm.TopMost = true;
                frm.Show(new WindowWrapper((IntPtr)Application.Hwnd));
            }
        }

        void _oBtn_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Excel._Worksheet sheet = (Excel._Worksheet)Application.ActiveWorkbook.ActiveSheet;
            _xlsCell = (Excel.Range)Application.ActiveCell;
            string _address = _xlsCell.get_AddressLocal(1, 1, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, 0, 0).ToString();
            _address = _address.Replace("$", "");
            string formular = _xlsCell.Formula.ToString();
            if (frm == null)
            {
                frm = new QDAddIn(_address, Application, formular, _strConnect, _strConnectDes);
                frm.Config = _config;
                frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                //frm.Pos = _address;
                //frm.TopMost = true;
                IWin32Window wincurrent = new WindowWrapper((IntPtr)Application.Hwnd);
                frm.Show(wincurrent);
            }
            //else if (frm.DialogResult == System.Windows.Forms.DialogResult.Yes)
            //{
            //    frm.BringToFront();
            //    frm.GetQueryBuilderFromFomular(formular);
            //}
            else
            {
                frm.Close();
                frm = new QDAddIn(_address, Application, formular, _strConnect, _strConnectDes);
                frm.Config = _config;
                frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                //frm.Pos = _address;
                //frm.TopMost = true;
                IWin32Window wincurrent = new WindowWrapper((IntPtr)Application.Hwnd);
                frm.Show(wincurrent);
            }
        }

        void Application_SheetBeforeDoubleClick(object Sh, Microsoft.Office.Interop.Excel.Range Target, ref bool Cancel)
        {
            //string _address = "A1";
            _address = Target.get_AddressLocal(Target.Row, Target.Column, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, 0, 0).Replace("$", "");
            string formular = Target.Formula.ToString();

            if (formular.Contains("TT_XLB_EB") || formular.Contains("USER TABLE"))
            {

                //Target.set_Value(Type.Missing, formular);
                //Application.Undo();
                if (frm != null)
                {
                    frm.Close();
                    frm = new QDAddIn(_address, Application, formular, _strConnect, _strConnectDes);
                    frm.Config = _config;
                    frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                    //frm.Pos = _address;
                    //if (value.Contains("TT_XLB_ED"))
                    //    frm.GetQueryBuilderFromFomular(value);
                    //frm.TopMost = true;
                    frm.Show(new WindowWrapper((IntPtr)Application.Hwnd));
                }
                else
                {
                    frm = new QDAddIn(_address, Application, formular, _strConnect, _strConnectDes);
                    frm.Config = _config;
                    frm.FormClosed += new System.Windows.Forms.FormClosedEventHandler(frm_FormClosed);
                    //frm.Pos = _address;
                    //if (value.Contains("TT_XLB_ED"))
                    //    frm.GetQueryBuilderFromFomular(value);
                    //frm.TopMost = true;
                    //
                    frm.Show(new WindowWrapper((IntPtr)Application.Hwnd));
                }
                frm.Focus();
                Cancel = true;

            }
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
        void Application_SheetSelectionChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {

            if (frm != null && frm.Status == "I")
            {
                _address = Target.get_AddressLocal(Target.Row, Target.Column, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, 0, 0);
                string address = _address.Replace("$", "");
                string value = "";
                try { value = Target.get_Value(type).ToString(); }
                catch { }
                frm.SetValueFocus(address, value);
            }
        }
        void frm_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            if (frm.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                Excel._Worksheet a = (Excel._Worksheet)Application.ActiveWorkbook.ActiveSheet;
                if (frm.Status == "C")
                {
                    try
                    {
                        a.get_Range(frm.Pos, type).ClearComments();
                    }
                    catch { }
                    a.get_Range(frm.Pos, type).AddComment(frm.TTFormular);
                }
                else if (frm.Status == "L")
                {
                    try
                    {
                        //DataTable dt = frm.DataReturn;

                        System.Data.DataTable dt = frm.DataReturn;
                        //Microsoft.Office.Interop.Excel.DataTable dtEx;
                        //Excel.Workbook _wbook = (Excel.Workbook)Application.ActiveWorkbook;
                        //_wbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                        Excel.Worksheet _wsheet = (Excel.Worksheet)Application.ActiveWorkbook.ActiveSheet;
                        Microsoft.Office.Interop.Excel.Range currentRange = _wsheet.get_Range(_address, Type.Missing);

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            Excel.Range _range = (Excel.Range)_wsheet.Cells[currentRange.Row, i + currentRange.Column];
                            _range.Font.Bold = true;
                            _range.set_Value(Type.Missing, dt.Columns[i].ColumnName);
                        }
                        for (int i = 0; i < dt.Rows.Count; i++)
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                Excel.Range _range = (Excel.Range)_wsheet.Cells[i + currentRange.Row + 1, j + currentRange.Column];
                                _range.set_Value(Type.Missing, dt.Rows[i][j]);
                            }
                        //string add = _wsheet.Name + "!R1C1:R" + (dt.Rows.Count + 1) + "C" + dt.Columns.Count;

                        //_wbook.Sheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                        //Excel.Worksheet _wpivotsheet = (Excel.Worksheet)Application.ActiveWorkbook.ActiveSheet;
                        //string des = _wpivotsheet.Name + "!R3C1";
                        //_wbook.PivotCaches().Add(Microsoft.Office.Interop.Excel.XlPivotTableSourceType.xlDatabase, add).CreatePivotTable(des, "PivotTable1", Type.Missing, Microsoft.Office.Interop.Excel.XlPivotTableVersionList.xlPivotTableVersion10);
                    }
                    catch (Exception ex) { BUS.CommonControl.AddLog("ErroLog", __documentDirectory + "\\Log", "[Addin] [" + DateTime.Now.ToString() + "] : " + ex.Message + "\n\t" + ex.Source + "\n\t" + ex.StackTrace); }
                    //a.get_Range(frm.Pos, type).AddComment(frm.TTFormular);
                }
                else
                    a.get_Range(frm.Pos, type).set_Value(type, frm.TTFormular);
            }
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
