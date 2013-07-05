using System;
using System.Collections.Generic;
using System.Text;
using QueryBuilder;
using System.Configuration;
using System.IO;
using System.Data;
using BUS;
using DTO;
using System.Diagnostics;
using dCube;

namespace QDCommand
{
    class Program
    {
        private static SQLBuilder _sqlBuilder = new SQLBuilder(processingMode.Details);
        //GridViewComboBoxColumn customerColumn = new GridViewComboBoxColumn();
        static string _strConnect = ConfigurationSettings.AppSettings["strConnect"].ToString();
        static string _strConnectDes = ConfigurationSettings.AppSettings["strConnect"].ToString();
        bool flag_view = false;
        static String sErr = "";
        string _strY = "Y";
        string _strN = "N";
        string THEME = "Office2010";
        Node[] _arrNodes = null;
        string _key = "newoppo123456789";
        string _iv = "12345678";
        string _padMode = "PKCS7";
        string _opMode = "CBC";
        string owner = "";
        string _pathLicense = Environment.CurrentDirectory + "\\License.bin";
        string _datePaterm = @"^([0]?[1-9]|[1|2][0-9]|[3][0|1])[./-]([0]?[1-9]|[1][0-2])[./-]([0-9]{4}|[0-9]{3}|[0-9]{2}|[0-9]{1})$";
        public static string __documentDirectory = string.Empty;
        public static string __reportPath = ConfigurationSettings.AppSettings["ReportPath"].ToString();
        public static string __templatePath = ConfigurationSettings.AppSettings["TemplatePath"].ToString();
        static string _sqlText = "";
        bool flagOpen = true;
        static QDConfig _config = new QDConfig();
        static int Main(string[] args)
        {
            int sErr = 1;
            try
            {
                InitDocument();
                if (args.Length > 5)
                {
                    string conn = args[0];
                    if (conn == "ZZZ")
                        conn = "";
                    string dtb = args[1];
                    string user = args[2];
                    string pass = args[3];
                    string method = args[4];
                    _sqlBuilder.ConnID = conn;
                    LoadConfig(conn);
                    ReportGenerator.Config = _config;
                    if (user == "TVC" && pass == "TVCSYS")
                    {
                        if (method == Command.OPEN.ToString())
                        {
                            OpenReport(args);
                        }
                        else if (method == Command.EXEC.ToString())
                        {
                            ExexReport(args);
                        }
                        else if (method == Command.EXECPDF.ToString())
                        {
                            ExexPDF(args);
                        }
                        else if (method == Command.OPENPDF.ToString())
                        {
                            OpenPDF(args);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                sErr = 0;
                BUS.CommonControl.AddLog("ErroLog", __documentDirectory + "\\Log", "[QDCommand]\t[" + DateTime.Now.ToString() + "]:\t" + ex.Message + ", " + ex.Source + ", " + ex.StackTrace);
            }
            return sErr;
        }

        private static void OpenPDF(string[] args)
        {
            string filename = ExexPDF(args);
            if (sErr == "" && File.Exists(filename))
            {
                Process.Start(filename);
            }
        }

        private static string ExexPDF(string[] args)
        {
            string dtb = args[1];

            string qdid = "";
            int index = 0;
            string value = "";
            if (args.Length >= 6) qdid = args[5];
            if (args.Length >= 7) index = Convert.ToInt32(args[6]);
            if (args.Length >= 8) value = args[7];

            BUS.LIST_QDControl qdCtr = new LIST_QDControl();
            DTO.LIST_QDInfo qdInfo = qdCtr.Get_LIST_QD(dtb, qdid, ref sErr);
            if (qdInfo.QD_ID != "")
            {
                _sqlText = qdInfo.SQL_TEXT;
                try
                {
                  //  ;General Timeout=100
                    LoadQD(qdInfo);
                    if (value != "")
                    {
                        _sqlBuilder.Filters[index].FilterFrom = _sqlBuilder.Filters[index].FilterTo = _sqlBuilder.Filters[index].ValueTo = _sqlBuilder.Filters[index].ValueFrom = value;
                    }
                    BUS.DBAControl dbaCtr = new DBAControl();
                    DTO.DBAInfo dbaInf = dbaCtr.Get(dtb, ref sErr);
                    __templatePath = dbaInf.REPORT_TEMPLATE_DRIVER;
                    ReportGenerator report = new ReportGenerator(_sqlBuilder, qdInfo.QD_ID, _sqlText, _strConnectDes, __templatePath, __reportPath, __documentDirectory);
                    return report.ExportPDFToPath(__reportPath);
                }
                catch (Exception ex)
                {
                    sErr = ex.Message;
                    BUS.CommonControl.AddLog("ErroLog", __documentDirectory + "\\Log", String.Format("[QDCommand]\t[{0}]:\t{1}, {2}, {3}", DateTime.Now, ex.Message, ex.Source, ex.StackTrace));
                    return "";
                }
            }
            return "";
        }

        private static string ExexReport(string[] args)
        {
            string dtb = args[1];
            string qdid = "";
            int index = 0;
            string value = "";
            if (args.Length >= 6) qdid = args[5];
            if (args.Length >= 7) index = args[6] == "ZZZ" ? 0 : Convert.ToInt32(args[6]);
            if (args.Length >= 8) value = args[7] == "ZZZ" ? "" : args[7];
            string path = "";
            if (args.Length >= 9) path = args[8] == "ZZZ" ? "" : args[8];
            path = path.Replace("%20", " ");
            string filename = "";
            if (args.Length >= 10) filename = args[9] == "ZZZ" ? "" : args[9];
            filename = filename.Replace("%20", " ");

            BUS.LIST_QDControl qdCtr = new LIST_QDControl();
            DTO.LIST_QDInfo qdInfo = qdCtr.Get_LIST_QD(dtb, qdid, ref sErr);
            if (qdInfo.QD_ID != "")
            {
                _sqlText = qdInfo.SQL_TEXT;
                try
                {
                    LoadQD(qdInfo);
                    if (value != "")
                    {
                        _sqlBuilder.Filters[index].FilterFrom = _sqlBuilder.Filters[index].FilterTo = _sqlBuilder.Filters[index].ValueTo = _sqlBuilder.Filters[index].ValueFrom = value;
                    }
                    BUS.DBAControl dbaCtr = new DBAControl();
                    DTO.DBAInfo dbaInf = dbaCtr.Get(dtb, ref sErr);
                    __templatePath = dbaInf.REPORT_TEMPLATE_DRIVER;

                    ReportGenerator report = new ReportGenerator(_sqlBuilder, qdInfo.QD_ID, _sqlText, _strConnectDes, __templatePath, __reportPath, __documentDirectory);
                   
                    if (path == "" && filename == "")
                        return report.ExportExcelToPath(__reportPath);
                    else if (path != "" && filename == "")
                        return report.ExportExcelToPath(path);
                    else if (path == "" && filename != "")
                        return report.ExportExcelToFile(__reportPath, filename);
                    else
                        return report.ExportExcelToFile(path, filename);
                }
                catch (Exception ex)
                {
                    sErr = ex.Message;
                    BUS.CommonControl.AddLog("ErroLog", __documentDirectory + "\\Log", "[QDCommand]\t[" + DateTime.Now.ToString() + "]:\t" + ex.Message + ", " + ex.Source + ", " + ex.StackTrace);
                    return "";
                }
            }
            return "";
        }
        private static void LoadQD(LIST_QDInfo info)
        {
            _sqlBuilder.Filters.Clear();
            _sqlBuilder.SelectedNodes.Clear();
            //radTabStrip1.Se
            _sqlBuilder = SQLBuilder.LoadSQLBuilderFromDataBase(info.QD_ID, info.DTB, info.ANAL_Q0.Trim());
        }
        private static void OpenReport(string[] args)
        {
            string filename = ExexReport(args);
            if (sErr == "" && File.Exists(filename))
            {
                Process.Start(filename);
            }
        }
        public static void InitDocument()
        {
            string filename = Environment.CurrentDirectory + "\\Configuration\\xmlConnect.xml";
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
            string logFolder = __documentDirectory + "\\Log";
            if (!Directory.Exists(logFolder))
            {
                Directory.CreateDirectory(logFolder);
            }
        }
        private static void LoadConfig(string strAP)
        {
            if (File.Exists(__documentDirectory + "\\Configuration\\QDConfig.tvc"))
            {
                _config.LoadConfig(__documentDirectory + "\\Configuration\\QDConfig.tvc");
                string key = "";
                _strConnect = _config.GetConnection(ref key, "QD");
                QueryBuilder.SQLBuilder.SetConnection(_strConnect);
                CommonControl.SetConnection(_strConnect);
                _strConnectDes = _config.GetConnection(ref strAP, "AP");
                _sqlBuilder.ConnID = strAP;


                if (_config.DIR.Count > 0)
                {
                    __templatePath = _config.DIR[0]["TMP"].ToString();
                    __reportPath = _config.DIR[0]["RPT"].ToString();
                }
                if (_config.SYS.Count > 0)

                    ReportGenerator.User2007 = (bool)_config.SYS[0][_config.SYS.USE2007Column];
            }


        }
    }
}
