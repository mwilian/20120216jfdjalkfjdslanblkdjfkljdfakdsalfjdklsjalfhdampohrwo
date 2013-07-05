using System;
using System.Collections.Generic;

using System.Windows.Forms;
using System.IO;
using System.ComponentModel;

namespace dCube
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            bool flagCmd = false;
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].Contains("tavico://"))
                {
                    flagCmd = true;
                    break;
                }
            }
            if (flagCmd)
            {
                if (args.Length >= 2)
                {
                    LoadConfig(args[0]);
                    string[] arrCmd = args[1].Split('?');
                    string sErr = CmdManager.RunCmd(arrCmd[0], arrCmd[1]);
                }
            }
            else
                Application.Run(new Form_QD(args));
        }
        private static void LoadConfig(string db)
        {
            string __documentDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\" + Form_QD.DocumentFolder;
            string _strConnect = "";
            string _strConnectDes = "";
            string __templatePath = "";
            string __reportPath = "";
            string _appPath = Application.StartupPath;
            try
            {
                QDConfig _config = new QDConfig();
                if (File.Exists(__documentDirectory + "\\Configuration\\QDConfig.tvc"))
                {
                    _config.LoadConfig(__documentDirectory + "\\Configuration\\QDConfig.tvc");

                    string key = "";
                    _strConnect = _config.GetConnection(ref key, "QD");
                    key = "";
                    _strConnectDes = _config.GetConnection(ref key, "AP");

                    if (_config.DIR.Rows.Count > 0)
                    {
                        __templatePath = _config.DIR.Rows[0]["TMP"].ToString();
                        __reportPath = _config.DIR.Rows[0]["RPT"].ToString();
                    }

                    if (_config.SYS.Rows.Count > 0)
                    {
                        ReportGenerator.User2007 = (bool)_config.SYS.Rows[0][_config.SYS.USE2007Column];
                    }

                }
                QueryBuilder.SQLBuilder.SetConnection(_strConnect);
                ReportGenerator.Config = _config;
                CmdManager.Db = db;
                CmdManager.AppConnect = _strConnect;
                CmdManager.RepConnect = _strConnectDes;
                CmdManager.ReptPath = __reportPath;
                CmdManager.TempPath = __templatePath;
            }
            catch (Exception ex)
            {
            }
        }
    }
}
