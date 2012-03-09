using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using FlexCel.Core;
using System.IO;
using System.Text.RegularExpressions;

namespace QueryDesigner
{
    public static class CmdManager
    {
        private static string _db = "";
        private static string _appConnect = "";
        private static string _repConnect = "";
        private static string _tempPath = "";
        private static string _reptPath = "";
        static string sErr = "";

        public static string TempPath
        {
            get { return CmdManager._tempPath; }
            set { CmdManager._tempPath = value; }
        }
        public static string Db
        {
            get { return CmdManager._db; }
            set { CmdManager._db = value; }
        }
        public static string AppConnect
        {
            get { return CmdManager._appConnect; }
            set { CmdManager._appConnect = value; }
        }
        public static string RepConnect
        {
            get { return CmdManager._repConnect; }
            set { CmdManager._repConnect = value; }
        }
        public static string ReptPath
        {
            get { return CmdManager._reptPath; }
            set { CmdManager._reptPath = value; }
        }
        public static string RunCmd(string cmd, string query)
        {
            sErr = "";
            string[] param = query.Split('&');
            Dictionary<string, object> valueList = new Dictionary<string, object>();
            for (int i = 0; i < param.Length; i++)
            {
                string[] value = param[i].Split('=');
                valueList.Add(value[0], value[1]);
            }
            switch (cmd)
            {
                case "tavico://TASK":
                case "TASK":
                    return TaskCmd(valueList);
            }
            return sErr;
        }

        private static string TaskCmd(Dictionary<string, object> valueList)
        {
            BUS.LIST_TASKControl ctr = new BUS.LIST_TASKControl();
            if (_appConnect != "")
                BUS.CommonControl.SetConnection(_appConnect);
            object id = null;
            if (valueList.TryGetValue("id", out id))
            {
                DTO.LIST_TASKInfo infTask = ctr.Get(_db, id.ToString(), ref sErr);
                ReportGenerator rgAtt = null;
                ReportGenerator rgCnt = null;
                if (infTask.AttQD_ID != "")
                {
                    QueryBuilder.SQLBuilder sqlBuiderA = QueryBuilder.SQLBuilder.LoadSQLBuilderFromDataBase(infTask.AttQD_ID, _db, "");
                    rgAtt = new ReportGenerator(sqlBuiderA, infTask.AttQD_ID, "", _repConnect, _tempPath, _reptPath);
                }
                else
                {
                    DataSet ds = new DataSet();
                    rgAtt = new ReportGenerator(ds, infTask.AttTmp, _db, _reptPath, _tempPath, _reptPath);
                }
                if (infTask.CntQD_ID != "")
                {
                    QueryBuilder.SQLBuilder sqlBuiderC = QueryBuilder.SQLBuilder.LoadSQLBuilderFromDataBase(infTask.CntQD_ID, _db, "");
                    rgCnt = new ReportGenerator(sqlBuiderC, infTask.CntQD_ID, "", _repConnect, _tempPath, _reptPath);
                }
                else
                {
                    DataSet ds = new DataSet();
                    rgCnt = new ReportGenerator(ds, infTask.CntTmp, _db, _repConnect, _tempPath, _reptPath);
                }
                rgAtt.ValueList = valueList;
                rgCnt.ValueList = valueList;
                ExcelFile xls = rgAtt.CreateReport();
                rgCnt.Close();
                bool flagRun = false;
                string[] arrVRange = infTask.ValidRange.Split(';');
                if (arrVRange.Length >= 1)
                    for (int i = 1; i <= xls.SheetCount; i++)
                    {
                        TXlsNamedRange range = xls.GetNamedRange(arrVRange[0], 0);
                        if (range != null)
                        {
                            xls.ActiveSheet = range.SheetIndex;
                            object flag = xls.GetCellValue(range.Top, range.Left);
                            if (flag != null && !String.IsNullOrEmpty(flag.ToString().Trim()) && flag.ToString().Trim() != "0")
                            {
                                flagRun = true; break;
                            }
                        }
                    }
                string title = infTask.Description;

                if (flagRun)
                {
                    try
                    {
                        using (TextWriter wt = rgCnt.ExportHTML(_reptPath))
                        {
                            ExcelFile xls1 = rgCnt.XlsFile;
                            if (arrVRange.Length >= 2)
                            {
                                for (int i = 1; i <= xls1.SheetCount; i++)
                                {
                                    TXlsNamedRange range = xls1.GetNamedRange(arrVRange[1], 0);
                                    if (range != null)
                                    {
                                        xls1.ActiveSheet = range.SheetIndex;
                                        object flag = xls1.GetCellValue(range.Top, range.Left);
                                        if (flag != null && !String.IsNullOrEmpty(flag.ToString()))
                                        {
                                            title = flag.ToString();
                                            break;
                                        }
                                    }
                                }
                            }

                            string content = wt.ToString();
                            string filename = rgAtt.ExportExcelToFile(_reptPath, infTask.Description + ".xls");
                            Sendmail sendMail = new Sendmail(infTask.UserID, infTask.Password, infTask.Server, infTask.Protocol, Convert.ToInt32(infTask.Port));
                            string[] emails = infTask.Emails.Split(',');
                            Dictionary<string, string> arrayMail = new Dictionary<string, string>();
                            for (int i = 0; i < emails.Length; i++)
                            {
                                Match name = Regex.Match(emails[i], "\".+\"");
                                Match mail = Regex.Match(emails[i], "<.+>");
                                arrayMail.Add(mail.Value.Substring(1, mail.Value.Length - 2), name.Value.Substring(1, name.Value.Length - 2));
                            }
                            sErr = sendMail.SendMail(title, content, arrayMail, filename, true, true);

                        }
                    }
                    catch (Exception ex)
                    {
                        sErr = ex.Message;
                    }
                }

            }
            return sErr;
        }
    }
}
