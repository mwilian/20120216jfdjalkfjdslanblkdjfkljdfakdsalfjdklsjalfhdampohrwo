using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
namespace DAO
{
    public class ErrorLog
    {        
        internal static void SetLog(string sErr)
        {
            string result = DateTime.Now.ToString() + ": " + sErr;
            string fileName = "LogErro.log";
        ////    StreamWriter streamWr = new StreamWriter(fileName);
        //    streamWr.NewLine = result;
        //    streamWr.Close();
        }
    }
}
    