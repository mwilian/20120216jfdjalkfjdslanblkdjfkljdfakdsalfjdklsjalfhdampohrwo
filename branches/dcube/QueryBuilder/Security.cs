using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
//using System.Drawing;
using System.Diagnostics;
//using System.Windows.Forms;
namespace QueryBuilder
{
    //  Class Security
    // following class was VB module
    static public class Security
    {
        //  Method login
        public static bool login()
        {
            try
            {

            }
            catch (Exception ex)
            {
                throw new ArgumentException(ex.Message);
            }
            return false;
        }


        //  Method ActiveConStr
        public static string ActiveConStr()
        {
           // return pbs.Helper.Database.SUNDBConnection;
            return "";
        }


        //  Method NZ
        public static string NZ(string str, string nullValue)
        {
            if (str == null)
            {
                return nullValue;
            }
            return str.Trim();
        }

        // TRANSWARNING: Automatically generated because of optional parameter(s) 
        //  Method NZ
        public static string NZ(string str)
        {
            return NZ(str, "");
        }

    }


}
