using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
namespace QueryBuilder
{
    public class CoreCommonControl
    {
        CoreConnection _conn = new CoreConnection();
        public static string GetParseExpressionDate(string columnDate, string type)
        {
            string result = "SUBSTRING(CONVERT(" + columnDate + ",System.String),7,2)+'/'+SUBSTRING(CONVERT(" + columnDate + ",System.String),5,2)+'/'+SUBSTRING(CONVERT(" + columnDate + ",System.String),1,4)";
            switch (type)
            {
                case "A":
                    result = "SUBSTRING(CONVERT(" + columnDate + ",System.String),5,2)+'/'+SUBSTRING(CONVERT(" + columnDate + ",System.String),7,2)+'/'+SUBSTRING(CONVERT(" + columnDate + ",System.String),1,4)";
                    break;
                case "B":
                    result = "SUBSTRING(CONVERT(" + columnDate + ",System.String),7,2)+'/'+SUBSTRING(CONVERT(" + columnDate + ",System.String),5,2)+'/'+SUBSTRING(CONVERT(" + columnDate + ",System.String),1,4)";
                    break;
                case "C":
                    result = "SUBSTRING(CONVERT(" + columnDate + ",System.String),1,4)+'/'+SUBSTRING(CONVERT(" + columnDate + ",System.String),5,2)+'/'+SUBSTRING(CONVERT(" + columnDate + ",System.String),7,2)";
                    break;
            }
            return result;
        }
        public static string GetParseExpressionPeriod(string columnDate)
        {
            string result = "SUBSTRING(CONVERT(" + columnDate + ",System.String),5,3)+'/'+SUBSTRING(CONVERT(" + columnDate + ",System.String),1,4)";
            return result;
        }
        public IDataReader executeQuery(string sqlString)
        {
            _conn.connect();
            IDataReader tmp = _conn.executeQuery(sqlString);
            _conn.disconnect();
            return tmp;
        }
        public DataTable executeSelectQuery(string sqlString)
        {
            return _conn.executeSelectQuery(sqlString);
        }
        public DataTable executeSelectQuery(string sqlString, string strConnection)
        {
            return _conn.executeSelectQuery(sqlString, strConnection);
        }
        public void executeNonQuery(string sqlString)
        {
            _conn.connect();
            _conn.executeNonQuery(sqlString);
            _conn.disconnect();
        }

        public object executeScalar(string sqlString)
        {
            _conn.connect();
            object tmp = _conn.executeScalar(sqlString);
            _conn.disconnect();
            return tmp;
        }
       

        public object executeScalar(string sqlString, string connectString)
        {
            return _conn.executeScalar(sqlString, connectString);
        }
        public static string ToRoman(int arabic)
        {
            //string result = "";
            //int[] arabic = new int[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            //string[] roman = new string[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            //int i = 0;
            //while (n >= arabic[i])
            //{
            //    n = n - arabic[i];
            //    result = result + roman[i];

            //    i = i + 1;
            //}
            //return result;
            /* Arabic Roman relation
            * 1000 = M
            * 900 = CM
            * 500 = D
            *400 = CD
            *100 = C
            *90 = XC
            *50 = L
            *40 = XL
            *10 = L
            *9 = IX
            *5 = V
            *4 = IV
            *1 = 1
            */
            string result = "";
            for (int i = 0; i < arabic; i++)
            {
                while (arabic >= 1000)
                {//check for thousands place

                    result = result + "M";
                    arabic = arabic - 1000;
                }
                while (arabic >= 900)
                {
                    //check for nine hundred place
                    result = result + "CM";
                    arabic = arabic - 900;
                }
                while (arabic >= 500)
                {
                    //check for five hundred place
                    result = result + "D";
                    arabic = arabic - 500;
                }
                while (arabic >= 400)
                {
                    //check for four hundred place
                    result = result + "CD";
                    arabic = arabic - 400;
                }
                while (arabic >= 100)
                {
                    //check for one hundred place
                    result = result + "C";
                    arabic = arabic - 100;
                }
                while (arabic >= 90)
                {
                    //check for ninety place
                    result = result + "XC";
                    arabic = arabic - 90;
                }
                while (arabic >= 50)
                {
                    //check for fifty place
                    result = result + "L";
                    arabic = arabic - 50;
                }
                while (arabic >= 40)
                {
                    // check for forty place
                    result = result + "XL";
                    arabic = arabic - 40;
                }

                while (arabic >= 10)
                {
                    // check for tenth place
                    result = result + "X";
                    arabic = arabic - 10;
                }
                while (arabic >= 9)
                {
                    //check for nineth place
                    result = result + "IX";
                    arabic = arabic - 9;
                }
                while (arabic >= 5)
                {
                    //check for fifth place
                    result = result + "V";
                    arabic = arabic - 5;
                }
                while (arabic >= 4)
                {
                    //check for fourth place
                    result = result + "IV";
                    arabic = arabic - 4;
                }
                while (arabic >= 1)
                {
                    //check for first place
                    result = result + "I";
                    arabic = arabic - 1;
                }
            }
            return result;
        }
        public static void SetConnection(string connect)
        {
            CoreConnection.SetConnection(connect);
        }

        public static string GetConnection()
        {
            return CoreConnection.ConnectionString;
        }



        public static void AddLog(string type, string path, string message)
        {
            string erroFile = path + "\\" + type + ".log";
            //System.IO.StreamWriter sw = new System.IO.StreamWriter(erroFile);
            if (!File.Exists(erroFile))
            {
                StreamWriter swt = File.CreateText(erroFile);
                swt.WriteLine(message);
                swt.Close();

            }
            else
            {
                FileStream file = new FileStream(erroFile, FileMode.Append);
                StreamWriter sw = new StreamWriter(file);
                sw.WriteLine(message);
                sw.Close();
                file.Close();
            }
        }
        public static string RemoveAttribute(String input, string attribute)
        {
            int indexS = input.IndexOf(attribute, 0);
            if (indexS != -1)
            {
                int indexE = input.IndexOf(";", indexS + attribute.Length);
                if (indexE == -1 || indexE == input.Length - 1)
                    return input.Substring(0, indexS);
                return input.Substring(0, indexS) + input.Substring(indexE + 1);
            }
            else
                return input;
        }
        public static string HiddenAttribute(String input, string attribute)
        {
            int indexS = input.IndexOf(attribute, 0);
            if (indexS != -1)
            {
                int indexE = input.IndexOf(";", indexS + attribute.Length);
                if (indexE == -1 || indexE == input.Length - 1)
                    return attribute + "=******;" + input.Substring(0, indexS);
                return input.Substring(0, indexS) + attribute + "=*****;" + input.Substring(indexE + 1);
            }
            else
                return attribute + "=*****;" + input;
        }
    }
}
