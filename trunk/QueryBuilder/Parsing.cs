using System.ComponentModel;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
//using System.Drawing;
using System.Diagnostics;
//using System.Windows.Forms;
namespace QueryBuilder
{
    //  Class Parsing
    public class Parsing
    {

        public const string STR_SUM = "SUM";
        public const string STR_COUNT = "COUNT";
        public const string STR_AVERAGE = "AVG";
        public const string STR_MINIMUM = "MIN";
        public const string STR_MAXIMUM = "MAX";
        public const string STR_DISTINCTSUM = "DISTINCT SUM";
        public const string STR_DISTINCTCOUNT = "DISTINCT COUNT";
        public const string STR_DISTINCTAVERAGE = "DISTINCT AVG";

        private const string regexDB = @"(?<=0\,2\,)[^\,]+(?=\,)";
        private const string regexTable = @"(?<=0\,2\,[^\,]+\,)[_A-Za-z0-9]{1,}";
        private const string regexLedger = @"(?<=F\=)[^\,]+(?=\,K\=\/LA\/Ledger)";
        private const string regexDTB = "dtb=[^;]+";
        private const string regexTBL = "tbl=[^;]+";
        private const string regexLDG = "ldg=[^;]+";
        private const string regexFIL = "fil=((?!};).)+}";
        private const string regexOUT = "out=((?!};).)+}";
        private const string regexFROM = "f=[^;]+";
        private const string regexTO = "t=[^;]+";
        private const string regexOPERATOR = "o=[^;]+";
        private const string regexIS = "i=[^;]+";
        private const string regexKEY = "k=[^;]+";
        private const string regexAGGREGATE = "a=[^;]+";
        private const string regexC = "c=[^;]+";
        ///  <summary>
        ///    Step2 :Parse TTformular to sqlBuilder Object
        ///  </summary>
        ///  <param name="ParseString">TTformular </param>
        ///  <param name="_SQLBuilder">sqlBuilder Object</param>
        ///  <remarks></remarks>
        public static void Formular2SQLBuilder(string ParseString, ref SQLBuilder _SQLBuilder)
        {
            if (string.IsNullOrEmpty(ParseString))
            {
                return;
            }

            ParseString = Regex.Replace(ParseString, ".*(?=TT_XLB_EB)", string.Empty);

            string[] FromTo = new string[] { };
            string vParamsString = null; //  the string, contains parameters in TTformular
            string[] vParameter = new string[] { }; //  array of all parameters ex .$H$1,$G14,$H14,J$11,J$11
            int vparacount = 0;
            string vPosition = null; //  address of the formular

            // Return Database, table      
            string vTable = null;
            string vDatabase = null;
            vDatabase = Regex.Match(ParseString, regexDB).Value.ToString();
            vTable = Regex.Match(ParseString, regexTable).Value.ToString();

            _SQLBuilder.Table = vTable;

            string vFilter = null; // ex part of TTformular , contains filters
            vFilter = Regex.Match(ParseString, @"(?<=\,K\=)[^\,,.]+").Value.ToString();
            int i = 0;
            int n = 0;
            n = Regex.Matches(ParseString, @"F\=.*?,K").Count;

            // fill FromTo array
            if (n > 0)
            {
                FromTo = new string[n];
                foreach (System.Text.RegularExpressions.Match ft in Regex.Matches(ParseString, @"F\=.*?,K"))
                {
                    FromTo[i] = ft.Value.ToString();
                    i = i + 1;
                }
            }
            n = 0;
            i = 0;

            //  get string , contains parameters
            vParamsString = Regex.Match(ParseString, @"\" +  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ System.Convert.ToChar(34) + @"\,.+?\)").Value.ToString();

            // fill to parameter Array
            if (!(string.IsNullOrEmpty(vParamsString)))
            {
                vParamsString =  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vParamsString.Substring(1);
                vParamsString = vParamsString.Substring(1, vParamsString.Length - 2);// Strings.Mid(vParamsString, 1,  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vParamsString.Length - 1); 
                vParamsString = vParamsString + ","; //  them dau , cho de xu ly
                n = Regex.Matches(vParamsString, ".*?,").Count; //  cac tham so
                if (n > 0)
                {
                    vParameter = new string[n]; // tham so dau tien la vi tri cua cong thuc
                    foreach (System.Text.RegularExpressions.Match p in Regex.Matches(vParamsString, ".*?,"))
                    {
                        i = i + 1;
                        if (i == 1)
                        {
                            vPosition = p.Value.ToString().Replace(",", string.Empty);
                            _SQLBuilder.Pos = vPosition;
                        }
                        else
                        {
                            vParameter[i - 1] = p.Value.ToString().Replace(",", string.Empty);
                        }

                    }
                }
            }

            if (vDatabase.Contains("{P}"))
            {
                _SQLBuilder.DatabaseP = vDatabase;
                _SQLBuilder.Database = vParameter[System.Convert.ToInt32(double.Parse(vDatabase.Replace("{P}", string.Empty)))];
                vparacount = vparacount + 1;
                _SQLBuilder.DatabaseV = _SQLBuilder.ParaValueList[vparacount];
            }
            else
            {
                _SQLBuilder.Database = vDatabase;
                _SQLBuilder.DatabaseV = vDatabase;
                _SQLBuilder.DatabaseP = "";

            }

            i = 0;

            string vf = null;
            string vt = null;
            string vf1 = "";
            string vt1 = "";
            string filterf = null;
            string filtert = null;

            // identifying filters
            MatchCollection matchCollect = Regex.Matches(ParseString, @"(?<=\,K\=)[^\,,.]+");
            foreach (System.Text.RegularExpressions.Match m in matchCollect)
            {

                vFilter = m.Value.ToString();

                vf = Regex.Match(FromTo[i], "F=.*?,").Value.ToString();
                if (!(string.IsNullOrEmpty(vf)))
                {
                    vf =  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vf.Substring(2);
                    vf = vf.Substring(0, vf.Length - 1);// Strings.Mid(vf, 1,  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vf.Length - 1); 
                }
                //if (vf != "")
                //    vt = Regex.Match(FromTo[i].Replace(vf, "_"), "T.+?,").Value.ToString();
                //else
                vt = Regex.Match(FromTo[i], "T=.*?,").Value.ToString();
                if (!(string.IsNullOrEmpty(vt)))
                {
                    vt =  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vt.Substring(2);
                    vt = vt.Substring(0, vt.Length - 1);// Strings.Mid(vt, 1,  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vt.Length - 1); 
                }
                filterf = "";
                filtert = "";

                if (Regex.IsMatch(vf, "{P}"))
                {
                    filterf = vParameter[System.Convert.ToInt32(double.Parse(  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vf.Substring(3)))];
                }

                if (Regex.IsMatch(vt, "{P}"))
                {
                    filtert = vParameter[System.Convert.ToInt32(double.Parse(  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vt.Substring(3)))];
                }

                if (!(string.IsNullOrEmpty(vFilter)))
                {
                    vFilter =  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vFilter.Substring(1);
                    if (string.IsNullOrEmpty(filterf))
                    {
                        filterf = vf;
                    }
                    if (string.IsNullOrEmpty(filtert))
                    {
                        filtert = vt;
                    }
                    if (vFilter.ToUpper() == "LA/LEDGER")
                    { //  ledger lam rieng


                        if (Regex.IsMatch(vf, "{P}"))
                        {
                            _SQLBuilder.LedgerP = vf;
                            _SQLBuilder.Ledger = vParameter[System.Convert.ToInt32(double.Parse(  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ vf.Substring(3)))];
                            vparacount = vparacount + 1;
                            _SQLBuilder.LedgerV = _SQLBuilder.ParaValueList[vparacount];

                        }
                        else
                        {
                            _SQLBuilder.Ledger = vf;
                            _SQLBuilder.LedgerV = vf;
                            _SQLBuilder.LedgerP = "";
                        }
                    }
                    else
                    {
                        if (Regex.IsMatch(vf, "{P}"))
                        {
                            vparacount = vparacount + 1;
                            vf1 = _SQLBuilder.ParaValueList[vparacount]; // gia tri
                        }
                        else
                        {
                            vf1 = vf;
                        }

                        if (Regex.IsMatch(vt, "{P}"))
                        {
                            vparacount = vparacount + 1;
                            vt1 = _SQLBuilder.ParaValueList[vparacount]; // gia tri
                        }
                        else
                        {
                            vt1 = vt;
                        }

                        foreach (Node _node in SchemaDefinition.GetDecorateTableByCode(vTable, _SQLBuilder.Database))
                        {
                            if (_node.Code == vFilter)
                            {
                                _SQLBuilder.Filters.Add(new Filter(new Node(vFilter, _node.Description), filterf, filtert, vf1, vt1, vf, vt));

                                //  _SQLBuilder.SelectedNodes.Add(New Node(vOutputAgr(i), Output, _node.Description, _node.FType))
                            }
                        }
                        //  _SQLBuilder.Filters.Add(New Filter(New Node(vFilter, vFilter), filterf, filtert, vf1, vt1, vf, vt))

                    }


                }
                i = i + 1;
            }

            string Output = null;
            string[] vOutputAgr = null;

            n = Regex.Matches(ParseString, @"E\=.+?,").Count;
            i = 0;
            if (n > 0)
            {
                vOutputAgr = new string[n];
                foreach (System.Text.RegularExpressions.Match oe in Regex.Matches(ParseString, @"E\=.+?,"))
                {

                    vOutputAgr[i] = oe.Value.ToString().Substring(2, 1);// Strings.Mid(oe.Value.ToString(), 3, 1); 
                    i = i + 1;
                }
                i = 0;
                foreach (System.Text.RegularExpressions.Match o in Regex.Matches(ParseString, @"O\=.+?,"))
                {
                    Output = o.Value.ToString();
                    if (!(string.IsNullOrEmpty(Output)))
                    {
                        Output = Output.Replace(",", string.Empty);
                    }



                    vOutputAgr[i] = AgregateN2Code(vOutputAgr[i]);

                    if (!(string.IsNullOrEmpty(Output)))
                    {
                        Output = Output.Replace("O=/", string.Empty);
                        foreach (Node _node in SchemaDefinition.GetDecorateTableByCode(vTable, _SQLBuilder.Database))
                        {
                            // If Regex.IsMatch(_node.Code, Output & "$") Then
                            if (_node.Code.ToUpper() == Output.ToUpper())
                            {
                                _SQLBuilder.SelectedNodes.Add(new Node(vOutputAgr[i], Output, _node.Description, _node.FType, _node.NodeDesc));
                                break; /* TRANSWARNING: check that break is in correct scope */
                            }
                        }
                    }
                    i = i + 1;
                }
            }
            string sErr = "";
            CoreQD_SCHEMAControl schctr = new CoreQD_SCHEMAControl();
            CoreQD_SCHEMAInfo schInf = schctr.Get(_SQLBuilder.Database, _SQLBuilder.Table, ref sErr);
            _SQLBuilder.ConnID = schInf.DEFAULT_CONN;
        }


        public static SQLBuilder TVCFormular2SQLBuilder(string ParseString, ref SQLBuilder _SQLBuilder)
        {
            _SQLBuilder.Filters.Clear();
            _SQLBuilder.SelectedNodes.Clear();
            if (string.IsNullOrEmpty(ParseString))
            {
                return null;
            }
            if (Regex.IsMatch(ParseString, @".*(?=TVC_QUERY)"))
            {
                ParseString = Regex.Replace(ParseString, ".*(?=TVC_QUERY)", string.Empty);
                string vParamsString = Regex.Match(ParseString, @"\" +  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ System.Convert.ToChar(34) + @"\,.+?\)").Value.ToString();
                vParamsString = vParamsString.Substring(2, vParamsString.Length - 3);
                string[] arrParam = vParamsString.Split(',');
                _SQLBuilder.Pos = arrParam[0];

                //string[] arrParam = new string[arrParam1.Length - 1];
                //for (int i = 1; i < arrParam1.Length; i++)
                //    arrParam[i - 1] = arrParam1[i];

                string tmp = ParseString.Substring(ParseString.IndexOf('"') + 1);
                string value = tmp.Substring(0, tmp.IndexOf('"'));
                if (value != "")
                {
                    string formular = value.Substring(1, value.Length - 2);

                    if (Regex.IsMatch(formular, regexDTB))
                    {
                        Match m = Regex.Match(formular, regexDTB);
                        value = m.Value.Replace("dtb=", string.Empty).Replace(";", "");
                        if (Regex.IsMatch(value, "{P}"))
                        {
                            _SQLBuilder.DatabaseP = value;
                            int indexP = 0;
                            if (int.TryParse(value.Replace("{P}", ""), out indexP))
                            {
                                _SQLBuilder.Database = arrParam[indexP];
                                _SQLBuilder.DatabaseV = _SQLBuilder.ParaValueList[indexP];
                            }
                        }
                        else
                            _SQLBuilder.Database = value;
                    }
                    if (Regex.IsMatch(formular, regexTBL))
                    {
                        Match m = Regex.Match(formular, regexTBL);
                        value = m.Value.Replace("tbl=", string.Empty).Replace(";", "");

                        _SQLBuilder.Table = value;
                    }
                    if (Regex.IsMatch(formular, regexLDG))
                    {
                        Match m = Regex.Match(formular, regexLDG);
                        value = m.Value.Replace("ldg=", string.Empty).Replace(";", "");

                        //value = Regex.Replace(regexLDG, regexLDG, string.Empty);
                        if (Regex.IsMatch(value, "{P}"))
                        {
                            _SQLBuilder.LedgerP = value;
                            int indexP = 0;
                            if (int.TryParse(value.Replace("{P}", ""), out indexP))
                            {
                                _SQLBuilder.Ledger = arrParam[indexP];
                                _SQLBuilder.LedgerV = _SQLBuilder.ParaValueList[indexP];
                            }
                        }
                        else
                            _SQLBuilder.Ledger = value;


                    }
                    if (Regex.IsMatch(formular, regexFIL))
                    {
                        foreach (Match m in Regex.Matches(formular, regexFIL))
                        {
                            Match mtemp = Regex.Match(m.Value, "{.+}");
                            Filter f = Parse2Filter(mtemp.Value.Substring(1, mtemp.Value.Length - 2), _SQLBuilder.Table, _SQLBuilder.Database, arrParam, _SQLBuilder.ParaValueList);
                            if (f != null)
                                _SQLBuilder.Filters.Add(f);
                        }
                    }
                    if (Regex.IsMatch(formular, regexOUT))
                    {
                        foreach (Match m in Regex.Matches(formular, regexOUT))
                        {
                            Match mtemp = Regex.Match(m.Value, "{.+}");
                            Node n = Parse2Node(mtemp.Value.Substring(1, mtemp.Value.Length - 2), _SQLBuilder.Table, _SQLBuilder.Database, arrParam, _SQLBuilder.ParaValueList);
                            if (n != null)
                                _SQLBuilder.SelectedNodes.Add(n);
                        }
                    }



                }



            }
            return _SQLBuilder;
        }

        private static Node Parse2Node(string p, string tbl, string dtb, string[] arrparam, string[] arrvalue)
        {
            string key = "";
            string a = "";
            string c = "";

            if (Regex.IsMatch(p, regexKEY))
            {
                Match m = Regex.Match(p, regexKEY);
                key = m.Value.Replace("k=", string.Empty);
            }
            if (Regex.IsMatch(p, regexAGGREGATE))
            {
                Match m = Regex.Match(p, regexAGGREGATE);
                a = m.Value.Replace("a=", string.Empty);
            }
            if (Regex.IsMatch(p, regexC))
            {
                Match m = Regex.Match(p, regexC);
                c = m.Value.Replace("c=", string.Empty);
            }


            if (Regex.IsMatch(c, "{P}"))
            {
                int indexP = 0;
                if (int.TryParse(c.Replace("{P}", ""), out indexP))
                {
                    c = arrparam[indexP];
                }
            }
            foreach (Node _node in SchemaDefinition.GetDecorateTableByCode(tbl, dtb))
            {
                if (_node.Code.ToUpper() == key.ToUpper())
                {
                    return new Node(a, key,_node.Description, _node.FType, _node.NodeDesc);// "", "", "");//
                }
            }
            return null;
        }

        private static Filter Parse2Filter(string p, string tbl, string dtb, string[] arrparam, string[] arrvalue)
        {
            string vFilter = "";
            string f = "";
            string t = "";
            string fp = "";
            string tp = "";
            string vf = "";
            string vt = "";
            string ai = "";
            string op = "";

            if (Regex.IsMatch(p, regexKEY))
            {
                Match m = Regex.Match(p, regexKEY);
                vFilter = m.Value.Replace("k=", string.Empty);
            }
            if (Regex.IsMatch(p, regexFROM))
            {
                Match m = Regex.Match(p, regexFROM);
                f = m.Value.Replace("f=", string.Empty);
            }
            if (Regex.IsMatch(p, regexTO))
            {
                Match m = Regex.Match(p, regexTO);
                t = m.Value.Replace("t=", string.Empty);
            }
            if (Regex.IsMatch(p, regexIS))
            {
                Match m = Regex.Match(p, regexIS);
                ai = m.Value.Replace("i=", string.Empty);
            }
            if (Regex.IsMatch(p, regexOPERATOR))
            {
                Match m = Regex.Match(p, regexOPERATOR);
                op = m.Value.Replace("o=", string.Empty);
            }

            if (Regex.IsMatch(f, "{P}"))
            {
                int indexP = 0;
                if (int.TryParse(f.Replace("{P}", ""), out indexP))
                {
                    fp = f;
                    f = arrparam[indexP];
                    vf = arrvalue[indexP];
                }
            }
            else vf = f;
            if (Regex.IsMatch(t, "{P}"))
            {
                int indexP = 0;
                if (int.TryParse(t.Replace("{P}", ""), out indexP))
                {
                    tp = t;
                    t = arrparam[indexP];
                    vt = arrvalue[indexP];
                }
            }
            else vt = t;
            foreach (Node _node in SchemaDefinition.GetDecorateTableByCode(tbl, dtb))
            {
                if (_node.Code == vFilter)
                {
                    Filter resuslt = new Filter(new Node(vFilter, _node.Description), f, t, vf, vt, fp, tp);//
                    resuslt.IsNot = ai;
                    resuslt.Operate = op;
                    return resuslt;
                }
            }
            return null;
        }

        private static string AgregateN2Code(string N)
        {
            switch (N)
            {
                case "1":
                    return STR_SUM;
                case "2":
                    return STR_COUNT;
                case "3":
                    return STR_AVERAGE;
                case "4":
                    return STR_MINIMUM;
                case "5":
                    return STR_MAXIMUM;
                case "6":
                    return STR_DISTINCTSUM;
                case "7":
                    return STR_DISTINCTCOUNT;
                case "8":
                    return STR_DISTINCTAVERAGE;
                default:
                    return string.Empty;
            }

        }

        ///  <summary>
        ///  Parsing Formular without any parameter to SQL Builder.
        ///  Parameters must be replacesd by its value before using this method
        ///  SQLbuilder created by this method is used drectly. without using Dform
        ///  </summary>
        ///  <param name="ParseString"></param>
        ///  <param name="_SQLBuilder"></param>
        ///  <remarks></remarks>
        public static void Formular2SQLBuilderLight(string ParseString, ref SQLBuilder _SQLBuilder)
        {
            if (string.IsNullOrEmpty(ParseString))
            {
                return;
            }
            ParseString = ParseString.Replace(" ", string.Empty);
            ParseString = ParseString.Replace(@"\", "/");

            // Build Common -----------------------------------------------------------------------------------------------
            string vDatabase = Regex.Match(ParseString, regexDB).Value.ToString();
            string vTable = Regex.Match(ParseString, regexTable).Value.ToString();
            string vLedger = Regex.Match(ParseString, regexLedger).Value.ToString();

            _SQLBuilder.Database = vDatabase;
            _SQLBuilder.Table = vTable;
            _SQLBuilder.LedgerV = vLedger;

            // Build Filters ---------------------------------------------------------------------------------------------
            MatchCollection filterList = Regex.Matches(ParseString, @"(?<=\,[Kk]\=)[^\,]+");
            MatchCollection filterValues = Regex.Matches(ParseString, @"(?<=F\=)[^\/|\\]*(?=\,[Kk]\=)");

            int maxkey = Math.Min(filterList.Count, filterValues.Count);
            string key = string.Empty;
            for (int i = 0; i <= maxkey - 1; i++)
            {
                key = filterList[i].Value;
                if (key == "/LA/Ledger")
                {
                    // do nothing
                }
                else
                {
                    Filter f = new Filter(new Node(key, string.Empty)); // description is not required for light version

                    string fromRegex = @"(?<=F\=)[^\/|\\]*(?=\,K\=" + Regex.Replace(key, @"[\/]+|[\\]+", @"\/") + ")";
                    string toRegex = @"(?<=T\=)[^\/|\\]*(?=\,K\=" + Regex.Replace(key, @"[\/]+|[\\]+", @"\/") + ")";

                    f.ValueFrom = Regex.Match(filterValues[i].Value, @".+(?=\,)").Value;
                    f.ValueTo = Regex.Match(filterValues[i].Value, @"(?<=T\=).+").Value.ToString();
                    // get Field Type from dictionary

                    _SQLBuilder.Filters.Add(f);
                }
            }

            // Build Ouputs ---------------------------------------------------------------------------------------------
            MatchCollection outList = Regex.Matches(ParseString, @"(?<=[Oo]\=)[^\,.]+");

            foreach (System.Text.RegularExpressions.Match output in outList)
            {
                key = output.Value.ToString();
                string argRegex = @"[0-9](?=\,O\=" + Regex.Replace(key, @"[\/]+|[\\]+", @"\/") + ")";
                Node n = new Node(key, string.Empty);

                n.Agregate = AgregateN2Code(Regex.Match(ParseString, argRegex).Value);

                if (string.IsNullOrEmpty(n.Agregate) && _SQLBuilder.Mode != processingMode.Details)
                {
                    n.Agregate = STR_COUNT;
                }

                _SQLBuilder.SelectedNodes.Add(n);
            }

        }
        public static DataTable GetListIsNot()
        {
            DataTable dt = new DataTable("IsNot");
            dt.Columns.AddRange(new DataColumn[] { new DataColumn("Code", typeof(String)), new DataColumn("Description", typeof(String)) });
            DataRow row = dt.NewRow();
            row["Code"] = "Y";
            row["Description"] = "Is not";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "N";
            row["Description"] = "Is";
            dt.Rows.Add(row);
            return dt;
        }
        public static DataTable GetListOperator(string type)
        {
            DataTable dt = new DataTable("Operator");
            dt.Columns.AddRange(new DataColumn[] { new DataColumn("Code", typeof(String)), new DataColumn("Description", typeof(String)) });
            DataRow row = dt.NewRow();
            row["Code"] = "=";
            row["Description"] = "Equal To";
            dt.Rows.Add(row);




            row = dt.NewRow();
            row["Code"] = "BEGIN";
            row["Description"] = "Begins  with";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "END";
            row["Description"] = "Ends  with";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "CONTAIN";
            row["Description"] = "Contains";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "BETWEEN";
            row["Description"] = "Between";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "<";
            row["Description"] = "Less than";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = ">";
            row["Description"] = "Greater than";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "<=";
            row["Description"] = "Less than or Equal";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = ">=";
            row["Description"] = "Greater than or Equal";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "<>";
            row["Description"] = "Different from";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "in";
            row["Description"] = "In";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "SPACE";
            row["Description"] = "SPACE";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "EXISTS";
            row["Description"] = "EXISTS";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = "-";
            row["Description"] = "-";
            dt.Rows.Add(row);
            return dt;
        }


        public static DataTable GetListNumberAgregate()
        {
            DataTable dt = new DataTable("Agregate");
            dt.Columns.AddRange(new DataColumn[] { new DataColumn("Code", typeof(String)), new DataColumn("Description", typeof(String)) });
            DataRow row = dt.NewRow();
            row["Code"] = row["Description"] = STR_SUM;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = row["Description"] = STR_COUNT;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = row["Description"] = STR_AVERAGE;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = row["Description"] = STR_MINIMUM;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = row["Description"] = STR_MAXIMUM;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = row["Description"] = STR_DISTINCTSUM;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = row["Description"] = STR_DISTINCTCOUNT;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = row["Description"] = STR_DISTINCTAVERAGE;
            dt.Rows.Add(row);
            return dt;
        }
        public static DataTable GetListStringAgregate()
        {
            DataTable dt = new DataTable("Agregate");
            dt.Columns.AddRange(new DataColumn[] { new DataColumn("Code", typeof(String)), new DataColumn("Description", typeof(String)) });
            DataRow row = dt.NewRow();
            row["Code"] = row["Description"] = STR_COUNT;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = row["Description"] = STR_MINIMUM;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = row["Description"] = STR_MAXIMUM;
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Code"] = row["Description"] = STR_DISTINCTCOUNT;
            dt.Rows.Add(row);

            return dt;
        }

    }


}

