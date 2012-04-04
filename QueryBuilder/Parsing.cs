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


        // Public Shared Sub Formular2SQLBuilder(ByVal ParseString As String, ByRef _SQLBuilder As SQLBuilder)
        //     If String.IsNullOrEmpty(ParseString) Then Exit Sub

        //     ParseString = Regex.Replace(ParseString, ".*(?=TT_XLB_EB)", String.Empty)

        //     Dim vTable As String      ' ex LA,CA
        //     Dim vDatabase As String

        //     Dim vFilter As String     'ex part of TTformular , contains filters
        //     Dim FromTo() As String = New String() {}
        //     Dim vParamsString As String   ' the string, contains parameters in TTformular
        //     Dim vParameter() As String = New String() {} ' array of all parameters ex .$H$1,$G14,$H14,J$11,J$11
        //     Dim vparacount As Integer = 0
        //     Dim vPosition As String  ' address of the formular

        //     'Return Database, table
        //     vDatabase = Regex.Match(ParseString, "(?<=0\,2\,)[A-Z0-9\{\}]{1,5}").Value.ToString
        //     vTable = Regex.Match(ParseString, "(?<=0\,2\,[A-Z0-9\{\}]{1,}\,)[A-Z0-9\{\}]{1,}").Value.ToString

        //     _SQLBuilder.Table = vTable

        //     vFilter = Regex.Match(ParseString, "(?<=\,K\=)[^\,,.]+").Value.ToString
        //     Dim i As Integer = 0
        //     Dim n As Integer = 0
        //     n = Regex.Matches(ParseString, "F.+?,K").Count

        //     'fill FromTo array
        //     If n > 0 Then
        //         ReDim FromTo(n)
        //         For Each ft As Match In Regex.Matches(ParseString, "F.+?,K")
        //             i = i + 1
        //             FromTo(i) = ft.Value.ToString

        //         Next
        //     End If
        //     n = 0
        //     i = 0

        //     ' get string , contains parameters
        //     vParamsString = Regex.Match(ParseString, "\" & Chr(34) & "\,.+?\)").Value.ToString

        //     'fill to parameter Array
        //     If Not String.IsNullOrEmpty(vParamsString) Then
        //         vParamsString = Mid(vParamsString, 3)
        //         vParamsString = Mid(vParamsString, 1, Len(vParamsString) - 1)
        //         vParamsString = vParamsString & ","  ' them dau , cho de xu ly
        //         n = Regex.Matches(vParamsString, ".*?,").Count ' cac tham so
        //         If n > 0 Then
        //             ReDim vParameter(n - 1)  'tham so dau tien la vi tri cua cong thuc
        //             For Each p As Match In Regex.Matches(vParamsString, ".*?,")
        //                 i = i + 1
        //                 If i = 1 Then
        //                     vPosition = p.Value.ToString.Replace(",", String.Empty)
        //                     _SQLBuilder.Pos = vPosition
        //                 Else
        //                     vParameter(i - 1) = p.Value.ToString.Replace(",", String.Empty)
        //                 End If

        //             Next
        //         End If
        //     End If

        //     If vDatabase.Contains("{P}") Then
        //         _SQLBuilder.DatabaseP = vDatabase
        //         _SQLBuilder.Database = vParameter(vDatabase.Replace("{P}", String.Empty))
        //         vparacount = vparacount + 1
        //         _sqlBuilder.Database = _SQLBuilder.ParaValueList(vparacount)
        //     Else
        //         _SQLBuilder.Database = vDatabase
        //         _sqlBuilder.Database = vDatabase

        //     End If

        //     i = 0

        //     Dim vf As String
        //     Dim vt As String
        //     Dim vf1 As String = ""
        //     Dim vt1 As String = ""
        //     Dim filterf As String
        //     Dim filtert As String

        //     'identifying filters
        //     For Each m As Match In Regex.Matches(ParseString, "(?<=\,K\=)[^\,,.]+")
        //         i = i + 1
        //         vFilter = m.Value.ToString

        //         vf = Regex.Match(FromTo(i), "F.+?,").Value.ToString
        //         If Not String.IsNullOrEmpty(vf) Then
        //             vf = Mid(vf, 3)
        //             vf = Mid(vf, 1, Len(vf) - 1)
        //         End If

        //         vt = Regex.Match(FromTo(i), "T.+?,").Value.ToString
        //         If Not String.IsNullOrEmpty(vt) Then
        //             vt = Mid(vt, 3)
        //             vt = Mid(vt, 1, Len(vt) - 1)
        //         End If
        //         filterf = ""
        //         filtert = ""

        //         If Regex.IsMatch(vf, "{P}") Then
        //             filterf = vParameter(Mid(vf, 4))
        //         End If

        //         If Regex.IsMatch(vt, "{P}") Then
        //             filtert = vParameter(Mid(vt, 4))
        //         End If

        //         If Not String.IsNullOrEmpty(vFilter) Then
        //             vFilter = Mid(vFilter, 2)
        //             If String.IsNullOrEmpty(filterf) Then filterf = vf
        //             If String.IsNullOrEmpty(filtert) Then filtert = vt
        //             If vFilter.ToUpper = "LA/LEDGER" Then ' ledger lam rieng

        //                 If Regex.IsMatch(vf, "{P}") Then
        //                     _SQLBuilder.LedgerP = vf
        //                     _SQLBuilder.Ledger = vParameter(Mid(vf, 4))
        //                     vparacount = vparacount + 1
        //                     _SQLBuilder.LedgerV = _SQLBuilder.ParaValueList(vparacount)

        //                 Else
        //                     _SQLBuilder.Ledger = vf
        //                     _SQLBuilder.LedgerV = vf
        //                 End If
        //             Else
        //                 If Regex.IsMatch(vf, "{P}") Then
        //                     vparacount = vparacount + 1
        //                     vf1 = _SQLBuilder.ParaValueList(vparacount) 'gia tri
        //                 Else
        //                     vf1 = vf
        //                 End If

        //                 If Regex.IsMatch(vt, "{P}") Then
        //                     vparacount = vparacount + 1
        //                     vt1 = _SQLBuilder.ParaValueList(vparacount) 'gia tri
        //                 Else
        //                     vt1 = vt
        //                 End If

        //                 For Each _node As Node In SchemaDefinition.GetDecorateTableByCode(vTable, _sqlBuilder.Database)
        //                     If _node.Code.Contains(vFilter) Then
        //                         _SQLBuilder.Filters.Add(New Filter(New Node(vFilter, _node.Description), filterf, filtert, vf1, vt1, vf, vt))

        //                         ' _SQLBuilder.SelectedNodes.Add(New Node(vOutputAgr(i), Output, _node.Description, _node.FType))
        //                     End If
        //                 Next
        //                 ' _SQLBuilder.Filters.Add(New Filter(New Node(vFilter, vFilter), filterf, filtert, vf1, vt1, vf, vt))

        //             End If


        //         End If
        //     Next

        //     Dim Output As String
        //     Dim vOutputAgr() As String

        //     n = Regex.Matches(ParseString, "E\=.+?,").Count
        //     i = 0
        //     If n > 0 Then
        //         ReDim vOutputAgr(n)
        //         For Each oe As Match In Regex.Matches(ParseString, "E\=.+?,")
        //             i = i + 1
        //             vOutputAgr(i) = Mid(oe.Value.ToString, 3, 1)
        //         Next oe
        //         i = 0
        //         For Each o As Match In Regex.Matches(ParseString, "O\=.+?,")
        //             Output = o.Value.ToString
        //             If Not String.IsNullOrEmpty(Output) Then
        //                 Output = Output.Replace(",", String.Empty)
        //             End If

        //             i = i + 1

        //             vOutputAgr(i) = AgregateN2Code(vOutputAgr(i))

        //             If Not String.IsNullOrEmpty(Output) Then
        //                 Output = Output.Replace("O=/", String.Empty)
        //                 For Each _node As Node In SchemaDefinition.GetDecorateTableByCode(vTable, _sqlBuilder.Database)
        //                     'If Regex.IsMatch(_node.Code, Output & "$") Then
        //                     If _node.Code.ToUpper = Output.ToUpper Then
        //                         _SQLBuilder.SelectedNodes.Add(New Node(vOutputAgr(i), Output, _node.Description, _node.FType))
        //                         Exit For
        //                     End If
        //                 Next
        //             End If

        //         Next
        //     End If

        // End Sub

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

                    string fromRegex = @"(?<=F\=)[^\/|\\]*(?=\,K\=" + Regex.Replace(key, @"\/]+|[\\]+", @"\/") + ")";
                    string toRegex = @"(?<=T\=)[^\/|\\]*(?=\,K\=" + Regex.Replace(key, @"\/]+|[\\]+", @"\/") + ")";

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
                string argRegex = @"[0-9](?=\,O\=" + Regex.Replace(key, @"\/]+|[\\]+", @"\/") + ")";
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

