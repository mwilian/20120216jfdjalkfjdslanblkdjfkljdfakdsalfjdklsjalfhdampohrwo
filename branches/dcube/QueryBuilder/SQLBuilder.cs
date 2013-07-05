using System.Text.RegularExpressions;

using System.Data.OleDb;
//using SatResources;


using System;
using System.Collections.Generic;
using System.Data;
//using System.Drawing;
//using System.Windows.Forms;


namespace QueryBuilder
{
    //  Class SQLBuilder
    [Serializable()]
    public class SQLBuilder
    {


        // Public Output() As Double
        // Public NoOfOutput As Integer
        public const string DocumentFolder = "TVC-QD";
        private processingMode _Mode = processingMode.Details;
        private bool _headerVisible = false;
        private static string periodParterm = @"[1-9]{1}[0-9]{3}0[0-9]{1}[1-9]{1}";
        //  Property HeaderVisible
        public bool HeaderVisible
        {
            get
            {
                return _headerVisible;
            }
            set
            {
                _headerVisible = value;
            }
        }
        //  Property Mode
        public processingMode Mode
        {
            get
            {
                return _Mode;
            }
        }
        private static bool _SQLDebugMode = false;



        public static void SetConnection(string val)
        {
            CoreCommonControl.SetConnection(val);
        }
        public static string GetConnection()
        {
            return CoreCommonControl.GetConnection();
        }
        string _strConnectDes = "";

        public string StrConnectDes
        {
            get { return _strConnectDes; }
            set { _strConnectDes = value; }
        }
        public static bool SQLDebugMode
        {
            get { return _SQLDebugMode; }
            set { _SQLDebugMode = value; }
        }
        //    Public ParseReverse As String
        public string[] ParaValueList = new string[51];

        public SQLBuilder(processingMode ProcMode)
        {
            for (int i = 0; i < ParaValueList.Length; i++)
                ParaValueList[i] = "";
            _Mode = ProcMode;
        }

        public SQLBuilder(SQLBuilder sqlBuilder)
        {
            for (int i = 0; i < ParaValueList.Length; i++)
                ParaValueList[i] = "";
            _Mode = sqlBuilder.Mode;
            _Database = sqlBuilder.Database;
            _databaseP = sqlBuilder.DatabaseP;
            _DatabaseV = sqlBuilder.DatabaseV;
            _ledger = sqlBuilder.Ledger;
            _LedgerP = sqlBuilder.LedgerP;
            _ledgerV = sqlBuilder.LedgerV;
            _Table = sqlBuilder.Table;
            _Pos = sqlBuilder.Pos;

            foreach (Node node in sqlBuilder.SelectedNodes)
            {
                Node newNode = node.CloneNode();
                SelectedNodes.Add(newNode);
            }
            foreach (Filter filter in sqlBuilder.Filters)
            {
                Filter newNode = new Filter(filter.Node, filter.FilterFrom, filter.FilterTo, filter.ValueFrom, filter.ValueTo, filter.FilterFromP, filter.FilterToP);
                newNode.Operate = filter.Operate;
                newNode.IsNot = filter.IsNot;
                Filters.Add(newNode);
            }
        }

        #region '" Constants and Regex "'
        /* L:40 */
        private const string Tab = "   ";
        private const string Space = " ";
        private const string Comma = ",";
        private const string LastComma = @"[\,]$";

        private const string STR_SELECT = "SELECT ";
        private const string STR_FROM = " FROM ";
        private const string STR_GROUPBY = " GROUP BY ";
        private const string STR_WHERE = " WHERE ";
        private const string STR_ORDERBY = " ORDER BY ";
        private const string STR_ASC = " ASC ";
        private const string STR_DESC = " DESC ";
        #endregion

        #region '" Members "'
        private string _connID = string.Empty;

        private System.ComponentModel.BindingList<Node> _nodes = new System.ComponentModel.BindingList<Node>();
        private System.ComponentModel.BindingList<Filter> _filters = new System.ComponentModel.BindingList<Filter>();

        private string _Database = string.Empty;
        private string _ledger = "A";
        private string _DatabaseV = string.Empty;
        private string _ledgerV = "A";
        private string _databaseP = string.Empty;
        private string _LedgerP = string.Empty;
        private string _Table = string.Empty;
        private string _Pos = "A1";


        private string SelectClause = string.Empty;
        private string FromClause = string.Empty;
        private string GroupByClause = string.Empty;
        private string WhereClause = string.Empty;
        private string OrderByClause = string.Empty;

        public string ConnID
        {
            get { return _connID; }
            set
            {
                //if (_connID != value)
                //    SchemaDefinition.JoinsDictionary = null;
                _connID = value;
            }
        }

        //  Property Database
        public string Database
        {
            get
            {
                return _Database;
            }
            set
            {
                _Database = value;
                if (!(_eval == null))
                {
                    _eval.Invoke(_Database, ref _databaseP, ref _DatabaseV);
                }
                else
                {
                    _DatabaseV = value;
                }
            }
        }
        //  Property Ledger
        public string Ledger
        {
            get
            {
                return _ledger;
            }
            set
            {
                _ledger = value;
                if (!(_eval == null))
                {

                    _eval.Invoke(_ledger, ref _LedgerP, ref _ledgerV);
                }
            }
        }

        //  Property DatabaseV
        public string DatabaseV
        {
            get
            {
                return _DatabaseV;
            }
            set
            {
                _DatabaseV = value;
            }
        }
        //  Property LedgerV
        public string LedgerV
        {
            get
            {
                return _ledgerV;
            }
            set
            {
                _ledgerV = value;
            }
        }
        //  Property DatabaseP
        public string DatabaseP
        {
            get
            {
                return _databaseP;
            }
            set
            {
                _databaseP = value;

            }
        }
        //  Property LedgerP
        public string LedgerP
        {
            get
            {
                return _LedgerP;
            }
            set
            {
                _LedgerP = value;
            }
        }

        //  Property Table
        public string Table
        {
            get
            {
                return _Table;
            }
            set
            {
                if (_Table == value)
                {
                    return;
                }
                _Table = value;
            }
        }

        //  Property Pos
        public string Pos
        {
            get
            {
                return _Pos;
            }
            set
            {
                _Pos = value;
            }
        }

        //  Property SelectedNodes
        public System.ComponentModel.BindingList<Node> SelectedNodes
        {
            get
            {
                return _nodes;
            }
            set
            {
                _nodes = value;
            }
        }

        //  Property Filters
        public System.ComponentModel.BindingList<Filter> Filters
        {
            get
            {
                return _filters;
            }
            set
            {
                _filters = value;
            }

        }
        #endregion

        //  Method AddOutputNodez
        public void AddOutputNode(Node _node)
        {
            if (_node == null)
            {
                return;
            }
            switch (Mode)
            {
                case processingMode.Balance:
                    if (_node.FType == "N")
                    {
                        _node.Agregate = Parsing.STR_SUM;
                    }
                    else { _node.Agregate = Parsing.STR_COUNT; }
                    break;
                case processingMode.Link:
                    _node.Agregate = Parsing.STR_MAXIMUM;
                    break;
            }

            _nodes.Add(_node);
        }


        #region '" Building Formular from sqlBuilder Object "'

        //  Method AddNextCell
        private string AddNextCell(string aPos)
        {
            string tmp = null;
            tmp = Regex.Match(aPos, "[A-Z]{1,4}").Value;
            if (!(string.IsNullOrEmpty(tmp)))
            {
                if (tmp.Length > 1)
                {
                    tmp = tmp.Substring(2);
                    //if (tmp != "Z")
                    //{
                    //    tmp.Substring(2) = System.Convert.ToString(System.Convert.ToChar(System.Convert.ToInt32(char.Parse(tmp.Substring(2))) + 1));
                    //}
                    //else
                    //{
                    //    tmp.Substring(2) = "A";
                    //    tmp.Substring(1) = System.Convert.ToString(System.Convert.ToChar(System.Convert.ToInt32(char.Parse(tmp.Substring(1))) + 1));
                    //} 
                }
                else
                {
                    if (tmp == "Z")
                    {
                        tmp = "AA";
                    }
                    else
                    {
                        tmp = Convert.ToString(Convert.ToChar(Convert.ToInt32(char.Parse(tmp)) + 1));
                    }
                }
            }

            return tmp + Regex.Match(aPos, "[0-9]{1,9}").Value;
        }


        // TRANSNOTUSED: Private Method NZDBL

        //		private double NZDBL( object SOURCE, double VALUE_IF_NULL ) 
        //		{ 
        //			if (   System.Convert.IsDBNull( SOURCE ) )
        //			{ 
        //				return VALUE_IF_NULL; 
        //			} 
        //			return System.Convert.ToDouble( SOURCE ); 
        //		} 
        //
        // TRANSWARNING: Automatically generated because of optional parameter(s) 
        // TRANSNOTUSED: Private Method NZDBL

        //		private double NZDBL( object SOURCE ) 
        //		{ 
        //			return NZDBL( SOURCE, 0 ); } 
        //

        ///  <summary>
        ///   Build Select Statement . return TTION part of Formula
        ///   includes Outputs and all parameters
        ///   ex:   E=1,O=/LA/AMOUNT,",J14,$H$1,$G14,$H14,J$11,J$11)
        ///   fill TTselect with  E=1,O=/LA/AMOUNT,",
        ///   fill TTParam with J14,$H$1,$G14,$H14,J$11,J$11)
        ///  </summary>
        ///  <param name="Pos">Cell Address - A1-A2 B5 .... Position of the output Cell</param>
        ///  <remarks></remarks>
        private string BuildFormular_SELECT(string Pos, ref string TTselect, ref string TTPara, ref int MaxPara)
        {
            string aPos = Pos;
            string tmp = string.Empty;
            string ARG = string.Empty;
            if (_nodes.Count <= 0)
            {
                return string.Empty;
            }

            for (int i = 0; i <= _nodes.Count - 1; i++)
            {

                switch (_nodes[i].Agregate)
                {
                    case Parsing.STR_SUM:
                        ARG = "1";
                        break;
                    case Parsing.STR_COUNT:
                        ARG = "2";
                        break;
                    case Parsing.STR_AVERAGE:
                        ARG = "3";
                        break;
                    case Parsing.STR_MINIMUM:
                        ARG = "4";
                        break;
                    case Parsing.STR_MAXIMUM:
                        ARG = "5";
                        break;
                    case Parsing.STR_DISTINCTSUM:
                        ARG = "6";
                        break;
                    case Parsing.STR_DISTINCTCOUNT:
                        ARG = "7";
                        break;
                    case Parsing.STR_DISTINCTAVERAGE:
                        ARG = "8";
                        break;
                    default:
                        ARG = string.Empty;
                        break;
                }


                if (i == 0)
                {
                    TTselect += string.Format("E={0},O=/{1},", ARG, _nodes[i].Code);
                    //  TTselect + "E=" + ARG + "," + "O=/" + _nodes(i).Code & ","
                }
                else
                {
                    MaxPara = MaxPara + 1;
                    //   TTselect = TTselect + "C={P}" & MaxPara & "," + "E=" + ARG + "," + "O=/" + _nodes(i).Code & ","
                    TTselect += string.Format("C={{P}}{0},E={1},O=/{2},", MaxPara, ARG, _nodes[i].Code);

                    // ---
                    aPos = AddNextCell(aPos);
                    // TTPara = TTPara + aPos + ","
                    TTPara += string.Concat(aPos, ",");

                }
            }
            return string.Concat(TTselect, TTPara);

        }


        ///  <summary>
        ///  return Filter part of TT formular
        ///  F=A,K=/LA/Ledger,F={P}2,T={P}3,K=/LA/ACCNT_CODE,F={P}4,T={P}5,K=/LA/PERIOD
        ///  </summary>
        ///  <param name="Pos">Cell Address - A1-A2 B5 .... Position of the output Cell</param>
        ///  <remarks></remarks>
        private void BuildFormular_WHERE(string Pos, ref string TTWhere, ref string TTPara, ref int MaxPara)
        {

            // Where by clause
            // TTwhere = "F=" & aLedger & ",K=/LA/Ledger,"
            if (Regex.IsMatch(DatabaseP, @"\{P\}"))
            {
                TTPara += string.Concat(Database, ",");
                MaxPara = MaxPara + 1;
            }
            // database thi khong can

            if (Regex.IsMatch(LedgerP, @"\{P\}"))
            {
                TTPara += string.Concat(Ledger, ",");
                MaxPara = MaxPara + 1;
                TTWhere += string.Format("F={{P}}{0},K=/LA/Ledger,", MaxPara);
            }
            else
            {
                TTWhere += string.Format("F={0},K=/LA/Ledger,", LedgerV);
            }

            for (int i = 0; i <= _filters.Count - 1; i++)
            {
                if (!(string.IsNullOrEmpty(_filters[i].ValueFrom)))
                {
                    if (Regex.IsMatch(_filters[i].FilterFromP, @"\{P\}"))
                    {
                        TTPara += string.Concat(_filters[i].FilterFrom, ",");
                        MaxPara = MaxPara + 1;
                        TTWhere += string.Format("F={{P}}{0},", MaxPara);
                    }
                    else
                    {
                        TTWhere += string.Format("F={0},", _filters[i].ValueFrom.Replace("'", string.Empty));
                    }
                }

                if (!(string.IsNullOrEmpty(_filters[i].ValueTo)))
                {
                    // TTwhere = TTwhere + "T=" & _filters(i).ValueTo & ","
                    if (Regex.IsMatch(_filters[i].FilterToP, @"\{P\}"))
                    {
                        TTPara += string.Concat(_filters[i].FilterTo, ",");
                        MaxPara = MaxPara + 1;
                        TTWhere += string.Format("T={{P}}{0},K=/{1},", MaxPara, _filters[i].Node);
                    }
                    else
                    {
                        TTWhere += string.Format("T={0},K=/{1},", _filters[i].ValueTo.Replace("'", string.Empty), _filters[i].Node);
                    }

                }
                else if (!(string.IsNullOrEmpty(_filters[i].ValueFrom)))
                {
                    TTWhere += string.Format("K=/{0},", _filters[i].Node);
                }
            }
        }


        ///  <summary>
        ///  Return TTION Formula to Position POS
        ///  </summary>
        ///  <param name="Pos"></param>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public string BuildTTformula(string Pos)
        {

            string TTwhere = "";
            string TTselect = "";
            string TTpara = "";
            int MaxPara = 0; //  total number of al filters and all outputs

            BuildFormular_WHERE(Pos, ref TTwhere, ref TTpara, ref MaxPara);
            BuildFormular_SELECT(Pos, ref TTselect, ref TTpara, ref MaxPara);

            string s = null;
            string Db = null;
            if (Regex.IsMatch(DatabaseP, @"\{P\}"))
            {
                Db = "{P}1";
            }
            else
            {
                Db = DatabaseV;
            }
            // s = "=TT_XLB_EB(" & Chr(34) & " 0,2," & Db & ",LA,V=4," & TTwhere & TTselect & Chr(34) & "," & Pos & "," & TTpara
            s = String.Format("=TT_XLB_EB({0} 0,2,{1},{2},V=4,{3}{4}{0},{5},{6}", Convert.ToChar(34), Db, Table, TTwhere, TTselect, Pos, TTpara);
            s = s.Substring(0, s.Length - 1) + ")";//Strings.Mid(s, 1, s.Length - 1) + ")";
            return s;
        }

        //  Method BuildTTformula
        public string BuildTTformula()
        {
            return BuildTTformula(Pos);
        }


        ///  <summary>
        ///  Return TTION Formula to Position POS
        ///  </summary>
        ///  <param name="Pos"></param>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public string BuildTVCformula(string Pos)
        {
            int indexPara = 1;
            string param = "";
            string result = "TVC_QUERY(\"{";
            if (Regex.IsMatch(DatabaseP, @"\{P\}"))
            {
                result += String.Format("dtb={P}{0};", indexPara);
                param += "," + Database;
                indexPara++;
            }
            else
                result += String.Format("dtb={0};", _DatabaseV);

            result += String.Format("tbl={0};", _Table);

            if (Regex.IsMatch(_LedgerP, @"\{P\}"))
            {
                result += String.Format("ldg={P}{0};", indexPara);
                param += "," + Ledger;
                indexPara++;
            }
            else
                result += String.Format("ldg={0};", _ledgerV);
            foreach (Filter f in Filters)
            {
                result += "fil={";
                if (Regex.IsMatch(f.FilterFromP, @"\{P\}"))
                {
                    result += String.Format("f={P}{0};", indexPara);
                    param += "," + f.FilterFrom;
                    indexPara++;
                }
                else
                {
                    result += String.Format("f={0};", f.ValueFrom);
                }
                if (Regex.IsMatch(f.FilterToP, @"\{P\}"))
                {
                    result += String.Format("t={P}{0};", indexPara);
                    param += String.Format(",{0}", f.FilterTo);
                    indexPara++;
                }
                else
                {
                    result += string.Format("t={0};", f.ValueTo);
                }
                result += string.Format("o={0};", f.Operate);
                result += String.Format("i={0};", f.IsNot);
                result += String.Format("k={0};", f.Node.Code);
                result += "};";
            }
            foreach (Node n in SelectedNodes)
            {
                result += "out={";
                result += String.Format("a={0};", n.Agregate);
                result += String.Format("k={0};", n.Code);
                result += "};";
            }

            result += String.Format("}}\",{0}{1}", Pos, param);
            result += ")";
            return result;
        }

        //  Method BuildTTformula
        public string BuildTVCformula()
        {
            return BuildTVCformula(Pos);
        }

        //  Method Convert2XML
        public string Convert2XML()
        {
            try
            {
                return "'" + BuildTTformula();

            }
            catch (Exception ex)
            {
                throw new ArgumentException(ex.Message);
            }
            return string.Empty;
        }


        #endregion

        #region '" SQL Scripts Processing "'

        //  Method BuildSELECT
        private void BuildSELECT()
        {

            if (SelectedNodes.Count == 0)
            {
                SelectClause = "SELECT * ";
                return;
            }

            SelectClause = string.Empty;

            System.Text.StringBuilder selectBuilder = new System.Text.StringBuilder(STR_SELECT);

            if (_nodes.Count <= 0)
            {
                return;
            }
            selectBuilder.Append(Environment.NewLine);
            for (int i = 0; i <= _nodes.Count - 1; i++)
            {
                if (_nodes[i].Expresstion == "")
                    selectBuilder.AppendFormat("{0}{1} {2},{3}", Tab, _nodes[i].AgregateMe(), String.Format(" as [{0}]", _nodes[i].Description), Environment.NewLine);
                //selectBuilder.AppendFormat("{0}{1} {2},{3}", Tab, _nodes[i].AgregateMe(), " as [" + _nodes[i].MyCode + "]", System.Environment.NewLine);
                //selectBuilder.AppendFormat("{0}{1} {2},{3}", Tab, _nodes[i].AgregateMe(), _nodes[i].MyAlias(), System.Environment.NewLine);
                // SelectClause += String.Concat(Tab, _nodes(i).AgregateMe, _nodes(i).MyAlias, Comma, System.Environment.NewLine)
            }
            // eliminate comma at the end
            SelectClause = Regex.Replace(selectBuilder.ToString().TrimEnd(), LastComma, Environment.NewLine);

        }


        //  Method BuildFROM
        private void BuildFROM()
        {
            // Me.DatabaseV = Me.Database
            //if (string.IsNullOrEmpty(this.Database))
            //{
            //    throw new ArgumentException("Database is empty");
            //}
            //if (string.IsNullOrEmpty(this.Table))
            //{
            //    throw new ArgumentException("Unknown Table source");
            //}

            FromClause = string.Empty;
            //if (_nodes.Count <= 0)
            //{
            //    return;
            //}

            // the list which registers the join clauses which is included
            List<string> RegisteredAliasList = new System.Collections.Generic.List<string>();

            // add all selected nodes to the sorted list
            SortedList<string, Node> sortedJoins = new System.Collections.Generic.SortedList<string, Node>();
            foreach (Node _node in _nodes)
            {
                if (!(sortedJoins.ContainsValue(_node)))
                {
                    sortedJoins.Add(_node.Code, _node);
                }
            }
            // add all filter nodes to the sorted list as well
            foreach (Filter _filter in _filters)
            {
                if (!(sortedJoins.ContainsValue(_filter.Node)))
                {
                    sortedJoins.Add(_filter.Node.Code, _filter.Node);
                }
            }

            FromClause += STR_FROM + Environment.NewLine;

            string currentFamily = string.Empty;
            for (int i = 0; i <= sortedJoins.Count - 1; i++)
            {
                FromClause += sortedJoins.Values[i].MyJoinClause(Database, currentFamily, ref RegisteredAliasList);

                currentFamily = sortedJoins.Values[i].MyFamily;
            }

        }


        //  Method BuildWhere
        private void BuildWhere()
        {

            //foreach (Filter f in this._filters)
            //{
            //    if (string.IsNullOrEmpty(f.FilterFrom))
            //    {
            //        throw new ArgumentException("Missing filter Value for " + f.Node.Code);
            //    }
            //}

            WhereClause = string.Empty;

            System.Text.StringBuilder WhereBuilder = new System.Text.StringBuilder();
            object filterItemCode = string.Empty;

            List<Filter> sortList = new System.Collections.Generic.List<Filter>();
            foreach (Filter f in _filters)
            {
                if (f.Operate == "SPACE" || f.Operate == "EXISTS" || (!string.IsNullOrEmpty(f.FilterFrom) && !Regex.IsMatch(f.Code, @"^@") && f.Code.Substring(0, 2) != "__"))
                    sortList.Add(f);

            }

            /* L:424 */
            sortList.Sort();
            /* L:424 */
            for (int i = 0; i <= sortList.Count - 1; i++)
            {
                if (filterItemCode.ToString() != sortList[i].Node.ToString())
                {
                    if (i == 0)
                    {
                        WhereBuilder.AppendFormat("{0}{1}{2}", Tab, sortList[i].MyWhereClause(), Environment.NewLine);
                    }
                    else
                    {
                        WhereBuilder.AppendFormat(" ){0} AND ({0} {1}{0}", Environment.NewLine, sortList[i].MyWhereClause());
                    }
                    filterItemCode = sortList[i].Node.ToString();
                }
                else
                {
                    // OR
                    WhereBuilder.AppendFormat(" OR {0}{1}", sortList[i].MyWhereClause(), Environment.NewLine);
                }


            }

            if (WhereBuilder.Length == 0)
            {
                return;
            }

            WhereClause = string.Format("WHERE {0} ({1}){0}", Environment.NewLine, WhereBuilder);

        }


        //  Method BuildGROUPBY
        private void BuildGROUPBY()
        {
            if (Mode == processingMode.Balance || Mode == processingMode.Link)
            {
                foreach (Node o in SelectedNodes)
                {
                    //if (string.IsNullOrEmpty(o.Agregate))
                    //{
                    //    throw new ArgumentException("Invalid Agregate");
                    //}
                }
            }

            GroupByClause = string.Empty;
            // Group by clause

            for (int i = 0; i <= _nodes.Count - 1; i++)
            {
                if (string.IsNullOrEmpty(_nodes[i].Agregate))
                {
                    if (_nodes[i].Expresstion == "")
                        GroupByClause += Tab + _nodes[i].FormatMe() + Comma + Environment.NewLine;
                }
            }

            if (string.IsNullOrEmpty(GroupByClause))
            {
                return;
            }

            GroupByClause = STR_GROUPBY + Space + Environment.NewLine + GroupByClause;
            // eliminate comma at the end
            GroupByClause = Regex.Replace(GroupByClause.Trim(), LastComma, Environment.NewLine);


        }
        private void BuildOREDERBY()
        {
            OrderByClause = string.Empty;
            // Group by clause

            for (int i = 0; i <= _nodes.Count - 1; i++)
            {
                if (_nodes[i].Sort != "")
                {
                    if (_nodes[i].Expresstion == "")
                    {
                        if (_nodes[i].Agregate != "")
                        {
                            string sort = _nodes[i].Sort.Trim() == STR_ASC.Trim() ? STR_ASC : STR_DESC;
                            OrderByClause += String.Format("{0}{1}({2}){3}{4}{5}", Tab, _nodes[i].Agregate, _nodes[i].FormatMe(), Space, sort, Comma);
                        }
                        else
                        {
                            string sort = _nodes[i].Sort.Trim() == STR_ASC.Trim() ? STR_ASC : STR_DESC;
                            OrderByClause += Tab + _nodes[i].FormatMe() + Space + sort + Comma;
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(OrderByClause))
            {
                return;
            }

            OrderByClause = STR_ORDERBY + Space + Environment.NewLine + OrderByClause;
            // eliminate comma at the end
            OrderByClause = Regex.Replace(OrderByClause.Trim(), LastComma, Environment.NewLine);
        }

        //  Method BuildSQL
        public string BuildSQL(string XXX, string aLedger)
        {
            BuildSELECT();
            BuildFROM();
            BuildWhere();
            BuildGROUPBY();
            BuildOREDERBY();

            if (string.IsNullOrEmpty(aLedger) | aLedger == "A")
            {
                aLedger = "L";
            }

            // Combine all together
            string retSQL = Regex.Replace(string.Concat(SelectClause, FromClause, WhereClause, GroupByClause, OrderByClause), "XXX", XXX);
            foreach (Filter x in Filters)
            {
                if (x.Node.MyCode.Substring(0, 1) == "@")
                    retSQL = retSQL.Replace(x.Code, x.ValueFrom);
            }
            //retSQL = Regex.Replace(retSQL, @"\/|\\", "_");

            return Regex.Replace(retSQL, "LLL", aLedger);
        }



        // TRANSWARNING: Automatically generated because of optional parameter(s) 
        //  Method BuildSQL
        public string BuildSQL(string XXX)
        {
            return BuildSQL(XXX, "A");
        }


        //  Method BuildSQL
        public string BuildSQL()
        {
            return BuildSQL(DatabaseV, LedgerV);
        }


        //  Method BuildWHEREClause
        public string BuildWHEREClause()
        {
            BuildWhere();
            return WhereClause;
        }


        //  Method ValidationCheck
        public void ValidationCheck()
        {
            //if (string.IsNullOrEmpty(this.Database))
            //{
            //    // throw new InvalidExpressionException(Localization.ResStr("DatabaseIsEmpty"));
            //    throw new InvalidExpressionException("DatabaseIsEmpty");
            //}
            DBInfo transTemp0 = null;
            if (!(DBInfoList.ContainsCode(Database, ref transTemp0)))
            {
                // throw new InvalidExpressionException(Localization.ResStr("DatabaseDoesNotExist"));
                throw new InvalidExpressionException("DatabaseDoesNotExist");
            }
            foreach (Filter flt in Filters)
            {
                if (string.IsNullOrEmpty(flt.FilterFrom))//|| string.IsNullOrEmpty(flt.FilterTo)
                {
                    //  throw new InvalidExpressionException(Localization.ResStr("FilterValueIsMissing"));
                    throw new InvalidExpressionException("FilterValueIsMissing");
                }
            }
            if (SelectedNodes.Count == 0)
            {
                // throw new InvalidExpressionException(Localization.ResStr("OutputIsEmpty"));
                throw new InvalidExpressionException("OutputIsEmpty");
            }
            switch (Mode)
            {
                case processingMode.Balance:

                    break;
                case processingMode.Details:

                    break;
                case processingMode.Link:

                    break;
            }

        }


        #endregion

        #region '"Evaluate Cell Addresses"'

        [NonSerialized()]
        public Delegation.EvaluateCell _eval;

        #endregion

        #region Method
        public static SQLBuilder LoadSQLBuilderFromDataBase(SQLBuilder kq, string qd_id, string database, string table)
        {
            string sErr = "";
            kq.Filters.Clear();
            CoreQD_SCHEMAControl schCtr = new CoreQD_SCHEMAControl();
            if (table != "")
                kq.Table = table.Trim();
            else
            {
                CoreQDControl qdCtr = new CoreQDControl();
                CoreQDInfo qdInf = qdCtr.Get_CoreQD(database, qd_id, ref sErr);
                table = kq.Table = qdInf.ANAL_Q0;
            }
            CoreQD_SCHEMAInfo schInf = schCtr.Get(database, table, ref sErr);
            kq._connID = schInf.DEFAULT_CONN;
            kq.Database = database;
            CoreQDDControl qddControl = new CoreQDDControl();
            //test
            //  DataTable dt = qddControl.GetAll_CoreQDD_By_QD_ID("DEMO4", "VUS", ref sErr);
            using (DataTable dt = qddControl.GetALL_CoreQDD_By_QD_ID(kq.Database, qd_id, ref sErr))
            {
                foreach (DataRow row in dt.Rows)
                {
                    CoreQDDInfo qddInfo = new CoreQDDInfo(row);
                    Node tmp = new Node(qddInfo.AGREGATE.Trim(), qddInfo.CODE.Trim(), qddInfo.DESCRIPTN.Trim(), qddInfo.F_TYPE, "");
                    if (qddInfo.IS_FILTER == true)
                    {
                        CoreQDD_FILTERControl filterCtr = new CoreQDD_FILTERControl();
                        CoreQDD_FILTERInfo filterInf = filterCtr.Get(database, qd_id, qddInfo.QDD_ID, ref sErr);
                        tmp.NodeDesc = qddInfo.EXPRESSION;
                        Filter tmpFilter = new Filter(tmp);
                        //tmpFilter.Description = qddInfo.DESCRIPTN;
                        //tmpFilter.Code = qddInfo.CODE;
                        if (filterInf.QD_ID != "")
                        {
                            tmpFilter.Operate = filterInf.OPERATOR;
                            tmpFilter.IsNot = filterInf.IS_NOT;
                        }
                        tmpFilter.ValueFrom = tmpFilter.FilterFrom = qddInfo.FILTER_FROM;
                        tmpFilter.ValueTo = tmpFilter.FilterTo = qddInfo.FILTER_TO;
                        kq.Filters.Add(tmpFilter);
                    }
                    else
                    {
                        tmp.Expresstion = qddInfo.EXPRESSION;
                        tmp.Sort = qddInfo.SORTING == "DES" ? "DESC" : qddInfo.SORTING;
                        kq.SelectedNodes.Add(tmp);
                    }
                }
            }
            if (table != "")
                kq.Table = table.Trim();
            else
            {
                CoreQDControl qdCtr = new CoreQDControl();
                CoreQDInfo qdInf = qdCtr.Get_CoreQD(database, qd_id, ref sErr);
                kq.Table = qdInf.ANAL_Q0;
            }
            return kq;

        }

        public static SQLBuilder LoadSQLBuilderFromDataBase(string qd_id, string database, string table)
        {
            string sErr = "";
            SQLBuilder kq = new SQLBuilder(processingMode.Details);
            CoreQD_SCHEMAControl schCtr = new CoreQD_SCHEMAControl();
            if (table != "")
                kq.Table = table.Trim();
            else
            {
                CoreQDControl qdCtr = new CoreQDControl();
                CoreQDInfo qdInf = qdCtr.Get_CoreQD(database, qd_id, ref sErr);
                table = kq.Table = qdInf.ANAL_Q0;
            }
            CoreQD_SCHEMAInfo schInf = schCtr.Get(database, table, ref sErr);
            kq._connID = schInf.DEFAULT_CONN;
            kq.Database = database;
            CoreQDDControl qddControl = new CoreQDDControl();
            //test
            //  DataTable dt = qddControl.GetAll_CoreQDD_By_QD_ID("DEMO4", "VUS", ref sErr);
            using (DataTable dt = qddControl.GetALL_CoreQDD_By_QD_ID(kq.Database, qd_id, ref sErr))
            {
                foreach (DataRow row in dt.Rows)
                {
                    CoreQDDInfo qddInfo = new CoreQDDInfo(row);
                    Node tmp = new Node(qddInfo.AGREGATE.Trim(), qddInfo.CODE.Trim(), qddInfo.DESCRIPTN.Trim(), qddInfo.F_TYPE, "");
                    if (qddInfo.IS_FILTER == true)
                    {
                        CoreQDD_FILTERControl filterCtr = new CoreQDD_FILTERControl();
                        CoreQDD_FILTERInfo filterInf = filterCtr.Get(database, qd_id, qddInfo.QDD_ID, ref sErr);
                        tmp.NodeDesc = qddInfo.EXPRESSION;
                        Filter tmpFilter = new Filter(tmp);
                        //tmpFilter.Description = qddInfo.DESCRIPTN;
                        //tmpFilter.Code = qddInfo.CODE;
                        if (filterInf.QD_ID != "")
                        {
                            tmpFilter.Operate = filterInf.OPERATOR;
                            tmpFilter.IsNot = filterInf.IS_NOT;
                        }
                        tmpFilter.ValueFrom = tmpFilter.FilterFrom = qddInfo.FILTER_FROM;
                        tmpFilter.ValueTo = tmpFilter.FilterTo = qddInfo.FILTER_TO;
                        kq.Filters.Add(tmpFilter);
                    }
                    else
                    {
                        tmp.Expresstion = qddInfo.EXPRESSION;
                        tmp.Sort = qddInfo.SORTING == "DES" ? "DESC" : qddInfo.SORTING;
                        kq.SelectedNodes.Add(tmp);
                    }
                }
            }
            if (table != "")
                kq.Table = table.Trim();
            else
            {
                CoreQDControl qdCtr = new CoreQDControl();
                CoreQDInfo qdInf = qdCtr.Get_CoreQD(database, qd_id, ref sErr);
                kq.Table = qdInf.ANAL_Q0;
            }
            return kq;

        }

        public Filter SelectFilter(int index)
        {
            if (index < _filters.Count)
                return _filters[index];

            return null;
        }
        public void UpdateFilter(int index, string from, string to)
        {
            if (index < _nodes.Count)
            {
                _filters[index].FilterFrom = from;
                _filters[index].FilterTo = to;
            }

            //       return null;
        }
        public void DeleteFilter(int index)
        {
            if (index < _nodes.Count)
            {
                _filters.RemoveAt(index);
            }

            //        return null;
        }
        public string BuildSQLEx(string sql_text)
        {
            string query = "";
            if (sql_text == "")
            {
                if (SelectedNodes.Count > 0)
                    query = BuildSQL();
                else
                    query = "";
            }
            else
            {
                query = sql_text.Replace("XXX", Database);
                foreach (Filter x in Filters)
                {
                    if (Regex.IsMatch(x.Code, @"^@"))
                        query = query.Replace(x.Code, x.FilterFrom);
                }
            }
            if (_SQLDebugMode)
            {
                //CoreCommonControl log = new CoreCommonControl();
                string __documentDirectory = String.Format("{0}\\{1}", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), DocumentFolder);
                CoreCommonControl.AddLog("cmdlog", __documentDirectory + "\\Log", String.Format("{0:yyyyMMdd}{1}:{2}", DateTime.Today, Table, query));

                //System.Windows.Forms.Clipboard.SetText(query);
            }
            return query;
        }
        /*
        public DataTable BuildDataTable(string sql_text)
        {
            DataTable dt = new DataTable();
            string query = BuildSQLEx(sql_text);
            CommoControl control = new CommoControl();
            dt = control.executeSelectQuery(query);
            foreach (Filter x in Filters)
            {
                if (x.Node.Code.Substring(0, 2) == "__")
                {
                    DataColumn col = new DataColumn(x.Node.Code);

                    if (Regex.IsMatch(x.ValueFrom, @"^\d$"))
                        col.Expression = x.ValueFrom;
                    else
                        col.Expression = "'" + x.ValueFrom + "'";
                    dt.Columns.Add(col);
                }
            }
            foreach (Node x in _nodes)
            {
                if (x.Expresstion != "")
                {
                    DataColumn col = new DataColumn(x.Code);
                    if (x.FType == "N")
                        col.DataType = typeof(Decimal);
                    else if (x.FType == "SDN")
                        col.DataType = typeof(Int32);
                    else
                        col.DataType = typeof(String);
                    col.Expression = x.Expresstion;
                    dt.Columns.Add(col);
                    //col.Expression = "";
                    //dt.DefaultView.DataViewManager.
                }
                else if (x.FType == "SDN")
                {
                    DataColumn col = new DataColumn(x.Code + "_");
                    col.Expression = CommoControl.GetParseExpressionDate(x.MyCode, "A");
                    dt.Columns.Add(col);
                    //col.Expression = "";
                }
                else if (x.FType == "SPN")
                {
                    DataColumn col = new DataColumn(x.Code + "_");
                    col.Expression = CommoControl.GetParseExpressionPeriod(x.MyCode);
                    dt.Columns.Add(col);
                    //col.Expression = "";
                }
                //dt.Columns[x.MyCode].Caption = x.Description;
            }
            DataTable result = new DataTable("data");
            result.TableName = dt.TableName;
            foreach (DataColumn col in dt.Columns)
            {
                DataColumn tmp = new DataColumn(col.ColumnName, col.DataType);                
                result.Columns.Add(tmp);
            }
            foreach (DataRow row in dt.Rows)
            {
                result.ImportRow(row);
            }
            DataSet dtSet = new DataSet();
            dtSet.Tables.Add(result);
            return result;
        }*/
        public object BuildObject(string sql_text, string connectString)
        {
            DataTable dt = new DataTable();
            string query = BuildSQLEx(sql_text);
            OleDbConnection connection = new OleDbConnection(connectString);
            OleDbCommand adapter = new OleDbCommand(query, connection);
            int timeout = 0;
            object result = null;
            try
            {
                connection.Open();
                string[] arr = _strConnectDes.Split(';');
                for (int i = 0; i < arr.Length; i++)
                {
                    string[] arrP = arr[i].Split('=');
                    if (arrP.Length == 2)
                    {
                        if (arrP[0] == "General Timeout")
                        {
                            timeout = Convert.ToInt32(arrP[1]);
                        }
                    }
                }
                adapter.CommandTimeout = timeout;
                result = adapter.ExecuteScalar();
            }
            catch (Exception ex)
            {
                result = "";
            }
            finally
            {
                connection.Close();
            }
            return result;
        }

        public DataTable BuildDataTable(string sql_text)
        {
            DataTable dt = new DataTable();
            string query = BuildSQLEx(sql_text);
            //[oledb]
            //; Everything after this line is an OLE DB initstring
            //Provider=SQLNCLI.1;Persist Security Info=False;User ID=sa;Initial Catalog=TVC_IC;Data Source=.
            int timeout = 0;
            using (OleDbConnection connection = new OleDbConnection(_strConnectDes))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);
                DataSet dSet = new DataSet();
                if (query != "")
                    try
                    {
                        string[] arr = _strConnectDes.Split(';');
                        for (int i = 0; i < arr.Length; i++)
                        {
                            string[] arrP = arr[i].Split('=');
                            if (arrP.Length == 2)
                            {
                                if (arrP[0] == "General Timeout")
                                {
                                    timeout = Convert.ToInt32(arrP[1]);
                                }
                            }
                        }
                        adapter.SelectCommand.CommandTimeout = timeout;
                        adapter.Fill(dSet);
                        dt = dSet.Tables[0];
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
            }
            //CommoControl control = new CommoControl();
            //dt = control.executeSelectQuery(query, connectString);
            foreach (Filter x in Filters)
            {
                if (x.Node.Code.Substring(0, 2) == "__")
                {
                    DataColumn col = new DataColumn(x.Node.Code);

                    if (Regex.IsMatch(x.ValueFrom, @"^\d$"))
                        col.Expression = x.ValueFrom;
                    else
                        col.Expression = String.Format("'{0}'", x.ValueFrom);
                    dt.Columns.Add(col);
                }
            }
            if (sql_text == "")
                foreach (Node x in _nodes)
                {
                    if (x.Expresstion != "")
                    {
                        DataColumn col = new DataColumn(x.Code);
                        if (x.FType == "N")
                            col.DataType = typeof(Decimal);
                        else if (x.FType == "SDN")
                            col.DataType = typeof(Int32);
                        else
                            col.DataType = typeof(String);
                        col.Expression = x.Expresstion;
                        if (!dt.Columns.Contains(col.ColumnName))
                            dt.Columns.Add(col);
                        //col.Expression = "";
                        //dt.DefaultView.DataViewManager.
                    }
                    else if (x.FType == "SDN")
                    {
                        DataColumn col = new DataColumn(x.Code + "_");
                        col.Expression = CoreCommonControl.GetParseExpressionDate(x.Description, "A");
                        if (!dt.Columns.Contains(col.ColumnName))
                            dt.Columns.Add(col);
                        //col.Expression = "";
                    }
                    else if (x.FType == "SPN")
                    {
                        DataColumn col = new DataColumn(x.Code + "_");
                        col.Expression = CoreCommonControl.GetParseExpressionPeriod(x.Description);
                        if (!dt.Columns.Contains(col.ColumnName))
                            dt.Columns.Add(col);
                        //col.Expression = "";
                    }
                    //dt.Columns[x.MyCode].Caption = x.Description;
                }
            DataTable result = new DataTable("data") { TableName = dt.TableName };
            foreach (DataColumn col in dt.Columns)
            {
                DataColumn tmp = new DataColumn(col.ColumnName, col.DataType);
                if (!result.Columns.Contains(tmp.ColumnName))
                    result.Columns.Add(tmp);
            }
            foreach (DataRow row in dt.Rows)
            {
                foreach (DataColumn col in dt.Columns)
                {
                    if (col.Expression == "" && row[col.ColumnName] == DBNull.Value)
                    {
                        if (col.DataType == typeof(string))
                            row[col.ColumnName] = "";
                        else if (col.DataType == typeof(DateTime))
                        { } //row[col.ColumnName] = new DateTime(1900, 1, 1);
                        else
                            row[col.ColumnName] = 0;
                    }
                }
                result.ImportRow(row);
            }
            using (DataSet dtSet = new DataSet())
            {
                dtSet.Tables.Add(result);
            }
            return result;
        }
        #endregion


        public static string SetFunctions(string formular)
        {
            foreach (System.Text.RegularExpressions.Match m in Regex.Matches(formular, @"<#.+(.+)>"))
            {
                string value = "";
                if (Regex.IsMatch(m.Value, @"<#PH(.+)>"))
                {
                    value = m.Value.Replace("<#PH(", "").Replace(")>", "");
                    if (Regex.IsMatch(value, periodParterm))
                    {
                        int year = Convert.ToInt32(value) / 1000;
                        int month = Convert.ToInt32(value) - year * 1000;
                        month--;
                        if (month == 0)
                        {
                            year--;
                            month = 12;
                        }
                        value = year + month.ToString("000");
                    }
                }
                else if (Regex.IsMatch(m.Value, @"<#PA(.+)>"))
                {
                    value = m.Value.Replace("<#PA(", "").Replace(")>", "");
                    if (Regex.IsMatch(value, periodParterm))
                    {
                        int year = Convert.ToInt32(value) / 1000;
                        int month = Convert.ToInt32(value) - year * 1000;
                        //month--;
                        //if (month == 0)
                        //{
                        //    year--;
                        //    month = 12;
                        //}
                        value = year + month.ToString("000");
                    }
                }
                else if (Regex.IsMatch(m.Value, @"<#PE(.+)>"))
                {
                    value = m.Value.Replace("<#PE(", "").Replace(")>", "");
                    if (Regex.IsMatch(value, periodParterm))
                    {
                        int year = Convert.ToInt32(value) / 1000;
                        int month = Convert.ToInt32(value) - year * 1000;
                        month = 12;
                        value = year + month.ToString("000");
                    }
                }
                else if (Regex.IsMatch(m.Value, @"<#YA(.+)>"))
                {
                    value = m.Value.Replace("<#YA(", "").Replace(")>", "");
                    if (Regex.IsMatch(value, periodParterm))
                    {
                        int year = Convert.ToInt32(value) / 1000;
                        int month = Convert.ToInt32(value) - year * 1000;
                        month = 1;
                        value = year + month.ToString("000");
                    }
                }
                else if (Regex.IsMatch(m.Value, @"<#YE(.+)>"))
                {
                    value = m.Value.Replace("<#YE(", "").Replace(")>", "");
                    if (Regex.IsMatch(value, periodParterm))
                    {
                        int year = Convert.ToInt32(value) / 1000;
                        int month = Convert.ToInt32(value) - year * 1000;
                        year--;
                        month = 1;
                        value = year + month.ToString("000");
                    }
                }
                else if (Regex.IsMatch(m.Value, @"<#YH(.+)>"))
                {
                    value = m.Value.Replace("<#YH(", "").Replace(")>", "");
                    if (Regex.IsMatch(value, periodParterm))
                    {
                        int year = Convert.ToInt32(value) / 1000;
                        int month = Convert.ToInt32(value) - year * 1000;
                        month = 12;
                        year--;
                        value = year + month.ToString("000");
                    }
                }
                else if (Regex.IsMatch(m.Value, @"<#YK(.+)>"))
                {
                    value = m.Value.Replace("<#YK(", "").Replace(")>", "");
                    if (Regex.IsMatch(value, periodParterm))
                    {
                        int year = Convert.ToInt32(value) / 1000;
                        int month = Convert.ToInt32(value) - year * 1000;
                        month = 12;
                        value = year + month.ToString("000");
                    }
                }
                else
                {
                    value = m.Value.Replace("<#.+(", "").Replace(")>", "");
                }
                formular = formular.Replace(m.Value, value);
            }
            return formular;
        }

        public DataTable BuildDataTableStruct(string sql_text, string connectString)
        {
            DataTable dt = new DataTable();
            string query = BuildSQLEx(sql_text);
            int index = query.IndexOf(STR_SELECT);
            query.Insert(index + 6, " TOP 0");
            using (OleDbConnection connection = new OleDbConnection(connectString))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);
                DataSet dSet = new DataSet();
                try
                {
                    adapter.Fill(dSet);
                    dt = dSet.Tables[0];
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            //CommoControl control = new CommoControl();
            //dt = control.executeSelectQuery(query, connectString);
            foreach (Filter x in Filters)
            {
                if (x.Node.Code.Substring(0, 2) == "__")
                {
                    DataColumn col = new DataColumn(x.Node.Code);

                    if (Regex.IsMatch(x.ValueFrom, @"^\d$"))
                        col.Expression = x.ValueFrom;
                    else
                        col.Expression = String.Format("'{0}'", x.ValueFrom);
                    dt.Columns.Add(col);
                }
            }
            if (sql_text == "")
                foreach (Node x in _nodes)
                {
                    if (x.Expresstion != "")
                    {
                        DataColumn col = new DataColumn(x.Code);
                        if (x.FType == "N")
                            col.DataType = typeof(Decimal);
                        else if (x.FType == "SDN")
                            col.DataType = typeof(Int32);
                        else
                            col.DataType = typeof(String);
                        col.Expression = x.Expresstion;
                        dt.Columns.Add(col);
                        //col.Expression = "";
                        //dt.DefaultView.DataViewManager.
                    }
                    else if (x.FType == "SDN")
                    {
                        DataColumn col = new DataColumn(x.Code + "_");
                        col.Expression = CoreCommonControl.GetParseExpressionDate(x.Description, "A");
                        dt.Columns.Add(col);
                        //col.Expression = "";
                    }
                    else if (x.FType == "SPN")
                    {
                        DataColumn col = new DataColumn(x.Code + "_");
                        col.Expression = CoreCommonControl.GetParseExpressionPeriod(x.Description);
                        dt.Columns.Add(col);
                        //col.Expression = "";
                    }
                    //dt.Columns[x.MyCode].Caption = x.Description;
                }

            return dt;
        }
    }
}
