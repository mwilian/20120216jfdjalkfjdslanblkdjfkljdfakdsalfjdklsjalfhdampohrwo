using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace QueryBuilder
{
    //  Class Node
    public class Node
    {
        private const string Tab = "  ";
        private const string Space = " ";

        private const string Comma = ",";
        private const string STR_SELECT = "SELECT ";
        private const string STR_FROM = " FROM ";
        private const string STR_GROUPBY = " GROUP BY ";
        private const string STR_WHERE = " WHERE ";
        private const string STR_DOT = ".";
        private const string STR_LEFTOUTERJOIN = "LEFT OUTER JOIN ";
        private const string STR_SUBSTRING = "SUBSTRING";
        private const string LeftSign = "(";
        private const string RightSign = ")";

        private const string Regex_Root = @"^[^\/,\\]*(?=[\/,\\]|^)";
        private const string Regex_Leaf = @"(?<=[\/,\\]|^)[^\/,\\]*$";
        private const string Regex_Family = @"[A-Z,a-z,0-9,_,\/,\\]+(?=[\/,\\])";
        private const string Regex_Child = @"(?<=[\/,\\])[A-Z,a-z,0-9,_,\/,\\]+";

        #region '" Members "'
        private string _Code = string.Empty;
        private string _NodeDesc = string.Empty;
        private string _description = string.Empty;
        private string _ftype = string.Empty;
        private string _sort = string.Empty;
        private string _agregate = string.Empty;
        private string _TreeCode = string.Empty;
        private string _expresstion = string.Empty;
        private int _idParentTree;

        private int _idTree;

        public string NodeDesc
        {
            get { return _NodeDesc; }
            set { _NodeDesc = value; }
        }
        public string Expresstion
        {
            get { return _expresstion; }
            set { _expresstion = value; }
        }
        public int IdTree
        {
            get { return _idTree; }
            set { _idTree = value; }
        }


        public int IdParentTree
        {
            get { return _idParentTree; }
            set { _idParentTree = value; }
        }

        //  Property TreeCode
        public string TreeCode
        {
            get
            {
                return _TreeCode;
            }
        }
        //  Property Code
        public string Code
        {
            get
            {
                return _Code;
            }
        }
        //  Property FType
        public string FType
        {
            get
            {
                if (string.IsNullOrEmpty(_ftype))
                {
                    return string.Empty;
                }
                else
                {
                    return (_ftype + new string(' ', 5)).Substring(0, 5).Trim();
                }
            }
        }
        public string FTypeFull
        {
            get
            {
                if (string.IsNullOrEmpty(_ftype))
                {
                    return "          ";
                }
                else
                {
                    return (_ftype + new string(' ', 100)).Substring(0, 100);
                }
            }
            set { _ftype = value; }
        }
        public int Index
        {
            get
            {
                int from = 0;
                if (Regex.IsMatch(FTypeFull.Substring(5, 2), @"^\d+$"))
                    from = Convert.ToInt32(FTypeFull.Substring(5, 2));
                return from;
            }
        }
        public int Length
        {
            get
            {
                int len = 0;
                if (Regex.IsMatch(FTypeFull.Substring(7, 2), @"^\d+$"))
                    len = Convert.ToInt32(FTypeFull.Substring(7, 2));
                return len;
            }
        }
        //  Property Description
        public string Description
        {
            get
            {
                return _description;
            }
            set
            {
                _description = value;
            }
        }
        //  Property Sort
        public string Sort
        {
            get
            {
                return _sort;
            }
            set
            {
                if (value == null)
                    _sort = "";
                else
                    _sort = value;
            }
        }
        //  Property Agregate
        public string Agregate
        {
            get
            {
                if (_agregate == null)
                    return "";
                return _agregate;
            }
            set
            {
                if (value == null)
                    _agregate = "";
                else
                    _agregate = value;
            }
        }

        ///  <summary>
        ///  Return the ID code of the node for exp LA\TREF return TREF
        ///  </summary>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public string MyCode
        {
            get
            {
                return GetLeaf(_Code);
            }
        }
        //  Property MyParent
        public string MyParent
        {
            get
            {
                return GetParent(_Code);
            }
        }
        //  Property MyAlias
        public string MyAlias
        {
            get
            {
                if (Regex.IsMatch(MyCode.Trim(), "^[0-9]+$"))
                {
                    return "_" + MyCode;
                }
                else
                {
                    return " as " + MyCode;
                }

            }
        }
        ///  <summary>
        ///  Return the family code of the node for exp LA\CA\TREF return LA\CA
        ///  </summary>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public string MyFamily
        {
            get
            {
                return GetFamily(_Code);
            }
        }

        #endregion

        #region '" Node Processing "'

        #region '" Business Specific "'

        //  Method AgregateMe
        public string AgregateMe()
        {
            if (string.IsNullOrEmpty(_agregate))
            {
                return FormatMe();
            }
            else
            {
                return String.Format("{0}({1})", _agregate, FormatMe());
            }
        }

        //  Method FormatMe
        public string FormatMe()
        {
            if (Length == 0)
                return FormatNode(_Code);
            else
                return STR_SUBSTRING + LeftSign + FormatNode(_Code) + Comma + Index + Comma + Length + RightSign;
        }

        //  Method FormatMyParameter
        public string FormatMyParameter(string param)
        {
            if (isEmpty())
            {
                return string.Empty;
            }
            switch (FType)
            {
                case null:
                case "":
                    return string.Format("N'{0}'", param.Replace("'", String.Empty));
                case "D":
                    if (param == "C")
                        param = DateTime.Now.ToString("yyyy-MM-dd");
                    return string.Format("N'{0}'", param.Replace("'", String.Empty));

                case "SDN":
                    if (param == "C")
                        param = DateTime.Now.ToString("yyyyMMdd");
                    return string.Format("{0}", param.Replace("'", String.Empty));

                case "SP":
                case "SPN":
                    if (Regex.IsMatch(param, "[0-9]{8}"))
                    {
                        string sp = string.Format("'{0}0{1}'", param.Substring(0, 4), param.Substring(4, 2));
                        return sp;
                    }
                    else if (param == "C")
                    {
                        string sp = string.Format("'{0}0{1}'", DateTime.Now.Year, DateTime.Now.Month.ToString("000"));
                        return sp;
                    }
                    

                    break;
                default:
                    return string.Format("{0}", param.Replace("'", String.Empty));//"N'{0}'"
                    break;


            }

            return param;
        }


        ///  <summary>
        ///  Return the Root code of the node for exp LA\CA\TREF return LA
        ///  </summary>
        ///  <remarks></remarks>
        public void AddMeToParent(string ParentCode)
        {
            _Code = String.Format(@"{0}\{1}", ParentCode, _Code);
            _TreeCode = String.Format(@"{0}\{1}", ParentCode, _TreeCode);
        }


        #region '" Build TreeView Control "'
        ///  <summary>
        ///  The point of view is the root. LA\CA\A0 - me is LA. my child is CA\A0.
        ///  Use this function only for generate treeView control
        ///  </summary>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public Node MyChild()
        {
            Node n = new Node() { _TreeCode = GetChild(_TreeCode), _Code = _Code, Agregate = Agregate, Description = _description, _ftype = _ftype, _sort = _sort, NodeDesc = NodeDesc };
            return n;
        }

        //  Method HasChild
        public bool HasChild()
        {
            return !(string.IsNullOrEmpty(GetChild(_TreeCode.Trim())));
        }

        #endregion

        ///  <summary>
        ///  LA\CA\NA - > Get full Joins LA-CA and CA-NA  with avoiding duplication using REgistered Alias list and stop at Family
        ///  </summary>
        ///  <param name="StopAtFamily"></param>
        ///  <param name="RegisteredAliasList">Avoid duplication of Alias</param>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public string MyJoinClause(string dtb, string StopAtFamily, ref List<string> RegisteredAliasList)
        {

            string temp = string.Empty;
            string family = MyFamily;

            if (RegisteredAliasList == null)
            {
                RegisteredAliasList = new List<string>();
            }

            while (String.Compare(StopAtFamily, family) < 0)
            {
                // Dim origin As String = SchemaDefinition.GetOriginFromAlias(GetParent(family))
                // If the family is alias - ex . ICAS\NA- then get CA instead of ICAS (define in _ALIAS XML file)
                if (!(RegisteredAliasList.Contains(GetLeaf(family))))
                { //  if the alias ex. CA is not registered yet 

                    temp = Tab + GetJoinClause(dtb, GetParent(family), GetLeaf(family)) + Space + System.Environment.NewLine + temp;
                    RegisteredAliasList.Add(GetLeaf(family));
                }

                family = GetFamily(family);
            }
            return temp;

        }


        //  Method Equals
        public override bool Equals(object obj)
        {
            if (!(obj is Node))
            {
                return false;
            }
            return _Code == ((Node)(obj)).Code;
        }


        //  Method ToString
        public override string ToString()
        {
            return _Code;
        }

        //  Method isEmpty
        public bool isEmpty()
        {
            if (string.IsNullOrEmpty(_Code))
            {
                return true;
            }
            return false;
        }

        #endregion

        #region '" Service shared Function "'

        ///  <summary>
        ///  give LA\CA\ACCNT_CODE return [CA].[ACCNT_CODE]
        ///  </summary>
        ///  <param name="nodeCode">LA\TRefe</param>
        ///  <returns>[LA].[Tref]</returns>
        ///  <remarks></remarks>
        public static string FormatNode(string nodeCode)
        {
            if (string.IsNullOrEmpty(nodeCode))
            {
                return string.Empty;
            }
            string leaf = GetLeaf(nodeCode);
            string parent = GetParent(nodeCode);
            if (Regex.IsMatch(leaf, @".+\(.+\)"))
            {
                string left = leaf.Substring(0, leaf.IndexOf('(') + 1);
                return left + BoxBracketWithDot(parent) + leaf.Substring(leaf.IndexOf('(') + 1);
            }
            return BoxBracketWithDot(parent) + BoxBracket(leaf);
        }

        ///  <summary>
        ///  Return the family code of the node for exp LA\CA\TREF return LA\CA
        ///  </summary>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public static string GetFamily(string nodeCode)
        {
            if (string.IsNullOrEmpty(nodeCode))
            {
                return string.Empty;
            }

            string _family = Regex.Match(nodeCode.Trim(), Regex_Family).ToString().Trim();
            return _family;
        }

        //  Method GetParent
        public static string GetParent(string nodeCode)
        {
            if (string.IsNullOrEmpty(nodeCode))
            {
                return string.Empty;
            }

            string _parent = Regex.Match(nodeCode.Trim(), Regex_Family).ToString().Trim();
            string _alias = Regex.Match(_parent, Regex_Leaf).ToString().Trim();
            return _alias;
        }

        //  Method GetLeaf
        public static string GetLeaf(string nodeCode)
        {
            if (string.IsNullOrEmpty(nodeCode))
            {
                return string.Empty;
            }

            return Regex.Match(nodeCode.Trim(), Regex_Leaf).ToString();
        }

        ///  <summary>
        ///  Return the Root code of the node for exp LA\CA\TREF return LA
        ///  </summary>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public static string GetRoot(string nodeCode)
        {
            if (string.IsNullOrEmpty(nodeCode))
            {
                return string.Empty;
            }

            return Regex.Match(nodeCode.Trim(), Regex_Root).ToString();
        }

        ///  <summary>
        ///  Return the Child code of the node for exp LA\CA\TREF return CA\TREF
        ///  </summary>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public static string GetChild(string nodeCode)
        {
            if (string.IsNullOrEmpty(nodeCode))
            {
                return string.Empty;
            }

            return Regex.Match(nodeCode.Trim(), Regex_Child).ToString();
        }


        ///  <summary>
        ///  LA\CA\NA - > Get full Joins LA-CA and CA-NA
        ///  </summary>
        ///  <param name="Parent"></param>
        ///  <param name="Child"></param>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public static string GetJoinClause(string dtb, string Parent, string Child)
        { // EX IR\ICAS . ex 2 ICAS\NA

            if (string.IsNullOrEmpty(Child.Trim()))
            {
                return string.Empty;
            }
            if (string.IsNullOrEmpty(Parent.Trim()))
            { //  no parent . mean get ICAS

                return SchemaDefinition.GetJoin(dtb, Child) + Space + Child;
            }
            else
            {
                // if child is alias . then lookup original IR\CA instead of IR\ICAS  , CA\NA instead of ICAS\NA
                string originOfChild = SchemaDefinition.GetOriginFromAlias(Child); //  = CA  ,   NA
                string originOfParent = SchemaDefinition.GetOriginFromAlias(Parent); //  =IR ,   CA

                // LEFT OUTER JOIN (SELECT CA) ICAS ON     , LEFT OUTER JOIN (SELECT NA) NA ON
                return String.Format("{0}{1}{2}{3} ON {4}", STR_LEFTOUTERJOIN, SchemaDefinition.GetJoin(dtb, originOfChild), Space, Child, Regex.Replace(SchemaDefinition.GetJoin(dtb, String.Format(@"{0}\{1}", originOfParent, Child)), "^" + originOfParent, Parent));
                // replace : on "IR.SALES_ACC = ICAS.CODE" to "IR.SALES_ACC = ICAS.CODE"/> 
                // replace : on "ICAS.ADDR_CODE = NA.ADDR_CODE" to "CA.ADDR_CODE = NA.ADDR_CODE"/>  
            }
        }


        #region '" Formating "'
        //  Method BoxBracket
        private static string BoxBracket(string original)
        {
            if (string.IsNullOrEmpty(original))
            {
                return string.Empty;
            }
            return String.Format("[{0}]", original.Trim());
        }

        //  Method BoxBracketWithDot
        private static string BoxBracketWithDot(string original)
        {
            if (string.IsNullOrEmpty(original))
            {
                return string.Empty;
            }
            return String.Format("[{0}].", original.Trim());
        }

        #endregion
        #endregion

        #endregion

        #region '" Constructors "'
        public Node(string agregate, string Code, string name, string FType, string nodeDesc)
        {
            _Code = Code;
            _TreeCode = Code;
            _description = name;
            _agregate = agregate;
            _ftype = FType;
            _NodeDesc = nodeDesc;
        }
        public Node(string Code, string name)
        {
            _Code = Code;
            _TreeCode = Code;
            _description = name;

            // improve speed by maintaining a list of all period node
            if (Code.Contains("Period".ToUpper()))
            {
                _ftype = "SPN";
            }

        }
        private Node()
        {
        }

        //  Method EmptyNode
        public static Node EmptyNode()
        {
            Node _node = new Node();
            return _node;
        }


        //  Method CloneNode
        public Node CloneNode()
        {
            Node n = new Node() { _Code = _Code, _description = _description, _ftype = _ftype, _sort = _sort, _agregate = _agregate, _TreeCode = this._TreeCode, NodeDesc = this.NodeDesc };
            return n;
        }


        //  Method DecorateDescriptn
        public void DecorateDescriptn(string desc)
        {
            _description = desc;
        }

        #endregion





        internal string FormatMyArrayParameter(string ValueFrom)
        {
            if (isEmpty())
            {
                return string.Empty;
            }
            string[] arrValueFrom = ValueFrom.Split(',');
            for (int i = 0; i < arrValueFrom.Length; i++)
            {
                string param = arrValueFrom[i];
                switch (FType)
                {
                    case null:
                    case "":
                        param = string.Format("N'{0}'", param.Replace("'", String.Empty));
                        break;
                    case "D":
                        if (param == "C")
                            param = DateTime.Now.ToString("yyyy-MM-dd");
                        param = string.Format("N'{0}'", param.Replace("'", String.Empty));
                        break;
                    case "SDN":
                        if (param == "C")
                            param = DateTime.Now.ToString("yyyyMMdd");
                        param = string.Format("{0}", param.Replace("'", String.Empty));
                        break;
                    case "SP":
                    case "SPN":
                        if (Regex.IsMatch(param, "[0-9]{8}"))
                        {
                            param = string.Format("'{0}0{1}'", param.Substring(0, 4), param.Substring(4, 2));
                        }
                        else if (param == "C")
                        {
                            param = string.Format("'{0}0{1}'", DateTime.Now.Year, DateTime.Now.Month.ToString("000"));
                        }

                        break;
                    default:
                        param = string.Format("{0}", param.Replace("'", String.Empty));//"N'{0}'"
                        break;


                }
                arrValueFrom[i] = param;
            }

            ValueFrom = string.Join(",", arrValueFrom);
            return String.Format("({0})", ValueFrom);
        }
    }
}
