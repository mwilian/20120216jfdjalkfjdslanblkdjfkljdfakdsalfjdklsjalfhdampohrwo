using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace QueryBuilder
{
    //  Class Filter
    public class Filter : IComparable
    {
        private const string Wildcard_Signature = @"^[\*][=]";

        private Node _node;
        private string _filterFrom = string.Empty;
        private string _filterTo = string.Empty;
        private string _filterFromP = string.Empty;
        private string _filterToP = string.Empty;
        private string _ValueFrom = string.Empty;
        private string _ValueTo = string.Empty;
        private const string _insert = "INSERT";
        private const string _update = "UPDATE";
        private const string _delete = "DELETE";
        private string _isNot = "N";
        private string _operate = "-";

        public string Operate
        {
            get { return _operate; }
            set { _operate = value; }
        }
        public string IsNot
        {
            get { return _isNot; }
            set { _isNot = value; }
        }
        //  Property Node
        public Node Node
        {
            get
            {
                return _node;
            }
        }
        public string Code
        {
            get
            {
                return _node.MyCode;
            }
        }
        //  Property Description
        public string Description
        {
            get
            {
                return _node.Description;
            }
        }
        //  Property FilterFrom
        public string FilterFrom
        {
            get
            {
                return _filterFrom;
            }
            set
            {

                if (value == null)
                    _filterFrom = "";
                else
                    _filterFrom = value.Trim();
                if (!(_eval == null))
                {

                    _eval.Invoke(_filterFrom, ref _filterFromP, ref _ValueFrom);
                }
            }
        }
        //  Property FilterTo
        public string FilterTo
        {
            get
            {
                return _filterTo;
            }
            set
            {
                if (value == null)
                    _filterTo = "";
                else
                    _filterTo = value.Trim();
                if (!(_eval == null))
                {

                    _eval.Invoke(_filterTo, ref _filterToP, ref _ValueTo);
                }
            }
        }
        //  Property FilterFromP
        public string FilterFromP
        {
            get
            {
                return _filterFromP;
            }
            set
            {
                _filterFromP = value;
                _filterFromP = _filterFromP.Trim();
            }
        }
        //  Property FilterToP
        public string FilterToP
        {
            get
            {
                return _filterToP;
            }
            set
            {
                _filterToP = value;
                _filterToP = _filterToP.Trim();
            }
        }
        //  Property ValueFrom
        public string ValueFrom
        {
            get
            {
                return _ValueFrom;
            }
            set
            {
                string tmp = value.ToUpper();

                if (tmp.Contains(_insert) || tmp.Contains(_update) || tmp.Contains(_delete))
                {
                    _ValueFrom = tmp.Replace(_insert, "").Replace(_update, "").Replace(_delete, "");
                }
                else
                {
                    _ValueFrom = value;
                    _ValueFrom = _ValueFrom.Trim();
                }
            }
        }
        //  Property ValueTo
        public string ValueTo
        {
            get
            {
                return _ValueTo;
            }
            set
            {
                string tmp = value.ToUpper();

                if (tmp.Contains(_insert) || tmp.Contains(_update) || tmp.Contains(_delete))
                {
                    _ValueTo = tmp.Replace(_insert, "").Replace(_update, "").Replace(_delete, "");
                }
                else
                {
                    _ValueTo = value;
                    _ValueTo = _ValueTo.Trim();
                }
            }
        }


        public Filter(Node aNode)
        {
            _node = aNode;
        }
        public Filter(Node aNode, string filterFrom, string filterTo, string valueFrom, string valueTo, string filterfromP, string filtertoP)
        {
            _node = aNode;
            _filterFrom = filterFrom.Trim();
            _filterTo = filterTo.Trim();
            _ValueFrom = valueFrom.Trim();
            _ValueTo = valueTo.Trim();
            _filterFromP = filterfromP.Trim();
            _filterToP = filtertoP.Trim();
        }

        //  Method usingWildCard
        private bool usingWildCard()
        {
            // string test = @"^[\*][=]";
            if (Regex.IsMatch(_ValueFrom.Trim(), @"(>|<|>=|<=|<<|\[|%|_|\]|\*>|\!)"))//Wildcard_Signature
            {
                return true;
            }


            //if (Regex.IsMatch(_ValueFrom.Trim(), @"%|_"))
            //{
            //    return true;
            //}
            return false;
        }


        //  Method MyWhereClause
        public string do_not(string iput)
        {
            string kq = String.Format(" NOT ({0} )", iput);
            return kq;
        }
        public string MyWhereClause()
        {
            string strW = "";
            if (_node.isEmpty())
            {
                return string.Empty;
            }
            if (IsNot == "Y" || Operate != "-")
            {
                if (Operate == "-")
                {
                    strW = do_not(CreateWhereOld(strW));
                }
                else
                {
                    strW = CreateWhereNew(strW);
                }
            }
            else
            {
                strW = CreateWhereOld(strW);
            }

            return strW;

        }

        private string CreateWhereNew(string strW)
        {
            string kq = "";
            string tmp = "";
            switch (Operate)
            {
                case "=":
                    strW = NormalFilter(strW);
                    break;
                case "BEGIN":
                    tmp = _node.FormatMyParameter(ValueFrom);
                    if (_node.FType == "")
                    {
                        tmp = tmp.Substring(0, tmp.Length - 1) + "%'";
                        kq = String.Format("( UPPER({0}) LIKE {1} )", _node.FormatMe(), tmp.ToUpper());
                    }
                    break;
                case "END":
                    tmp = _node.FormatMyParameter(ValueFrom);
                    if (_node.FType == "")
                    {
                        int tmpIndex = tmp.IndexOf("'");
                        if (tmpIndex >= 0)
                        {
                            tmp = String.Format("{0}%{1}", tmp.Substring(0, tmpIndex + 1), tmp.Substring(tmpIndex + 1));
                            kq = String.Format("( UPPER({0}) LIKE {1} )", _node.FormatMe(), tmp.ToUpper());
                        }
                    }
                    break;
                case "CONTAIN":

                    string[] arrVal = ValueFrom.Split(' ');
                    for (int i = 0; i < arrVal.Length; i++)
                    {
                        tmp = _node.FormatMyParameter(String.Format("%{0}%", arrVal[i]));
                        if (_node.FType == "")
                        {
                            kq += String.Format("AND UPPER({0}) LIKE {1} ", _node.FormatMe(), tmp.ToUpper());
                        }
                    }
                    if (kq.Length > 3)
                        kq = String.Format("({0})", kq.Substring(3));
                    break;
                case "BETWEEN":
                    if (_node.FType == "")
                    {
                        kq = String.Format("( UPPER({0}) BETWEEN {1} AND {2} )", _node.FormatMe(), _node.FormatMyParameter(ValueFrom).ToUpper(), _node.FormatMyParameter(ValueTo).ToUpper());
                    }
                    else kq = String.Format("({0} BETWEEN {1} AND {2})", _node.FormatMe(), _node.FormatMyParameter(ValueFrom), _node.FormatMyParameter(ValueTo));

                    break;
                case "<":
                case ">":
                case "<=":
                case ">=":
                case "<>":
                    if (_node.FType == "")
                    {
                        kq = String.Format("( UPPER({0}) {1} {2} )", _node.FormatMe(), Operate, _node.FormatMyParameter(ValueFrom).ToUpper());
                    }
                    else
                        kq = String.Format("({0} {1} {2})", _node.FormatMe(), Operate, _node.FormatMyParameter(ValueFrom));
                    break;
                case "in":
                    if (_node.FType == "")
                    {
                        kq = String.Format("( UPPER({0}) in {1})", _node.FormatMe(), _node.FormatMyArrayParameter(ValueFrom).ToUpper());
                    }
                    else kq = String.Format("({0} in {1})", _node.FormatMe(), _node.FormatMyArrayParameter(ValueFrom));
                    break;
                case "SPACE":
                    if (_node.FType == "")
                    {
                        kq = String.Format("({0} = '' OR  {0} is null )", _node.FormatMe());
                    }
                    else kq = String.Format("({0} is null )", _node.FormatMe());
                    break;
                case "EXISTS":
                    if (_node.FType == "")
                    {
                        kq = String.Format("({0} <> '' AND  {0} is not null )", _node.FormatMe());
                    }
                    else kq = String.Format("({0} is not null )", _node.FormatMe());
                    break;
                default:
                    kq = NormalFilter(strW);
                    break;

            }
            if (kq == "")
                kq = NormalFilter(strW);
            if (IsNot == "Y")
            {
                return do_not(kq);
            }

            return kq;
        }

        private string NormalFilter(string strW)
        {
            if (_node.FType == "")
                strW = String.Format("( UPPER({0}) = {1})", _node.FormatMe(), _node.FormatMyParameter(ValueFrom).ToUpper());
            else
                strW = String.Format("({0} = {1})", _node.FormatMe(), _node.FormatMyParameter(ValueFrom));
            return strW;
        }

        private string CreateWhereOld(string strW)
        {
            // using wildcard
            if (usingWildCard())
            {
                // eliminate wildcard signature
                string _wildCard = ValueFrom.Trim();
                if (_wildCard.Length >= 2 && _wildCard.Substring(0, 2) == "*>")
                {
                    if (string.IsNullOrEmpty(ValueTo))
                    {
                        strW = String.Format("({0} = {1})", _node.FormatMe(), _node.FormatMyParameter(ValueFrom.Replace("*>", "")));
                    }
                }
                else
                {
                    bool flag_not = false;
                    if (Regex.IsMatch(_wildCard, @"^(\^)"))
                    {
                        flag_not = true;
                        _wildCard = _wildCard.Replace(_wildCard[0], ' ').Trim();
                    }
                    //if (_node.FType == string.Empty)
                    //{
                    if (flag_not == false)
                        strW = GetWhereClause(_wildCard);
                    else
                    {

                        strW = do_not(GetWhereClause(_wildCard));
                    }
                }

                //}
                //else// if (_node.FType[0] == 'N')
                //{
                //    _wildCard = Regex.Replace(_wildCard, @"\*|\s", string.Empty);
                //    return "(" + _node.FormatMe() + _node.FormatMyParameter(_wildCard) + " )";
                //}
                //return "()";
            }
            // no input in filterTo
            else
                if (string.IsNullOrEmpty(ValueTo))
                {
                    strW = String.Format("({0} = {1})", _node.FormatMe(), _node.FormatMyParameter(ValueFrom));
                }
                else
                    strW = String.Format("({0} BETWEEN {1} AND {2} )", _node.FormatMe(), _node.FormatMyParameter(ValueFrom), _node.FormatMyParameter(ValueTo));


            return strW;
        }

        private string GetWhereClause(string _wildCard)
        {
            string strOperator = Regex.Replace(_wildCard, @"[0-9a-zA-Z,\-]", string.Empty).Trim();
            string strParam = Regex.Replace(_wildCard, @"\=|\s|\>|\<|\*>", string.Empty).Trim();
            //   _wildCard = Regex.Replace(_wildCard, @"\*|\=|\s|\>|\<", string.Empty);
            // Like clause support only text node

            if (Regex.IsMatch(strOperator, @"^(<>)"))
            {
                strOperator = "<>";
                strParam = Regex.Replace(_wildCard, @"^(<>)", string.Empty).Trim();

            }
            else if (Regex.IsMatch(strOperator, @"^(>>)"))
            {
                strOperator = ">>";
                strParam = Regex.Replace(_wildCard, @"^(>>)", string.Empty).Trim();
            }
            else if (Regex.IsMatch(strOperator, @"^(<<)"))
            {
                strOperator = "<<";
                strParam = Regex.Replace(_wildCard, @"^(<<)", string.Empty).Trim();
            }
            else if (Regex.IsMatch(strOperator, @"(%|_|\[|\])"))
            {
                strOperator = "like";
            }

            switch (strOperator)
            {
                case "!":
                    if (strParam == "!")
                    {
                        return String.Format("(({0} is null) or ( ltrim ({0}) = ''))", _node.FormatMe());
                    }
                    else
                        return String.Format("({0} LIKE (N'{1}') )", _node.FormatMe(), _wildCard);
                    break;
                case ">>":
                    return String.Format("(Upper({0}) LIKE Upper(N'{1}') )", _node.FormatMe(), strParam);
                    break;
                case "<<":
                    string[] test = strParam.Split(',');
                    string result = "";
                    for (int i = 0; i < test.Length; i++)
                    {
                        result += _node.FormatMyParameter(test[i]);
                        if (i != test.Length - 1)
                            result += ",";
                    }

                    return String.Format("({0} in ({1}) )", _node.FormatMe(), result);
                    break;
                case "like":
                    if (Regex.IsMatch(_wildCard, @"(^*=)"))
                        return String.Format("({0} LIKE (N'{1}') )", _node.FormatMe(), _wildCard.Substring(2));
                    return String.Format("({0} LIKE (N'{1}') )", _node.FormatMe(), _wildCard);
                    break;
                case "<>":
                    return String.Format("({0}{1}{2} )", _node.FormatMe(), strOperator, _node.FormatMyParameter(strParam));
                    break;
                default:
                    return String.Format("({0}{1}{2} )", _node.FormatMe(), strOperator, _node.FormatMyParameter(strParam));
                    break;
            }
        }


        #region '"Evaluate Cell Addresses"'
        /* L:123 */
        public Delegation.EvaluateCell _eval;

        #endregion

        //  Method CompareTo
        public int CompareTo(object obj)
        {
            if (obj == null)
            {
                return 1;
            }
            else
            {
                Filter fy = ((Filter)(obj));

                return String.Compare(Node.Code, fy.Node.Code);
                //{
                //    return 1;
                //}
                //if (this.Node.Code == fy.Node.Code)
                //{
                //    return 0;
                //}
                //if (this.Node.Code < fy.Node.Code)
                //{
                //    return -1;
                //}
            }
            return 0;
        }
        // interface methods implemented by CompareTo
        int System.IComparable.CompareTo(object obj)
        {
            return CompareTo(obj);
        }


    }
}
