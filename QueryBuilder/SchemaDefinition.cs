using System.ComponentModel;
using System.Xml;
using System.IO;
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
    //  Class TableItem
    public class TableItem
    {
        private string _code = string.Empty;
        private string _lookup = string.Empty;
        private string _description = string.Empty;

        //  Property Code
        public string Code
        {
            get
            {
                return _code;
            }
        }
        //  Property Lookup
        public string Lookup
        {
            get
            {
                return _lookup;
            }
            set { _lookup = value; }
        }
        //  Property Description
        public string Description
        {
            get
            {
                return _description;
            }

        }
        public TableItem(string code, string descr)
        {
            _code = code;
            _description = descr;
        }
    }

    //  Class SchemaDefinition
    public class SchemaDefinition
    {
        #region '" Constants "'
        //  Inherits Dictionary(Of String, String)
        private const string JoinsDefinition = "FROM";
        private const string AliasDefinition = "_ALIAS";
        private const string STR_Table = "table";
        private const string STR_Node = "node";
        private const string STR_Name = "name";
        private const string STR_NodeDesc = "nodeDesc";
        /* L:40 */
        private const string STR_Type = "type";
        /* L:40 */
        private const string STR_SUM = "SUM";
        private const string STR_COUNT = "COUNT";
        private const string STR_MAX = "MAX";

        private const string STR_isNumber = "N";
        private const string STR_isCOUNTABLE = "N2";
        private const string STR_isMAXMIN = "N3";
        private const string STR_isSubNode = "S";
        #endregion

        #region '" Shared service Functions "'

        private static BindingList<Node> cachedTable = null;
        private static string cachedTableName = string.Empty;
        //  Method InvalidateTable
        public void InvalidateTable()
        {
            cachedTable = null;
        }


        ///  <summary>
        ///  Load whole Dictionary, related to LA , or CA ,or AR ...
        ///  </summary>
        ///  <param name="table"> Table Code exl LA CA</param>
        ///  <returns>A binding list of Node . which can bind to a grid or tree using Loadtree</returns>
        ///  <remarks></remarks>
        private static BindingList<Node> GetTable(string dtb, string table)
        {

            // filling new Tabble
            BindingList<Node> _schema = new BindingList<Node>();

            StringReader stream = null;
            XmlTextReader reader = null;

            // Use origin if table is alias.
            string origin = GetOriginFromAlias(table);
            try
            {
                // read resource file. for ex "LA"
                stream = new StringReader(GetTableSchema(dtb, origin)); // if table = ICAS - load from origin CA
                reader = new XmlTextReader(stream);

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (reader.GetAttribute(STR_Table) == origin)
                        { // Origin = CA

                            string nodeCode = reader.GetAttribute(STR_Node);

                            string nodeName = reader.GetAttribute(STR_Name);

                            string nodeType = reader.GetAttribute(STR_Type);

                            string nodeDesc = reader.GetAttribute(STR_NodeDesc) == null ? "" : reader.GetAttribute(STR_NodeDesc);

                            string nodeAgregate = null; // ----------------make auto agreagate function here

                            switch (nodeType)
                            {
                                case STR_isNumber:
                                    nodeAgregate = Parsing.STR_SUM;
                                    break;
                                case STR_isCOUNTABLE:
                                    nodeAgregate = Parsing.STR_COUNT;
                                    break;
                                case STR_isMAXMIN:
                                    nodeAgregate = Parsing.STR_MAXIMUM;
                                    break;
                                default:
                                    nodeAgregate = string.Empty;
                                    break;
                            }
                            //  -----------------------------------finish agregating

                            _schema.Add(new Node(nodeAgregate, table + @"\" + nodeCode, System.Convert.ToString(nodeName), System.Convert.ToString(nodeType), System.Convert.ToString(nodeDesc)));

                            if (System.Convert.ToString(nodeType) == STR_isSubNode)
                            { //  for example "LA\CA"

                                foreach (Node _node in GetTableByCode(dtb, System.Convert.ToString(nodeCode)))
                                {
                                    _node.AddMeToParent(table); // assign code and tree code to subnode. ex CA\NA --> LA\CA\NA
                                    _schema.Add(_node);
                                }
                            }

                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw new ArgumentException("Missing Dictionary schema : " + table);
            }
            finally
            {
                if (!(reader == null))
                {
                    reader.Close();
                }
            }
            return _schema;
        }




        //  Method GetTableByCode
        public static BindingList<Node> GetTableByCode(string dtb, string table)
        {
            // if context changed . invalidate cached Table
            /* L:123 */
            if (cachedTable == null || (!(table.Equals(cachedTableName))))
            {
                cachedTableName = table;
                cachedTable = GetTable(dtb, table);
            }
            return cachedTable;
        }


        ///  <summary>
        ///  Return GetTablebyCode but decorate analysis Code and Description
        ///  Also remove unused analysis Category
        ///  </summary>
        ///  <param name="table"></param>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public static BindingList<Node> GetDecorateTableByCode(string table, string DB)
        {
            BindingList<Node> _schema = new BindingList<Node>();
            BindingList<Node> _unUsedAnalysisCategories = new BindingList<Node>();
            _schema = GetTableByCode(DB, table);
            // decoration
            int AnalCnt = 0;
            string cat = string.Empty;
            string catNode = string.Empty;
            // Dim ANAL_Regex As String = "(?<=ANAL_)[A,T,M,I,C,Q,F][0-9][0-9]|(?<=ANAL)[A,T,M,I,C,Q,F][0-9][0-9]|(?<=[^A-Za-z0-90-9])[A,T,M,I,C,Q,F][0-9][0-9]|(?<=Anal)[A,T,M,I,C,Q,F][0-9][0-9]"
            string ANAL_Regex = "[A,T,M,I,Q,F][0-9][0-9]";

            foreach (Node F in _schema)
            {
                // cat = Regex.Match(F.Code, pbs.Helper.pbsRegex.DB_ANAL_HEADER).ToString
                // catNode = Regex.Match(F.MyCode, pbs.Helper.pbsRegex.DB_ANAL_HEADER).ToString

                // cat = Regex.Match(F.Code, ANAL_Regex).ToString
                cat = Regex.Match(F.Code, ANAL_Regex).ToString();
                catNode = Regex.Match(F.MyCode, ANAL_Regex).ToString();

                if (!(string.IsNullOrEmpty(cat)))
                {
                    string transTemp0 = F.MyCode;
                    if (  /* TRANSINFO: .NET Equivalent of Microsoft.VisualBasic NameSpace */ transTemp0.Length == 3)
                    { // tam thoi, day la ma phan tich

                        AnalCnt = 0;
                        NDInfo info = NDInfo.EmptyNDInfo();
                        if (NDInfoList.GetNDInfoList(DB).ContainsCode(cat, ref info))
                        {
                            // F.DecorateDescriptn(catNode)
                            F.DecorateDescriptn(info.Description);

                        }
                    }
                    else
                    {

                        // decorate cac analysis code va description ... ben trong tung node analysis
                        AnalCnt = AnalCnt + 1;
                        if (AnalCnt > 1)
                        {
                            F.DecorateDescriptn(cat + " " + F.Description);
                        }

                    }
                    NDInfo transTemp1 = null;
                    if (!(NDInfoList.GetNDInfoList(DB).ContainsCode(cat, ref transTemp1)))
                    {
                        _unUsedAnalysisCategories.Add(F);
                        // If F.Description = "" Or F.Description = "<Blank>" Then
                    }
                    else
                    {
                        // F.DecorateDescriptn(ResStr(F.MyCode))
                    }
                }

            }
            // remove unused analysis code
            foreach (Node _f in _unUsedAnalysisCategories)
            {
                _schema.Remove(_f);
            }
            return _schema;
        }


        ///  <summary>
        ///  Read List.XML and show which table can be select from dictionary
        ///  </summary>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public static Dictionary<string, string> GetUsableCodeList(string dtb, ref Dictionary<string, string> dcLookup)
        {
            if (_list == null)
            {
                dcLookup = new Dictionary<string, string>();
                _list = new Dictionary<string, string>();
                StringReader stream = null;
                XmlTextReader reader = null;

                try
                {
                    stream = new StringReader(GetSchemaList(dtb));
                    reader = new XmlTextReader(stream);
                    string table = string.Empty;
                    string Description = string.Empty;
                    string lookup = string.Empty;

                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element & reader.Name == "row")
                        {
                            table = reader.GetAttribute("table");
                            Description = reader.GetAttribute("Description");
                            lookup = reader.GetAttribute("lookup");
                            _list.Add(table, Description);
                            dcLookup.Add(table, lookup);
                        }
                    }

                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Cannot Load Dictionary" + ex.Message);
                }
                finally
                {
                    if (!(reader == null))
                    {
                        reader.Close();
                    }
                }

            }
            return _list;
        }




        //  Method GetTableList
        public static BindingList<TableItem> GetTableList(string dtb)
        {
            BindingList<TableItem> bl = new System.ComponentModel.BindingList<TableItem>();
            Dictionary<string, string> dcLookup = null;
            foreach (KeyValuePair<string, string> item in GetUsableCodeList(dtb, ref dcLookup))
            {
                TableItem ti = new TableItem(item.Key, item.Value);
                string lookup = "";
                if (dcLookup.TryGetValue(item.Key, out lookup))
                {
                    ti.Lookup = lookup;
                }
                bl.Add(ti);
            }
            return bl;
        }


        ///  <summary>
        ///  Give join code LA/CA , get :  "LA.ACCNT_CODE = CA.ACCNT_CODE"/>
        ///  or give CA get : SSRFACC where SUN_DB = 'XXX'
        ///  </summary>
        ///  <param name="joinCode">give : fromcode ="T0" </param>
        ///  <returns>lookup="(SELECT * FROM SSRFANV WHERE CATEGORY = 'T0' AND SUN_DB = 'XXX')"/</returns>
        ///  <remarks></remarks>
        public static string GetJoin(string dtb, string joinCode)
        {
            if (GetJoinsDictionary(dtb).ContainsKey(joinCode))
            {
                Dictionary<string, string> temp = new Dictionary<string, string>();
                temp = GetJoinsDictionary(dtb);
                return temp[joinCode];
            }
            else
            {
                return joinCode;
            }
        }


        ///  <summary>
        ///  row alias="ICAS" origin="CA"
        ///  </summary>
        ///  <param name="_alias"></param>
        ///  <returns></returns>
        ///  <remarks></remarks>
        public static string GetOriginFromAlias(string _alias)
        {
            if (GetAliasDictionary().ContainsKey(_alias))
            {
                Dictionary<string, string> temp = new Dictionary<string, string>();
                temp = GetAliasDictionary();

                return temp[_alias];
            }
            else
            {
                return _alias;
            }
        }

        #endregion

        #region '"Dictionary Building"'

        private static Dictionary<string, string> _list = null;
        private static Dictionary<string, string> _joinsDictionary = null;

        public static Dictionary<string, string> JoinsDictionary
        {
            get { return SchemaDefinition._joinsDictionary; }
            set { SchemaDefinition._joinsDictionary = value; }
        }
        private static Dictionary<string, string> _aliasDictionary = null;
        //  Method InvalidateCache
        public static void InvalidateCache()
        {
            _joinsDictionary = null;
            _aliasDictionary = null;
            _list = null;
        }


        private SchemaDefinition()
            : base()
        {

        }

        ///  <summary>
        ///  Read FROM.XML and load joins definition to Schema dictionary
        ///  </summary>
        ///  <returns></returns>
        ///  <remarks></remarks>
        private static Dictionary<string, string> GetJoinsDictionary(string dtb)
        {
            if (_joinsDictionary == null)
            {
                _joinsDictionary = new Dictionary<string, string>();

                StringReader stream = null;
                XmlTextReader reader = null;

                try
                {
                    stream = new StringReader(GetJoinSchema(dtb));
                    reader = new XmlTextReader(stream);
                    string fromcode = string.Empty;
                    string lookup = string.Empty;

                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element & reader.Name == "row")
                        {
                            fromcode = reader.GetAttribute("fromcode");
                            lookup = reader.GetAttribute("lookup");
                            _joinsDictionary.Add(fromcode, lookup);
                        }
                    }

                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Joins Dictionary Building Problem" + ex.Message);
                }
                finally
                {
                    if (!(reader == null))
                    {
                        reader.Close();
                    }
                }
            }
            return _joinsDictionary;
        }




        ///  <summary>
        ///  Read ALIAS.XML and load ALIAS - ORIGIN NAME
        ///  </summary>
        ///  <returns></returns>
        ///  <remarks></remarks>
        private static Dictionary<string, string> GetAliasDictionary()
        {
            if (_aliasDictionary == null)
            {
                _aliasDictionary = new Dictionary<string, string>();

                StringReader stream = null;
                XmlTextReader reader = null;

                try
                {
                    stream = new StringReader(GetAliasChema());
                    reader = new XmlTextReader(stream);
                    string aliascode = string.Empty;
                    string origin = string.Empty;

                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.Element & reader.Name == "row")
                        {
                            aliascode = reader.GetAttribute("alias");
                            origin = reader.GetAttribute("origin");
                            _aliasDictionary.Add(aliascode, origin);
                        }
                    }

                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Alias Dictionary Building problem " + ex.Message);
                }
                finally
                {
                    if (!(reader == null))
                    {
                        reader.Close();
                    }
                }
            }
            return _aliasDictionary;
        }




        #endregion

        #region User defined Schemma
        private static string GetTableSchema(string dtb, string origin)
        {
            string sErr = "";
            string result = "";
            CoreQD_SCHEMAControl ctr = new CoreQD_SCHEMAControl();
            CoreQD_SCHEMAInfo inf = ctr.Get(dtb, origin, ref sErr);
            if (inf.SCHEMA_ID == "")
            {
                result = System.Convert.ToString(Properties.Resources.ResourceManager.GetObject(origin));
            }
            else
            {
                result = "<?xml version='1.0' encoding='utf-8' ?><SUN_SCHEMA></SUN_SCHEMA>";
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(result);
                XmlElement docele = doc.DocumentElement;

                //<row table="M5" node="Lookup" name="Lookup" type=""/>  	
                string schema = inf.FROM_TEXT;
                DataSet dset = new DataSet("Schema");
                //DataTable dtfrom = new DataTable("fromcode");
                //DataColumn[] colfrom = new DataColumn[] { new DataColumn("fromcode"), new DataColumn("lookup") };
                //dtfrom.Columns.AddRange(colfrom);

                DataTable dtfield = new DataTable("field");
                DataColumn[] colfield = new DataColumn[] { new DataColumn("node")
                    , new DataColumn("table")
                    , new DataColumn("name")
                    , new DataColumn("nodeDesc")
                    , new DataColumn("type")};
                dtfield.Columns.AddRange(colfield);

                //dset.Tables.Add(dtfrom);
                dset.Tables.Add(dtfield);
                //DataRelation relation = new DataRelation("R_field", dtfrom.Columns["fromcode"], dtfield.Columns["table"], true);
                //dset.Relations.Add(relation);
                StringReader strR = new StringReader(schema);
                dset.ReadXml(strR);
                strR.Close();
                //dset.ReadXml(schema);
                foreach (DataRow jrow in dtfield.Rows)
                {
                    XmlElement ele = doc.CreateElement("row");
                    ele.SetAttribute("table", jrow["table"].ToString());
                    ele.SetAttribute("node", jrow["node"].ToString());
                    ele.SetAttribute("name", jrow["name"].ToString());
                    ele.SetAttribute("nodeDesc", jrow["nodeDesc"] == null ? "" : jrow["nodeDesc"].ToString());
                    ele.SetAttribute("type", jrow["type"] == null ? "" : jrow["type"].ToString());
                    ele.SetAttribute("conn_id", inf.DEFAULT_CONN);
                    docele.AppendChild(ele);
                }


                result = doc.InnerXml;
            }
            return result;
        }
        private static string GetSchemaList(string dtb)
        {
            string sErr = "";
            string result = Properties.Resources.SCHEMALIST;
            CoreQD_SCHEMAControl ctr = new CoreQD_SCHEMAControl();
            DataTable dt = ctr.GetAll(dtb, ref sErr);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result);
            XmlElement docele = doc.DocumentElement;
            foreach (DataRow row in dt.Rows)
            {
                //<row table="AllMovement" Description="All Movement" module="AllMovement"/>
                if (row["SCHEMA_STATUS"].ToString().Trim() == "Y")
                {
                    XmlElement ele = doc.CreateElement("row");
                    ele.SetAttribute("table", row["SCHEMA_ID"].ToString());
                    ele.SetAttribute("Description", row["DESCRIPTN"].ToString());
                    ele.SetAttribute("module", row["SCHEMA_ID"].ToString());
                    ele.SetAttribute("lookup", row["LOOK_UP"].ToString());
                    docele.AppendChild(ele);
                }
            }
            result = doc.InnerXml;
            return result;
        }
        private static string GetJoinSchema(string dtb)
        {
            string sErr = "";
            string result = System.Convert.ToString(Properties.Resources.ResourceManager.GetObject(JoinsDefinition));
            CoreQD_SCHEMAControl ctr = new CoreQD_SCHEMAControl();
            DataTable dt = ctr.GetAll(dtb, ref sErr);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(result);
            XmlElement docele = doc.DocumentElement;
            foreach (DataRow irow in dt.Rows)
            {
                //<row fromcode ="Movement" lookup="(SELECT * FROM TVC_MFMOVXXX WHERE Not Hold='Y')"/>
                string schema = irow["FROM_TEXT"].ToString();
                DataSet dset = new DataSet("Schema");
                DataTable dtfrom = new DataTable("fromcode");
                DataColumn[] col = new DataColumn[] { new DataColumn("fromcode"), new DataColumn("lookup") };
                dtfrom.Columns.AddRange(col);
                dset.Tables.Add(dtfrom);
                StringReader strR = new StringReader(schema);
                dset.ReadXml(strR);
                strR.Close();
                foreach (DataRow jrow in dtfrom.Rows)
                {
                    XmlElement ele = doc.CreateElement("row");
                    ele.SetAttribute("fromcode", jrow["fromcode"].ToString());
                    ele.SetAttribute("lookup", jrow["lookup"].ToString());
                    docele.AppendChild(ele);
                }

            }
            result = doc.InnerXml;
            return result;
        }
        private static string GetAliasChema()
        {
            string result = System.Convert.ToString(Properties.Resources.ResourceManager.GetObject(AliasDefinition));
            return result;
        }
        #endregion User defined Schemma
    }







}
