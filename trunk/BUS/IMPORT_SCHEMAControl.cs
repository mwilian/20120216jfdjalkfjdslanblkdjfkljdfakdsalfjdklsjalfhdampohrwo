using System;
using System.Collections.Generic;
using System.Text;
using DTO;
using DAO;
using System.Data;
using System.Xml;
using System.Data.OleDb;
namespace BUS
{
    /// <summary> 
    ///Author: nnamthach@gmail.com 
    /// <summary>
    public class IMPORT_SCHEMAControl
    {
        #region Local Variable
        string _strConn = "";

        public string StrConn
        {
            get { return _strConn; }
            set { _strConn = value; }
        }
        DataTable _dtStruct;

        public DataTable DtStruct
        {
            get { return _dtStruct; }
            set { _dtStruct = value; }
        }
        List<ValueList> _listV = new List<ValueList>();

        public List<ValueList> ListV
        {
            get { return _listV; }
            set { _listV = value; }
        }
        List<string> _lKey = new List<string>();

        public List<string> LKey
        {
            get { return _lKey; }
            set { _lKey = value; }
        }
        private IMPORT_SCHEMADataAccess _objDAO;
        string _lookup = "";

        public string Lookup
        {
            get { return _lookup; }
            set { _lookup = value; }
        }
        #endregion Local Variable

        #region Method
        public IMPORT_SCHEMAControl()
        {
            _objDAO = new IMPORT_SCHEMADataAccess();
        }

        public IMPORT_SCHEMAInfo Get(
        String CONN_ID,
        String SCHEMA_ID,
        ref string sErr)
        {
            return _objDAO.Get(
            CONN_ID,
            SCHEMA_ID,
            ref sErr);
        }

        public DataTable GetAll(string con,
        ref string sErr)
        {
            return _objDAO.GetAll(con,
            ref sErr);
        }

        public Int32 Add(IMPORT_SCHEMAInfo obj, ref string sErr)
        {
            return _objDAO.Add(obj, ref sErr);
        }

        public string Update(IMPORT_SCHEMAInfo obj)
        {
            return _objDAO.Update(obj);
        }

        public string Delete(
        String CONN_ID,
        String SCHEMA_ID
        )
        {
            return _objDAO.Delete(
            CONN_ID,
            SCHEMA_ID
            );
        }
        public Boolean IsExist(
        String CONN_ID,
        String SCHEMA_ID
        )
        {
            return _objDAO.IsExist(
            CONN_ID,
            SCHEMA_ID
            );
        }

        public DataTableCollection Get_Page(IMPORT_SCHEMAInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            return _objDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }

        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {
            return _objDAO.Search(columnName, columnValue, condition, tableName, ref  sErr);
        }
        public string InsertUpdate(IMPORT_SCHEMAInfo obj)
        {
            string sErr = "";
            if (IsExist(
            obj.CONN_ID,
            obj.SCHEMA_ID
            ))
            {
                sErr = Update(obj);
            }
            else
                Add(obj, ref sErr);
            return sErr;
        }

        public DataTable GetTransferOut(string dtb, object from, object to, ref string sErr)
        {
            return _objDAO.GetTransferOut(dtb, from, to, ref sErr);
        }

        public DataTable ToTransferInStruct()
        {
            IMPORT_SCHEMAInfo inf = new IMPORT_SCHEMAInfo();
            return inf.ToDataTable();
        }

        public string TransferIn(DataRow row)
        {
            IMPORT_SCHEMAInfo inf = new IMPORT_SCHEMAInfo(row);
            return InsertUpdate(inf);
        }
        #endregion Method

        public static DataTable GetStruct(string xmlStr)
        {
            DataTable _dtField = new DataTable("field");
            DataColumn[] colfield = new DataColumn[] {  new DataColumn("PrimaryKey")
                ,new DataColumn("Key")
                , new DataColumn("DataTypeCode")
                , new DataColumn("Caption")
                , new DataColumn("DataMember")
                 , new DataColumn("AggregateFunction")
                , new DataColumn("Position")
                , new DataColumn("IsNull")
                , new DataColumn("Visible")
                , new DataColumn("Tag")};
            _dtField.Columns.AddRange(colfield);
            return GetDataTableFromXML(_dtField, xmlStr);
        }
        public static DataTable GetDataTableFromXML(DataTable dataTable, string p)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(p);
            dataTable.Rows.Clear();
            XmlNodeList nodeL = xmlDoc.GetElementsByTagName("Columns");
            foreach (XmlElement node in nodeL)
            {

                foreach (XmlElement ele in node.ChildNodes)
                {
                    DataRow row = dataTable.NewRow();
                    foreach (XmlElement child in ele.ChildNodes)
                    {
                        if (child.InnerText != "")
                            row[child.Name] = child.InnerText;
                    }

                    dataTable.Rows.Add(row);
                }

            }
            return dataTable;
        }
        public static string GetXMLFromDataTable(DataTable dataTable)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\"?><GridEXLayoutFile LayoutType=\"Janus.Windows.GridEX.GridEX\" LayoutVersion=\"1.1\">  <RootTable>    <Key>TSHInfoList</Key>    <Caption>TSHInfoList</Caption>    <Columns Collection=\"true\" ElementName=\"Column\">          </Columns>    <GroupCondition />  </RootTable>  <SearchColumnIndex>0</SearchColumnIndex>     <VisualStyle>Office2007</VisualStyle>    <RowHeaders>True</RowHeaders>  </GridEXLayoutFile>");
            XmlNodeList eleColumns = xmlDoc.GetElementsByTagName("Columns");
            foreach (XmlElement node in eleColumns)
            {
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    XmlElement ele = xmlDoc.CreateElement("Column" + i);
                    ele.SetAttribute("ID", dataTable.Rows[i]["Key"].ToString());
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        XmlElement child = xmlDoc.CreateElement(col.ColumnName);
                        child.InnerXml = dataTable.Rows[i][col].ToString();
                        ele.AppendChild(child);
                    }
                    node.AppendChild(ele);
                }
            }
            return xmlDoc.InnerXml;
        }

        public static DataTable GetDataTableStruct(DataTable dtStruct, string lookup)
        {
            DataTable dt = new DataTable(lookup);
            foreach (DataRow row in dtStruct.Rows)
            {
                dt.Columns.Add(row["Key"].ToString(), Type.GetType("System." + row["DataTypeCode"].ToString()));
                dt.Columns[row["Key"].ToString()].ExtendedProperties.Add("DataMember", row["DataMember"].ToString());
            }
            return dt;
        }

        public bool ContrainList(string key, object value, ref string message)
        {
            foreach (ValueList x in _listV)
            {
                if (x.Key == key)
                {
                    if (value == DBNull.Value && x.IsNull == false)
                    {
                        message = "Value is not null";
                        return false;
                    }
                    else
                    {
                        if (x.Content.Count == 0)
                            return true;
                        else
                        {
                            if (x.Content.Contains(value.ToString()))
                                return true;
                            else
                            {
                                message = x.Message;
                                return false;
                            }
                        }
                    }
                }
            }
            return true;
        }


        public string Import(DataTable dataTable, bool insert, bool update)
        {
            string sErr = "";
            foreach (DataRow row in dataTable.Rows)
                sErr += Import(row, insert, update, ref sErr) + "\\";
            return sErr;
        }
        public int Import(DataRow row, bool insert, bool update, ref string sErr)
        {
            if (insert && update)
                return ImportInsertUpdate(row, ref sErr);
            if (!insert && update)
                return ImportUpdate(row, ref sErr);
            if (insert && !update)
                return ImportInsert(row, ref sErr);
            return 0;
        }

        private int ImportInsert(DataRow row, ref string sErr)
        {
            string exist = CreateExistString(row);
            string query = CreateInsertString(row);
            if (exist != "")
                query = String.Format("IF NOT {0} BEGIN {1} SELECT 1 END ELSE SELECT 0", exist, query);
            else
            {
                query = String.Format("{0} SELECT 1", query);
            }
            return ExecQuery(query, ref sErr);
        }



        private int ImportUpdate(DataRow row, ref string sErr)
        {
            string exist = CreateExistString(row);
            string query = CreateUpdatetString(row);
            if (exist != "")
                query = String.Format("IF {0} BEGIN {1} SELECT 1 END ELSE SELECT 0", exist, query);
            else query = CreateInsertString(row) + " SELECT 1";
            return ExecQuery(query, ref sErr);
        }
        private int ImportInsertUpdate(DataRow row, ref string sErr)
        {
            string exist = CreateExistString(row);
            string insert = CreateInsertString(row);
            string update = CreateUpdatetString(row);
            if (exist != "")
                insert = String.Format("IF {0} BEGIN {1} SELECT 1 END ELSE BEGIN {2} SELECT 1 END", exist, update, insert);
            else
            {
                insert = String.Format("{0} SELECT 1", insert);
            }
            return ExecQuery(insert, ref sErr);
        }
        private string CreateInsertString(DataRow row)
        {
            string field = "";
            string values = "";
            foreach (DataColumn col in row.Table.Columns)
            {
                string fieldName = col.ColumnName;
                if (col.ExtendedProperties.ContainsKey("DataMember"))
                    fieldName = col.ExtendedProperties["DataMember"].ToString();

                field += ",[" + fieldName + "]";
                if (row[col] == DBNull.Value)
                {
                    values += ",NULL";
                }
                else
                {
                    if (col.DataType == typeof(String))
                        values += ",N'" + row[col].ToString() + "'";
                    else if (col.DataType == typeof(DateTime))
                        values += ",'" + ((DateTime)row[col]).Year + "-" + ((DateTime)row[col]).Month + "-" + ((DateTime)row[col]).Day + " " + ((DateTime)row[col]).ToLongTimeString() + "'";
                    else
                        values += "," + row[col].ToString();
                }
            }
            field = field.Substring(1);
            values = values.Substring(1);
            string query = String.Format("INSERT INTO " + _lookup + "({0}) VALUES({1})", field, values);
            return query;
        }

        private string CreateExistString(DataRow row)
        {
            if (_lKey.Count == 0)
                return "";
            string kq = "EXISTS(SELECT {0} FROM " + _lookup + " WHERE {1})";
            string field = "";
            string values = "";
            foreach (DataColumn col in row.Table.Columns)
            {
                string fieldName = col.ColumnName;
                if (col.ExtendedProperties.ContainsKey("DataMember"))
                    fieldName = col.ExtendedProperties["DataMember"].ToString();
                if (_lKey.Contains(fieldName))
                {
                    field += ",[" + fieldName + "]";
                    if (row[col] == DBNull.Value)
                    {
                        values += "And [" + fieldName + "]= NULL";// values += ",NULL";
                    }
                    else
                    {
                        if (col.DataType == typeof(String))
                            values += "And [" + fieldName + "]=N'" + row[col].ToString() + "'";
                        else if (col.DataType == typeof(DateTime))
                            values += "And [" + fieldName + "]=N'" + ((DateTime)row[col]).Year + "-" + ((DateTime)row[col]).Month + "-" + ((DateTime)row[col]).Day + " " + ((DateTime)row[col]).ToLongTimeString() + "'";
                        else
                            values += "And [" + fieldName + "]=" + row[col].ToString();
                    }
                }
            }
            if (field.Length > 1)
                field = field.Substring(1);
            if (values.Length > 3)
                values = values.Substring(3);
            kq = String.Format(kq, field, values);
            return kq;
        }
        private string CreateUpdatetString(DataRow row)
        {
            if (_lKey.Count == 0)
                return "";
            string kq = "UPDATE " + _lookup + " SET {0} WHERE {1}";
            string field = "";
            string where = "";
            foreach (DataColumn col in row.Table.Columns)
            {
                string fieldName = col.ColumnName;
                if (col.ExtendedProperties.ContainsKey("DataMember"))
                    fieldName = col.ExtendedProperties["DataMember"].ToString();
                if (!_lKey.Contains(fieldName))
                {
                    if (row[col] == DBNull.Value)
                    {

                    }
                    else
                    {
                        if (col.DataType == typeof(String))
                            field += ", [" + fieldName + "]=N'" + row[col].ToString() + "'";
                        else if (col.DataType == typeof(DateTime))
                            field += ", [" + fieldName + "]=N'" + ((DateTime)row[col]).Year + "-" + ((DateTime)row[col]).Month + "-" + ((DateTime)row[col]).Day + " " + ((DateTime)row[col]).ToLongTimeString() + "'";
                        else
                            field += ", [" + fieldName + "]=" + row[col].ToString();
                    }
                }
                else
                {
                    if (row[col] == DBNull.Value)
                    {

                    }
                    else
                    {

                        if (col.DataType == typeof(String))
                            where += "And [" + fieldName + "]=N'" + row[col].ToString() + "'";
                        else if (col.DataType == typeof(DateTime))
                            where += "And [" + fieldName + "]=N'" + ((DateTime)row[col]).Year + "-" + ((DateTime)row[col]).Month + "-" + ((DateTime)row[col]).Day + " " + ((DateTime)row[col]).ToLongTimeString() + "'";
                        else
                            where += "And [" + fieldName + "]=" + row[col].ToString();
                    }

                }
            }
            if (field.Length > 1)
                field = field.Substring(1);
            if (where.Length > 3)
                where = where.Substring(3);
            kq = String.Format(kq, field, where);
            return kq;
        }

        private int ExecQuery(string query, ref string sErr)
        {
            int result = 0;
            using (OleDbConnection conn = new OleDbConnection(_strConn))
            {
                try
                {
                    OleDbCommand command = new OleDbCommand(query, conn);
                    conn.Open();
                    result = Convert.ToInt32(command.ExecuteScalar());
                    conn.Close();
                }
                catch (Exception ex) { sErr = ex.Message; }
            }
            return result;
        }
    }

    public class ValueList
    {
        string _key = "";

        public string Key
        {
            get { return _key; }
            set { _key = value; }
        }
        List<String> _content = new List<string>();

        public List<String> Content
        {
            get { return _content; }
            set { _content = value; }
        }
        string _message = "";

        public string Message
        {
            get { return _message; }
            set { _message = value; }
        }
        bool _isNull = true;

        public bool IsNull
        {
            get { return _isNull; }
            set { _isNull = value; }
        }
    }
}
