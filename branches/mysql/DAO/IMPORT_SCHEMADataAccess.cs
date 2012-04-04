using System;
using System.Collections.Generic;
using System.Text;
using DTO;
using System.Data;
using System.Data.SqlClient;

namespace DAO
{
    /// <summary> 
    ///Author: nnamthach@gmail.com 
    /// <summary>
    public class IMPORT_SCHEMADataAccess : Connection
    {
        #region Local Variable
        private string _strSPInsertName = "procIMPORT_SCHEMA_add";
        private string _strSPUpdateName = "procIMPORT_SCHEMA_update";
        private string _strSPDeleteName = "procIMPORT_SCHEMA_delete";
        private string _strSPGetName = "procIMPORT_SCHEMA_get";
        private string _strSPGetAllName = "procIMPORT_SCHEMA_getall";
        private string _strSPGetPages = "procIMPORT_SCHEMA_getpaged";
        private string _strSPIsExist = "procIMPORT_SCHEMA_isexist";
        private string _strTableName = "IMPORT_SCHEMA";
        private string _strSPGetTransferOutName = "procIMPORT_SCHEMA_gettransferout";
        string prefix = "param";
        #endregion Local Variable

        #region Method
        public IMPORT_SCHEMAInfo Get(
        String CONN_ID,
        String SCHEMA_ID,
        ref string sErr)
        {
            IMPORT_SCHEMAInfo objEntr = new IMPORT_SCHEMAInfo();
            connect();
            InitSPCommand(_strSPGetName);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.CONN_ID.ToString(), CONN_ID);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.SCHEMA_ID.ToString(), SCHEMA_ID);

            DataTable list = new DataTable();
            try
            {
                list = executeSelectSP(command);
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();

            if (list.Rows.Count > 0)
                objEntr = (IMPORT_SCHEMAInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            IMPORT_SCHEMAInfo result = new IMPORT_SCHEMAInfo();
            result.CONN_ID = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.CONN_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][IMPORT_SCHEMAInfo.Field.CONN_ID.ToString()]));
            result.SCHEMA_ID = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.SCHEMA_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][IMPORT_SCHEMAInfo.Field.SCHEMA_ID.ToString()]));
            result.LOOK_UP = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.LOOK_UP.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][IMPORT_SCHEMAInfo.Field.LOOK_UP.ToString()]));
            result.DESCRIPTN = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.DESCRIPTN.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][IMPORT_SCHEMAInfo.Field.DESCRIPTN.ToString()]));
            result.FIELD_TEXT = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.FIELD_TEXT.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][IMPORT_SCHEMAInfo.Field.FIELD_TEXT.ToString()]));
            result.DB = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.DB.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][IMPORT_SCHEMAInfo.Field.DB.ToString()]));
            result.DAG = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.DAG.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][IMPORT_SCHEMAInfo.Field.DAG.ToString()]));
            result.SCHEMA_STATUS = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.SCHEMA_STATUS.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][IMPORT_SCHEMAInfo.Field.SCHEMA_STATUS.ToString()]));
            result.UPDATED = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.UPDATED.ToString()] == DBNull.Value ? 0 : Convert.ToInt32(dt.Rows[i][IMPORT_SCHEMAInfo.Field.UPDATED.ToString()]));
            result.ENTER_BY = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.ENTER_BY.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][IMPORT_SCHEMAInfo.Field.ENTER_BY.ToString()]));
            result.DEFAULT_CONN = (dt.Rows[i][IMPORT_SCHEMAInfo.Field.DEFAULT_CONN.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][IMPORT_SCHEMAInfo.Field.DEFAULT_CONN.ToString()]));

            return result;
        }

        public DataTable GetAll(string con,
        ref string sErr)
        {
            connect();
            InitSPCommand(_strSPGetAllName);
            AddParameter(prefix + "CONN_ID", con);
            DataTable list = new DataTable();
            try
            {
                list = executeSelectSP();
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();


            if (sErr != "") ErrorLog.SetLog(sErr);
            return list;
        }
        /// <summary>
        /// Return 1: Table is exist Identity Field
        /// Return 0: Table is not exist Identity Field
        /// Return -1: Erro
        /// </summary>
        /// <param name="tableName"></param>
        public Int32 Add(IMPORT_SCHEMAInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.CONN_ID.ToString(), objEntr.CONN_ID);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.SCHEMA_ID.ToString(), objEntr.SCHEMA_ID);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.LOOK_UP.ToString(), objEntr.LOOK_UP);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.DESCRIPTN.ToString(), objEntr.DESCRIPTN);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.FIELD_TEXT.ToString(), objEntr.FIELD_TEXT);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.DB.ToString(), objEntr.DB);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.DAG.ToString(), objEntr.DAG);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.SCHEMA_STATUS.ToString(), objEntr.SCHEMA_STATUS);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.UPDATED.ToString(), objEntr.UPDATED);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.ENTER_BY.ToString(), objEntr.ENTER_BY);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.DEFAULT_CONN.ToString(), objEntr.DEFAULT_CONN);

            try
            {
                //command.ExecuteNonQuery();
                object tmp = executeSPScalar();
                if (tmp != null && tmp != DBNull.Value)
                    ret = Convert.ToInt32(tmp);
                else
                    ret = 0;
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") ErrorLog.SetLog(sErr);

            return ret;
        }

        public string Update(IMPORT_SCHEMAInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.CONN_ID.ToString(), objEntr.CONN_ID);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.SCHEMA_ID.ToString(), objEntr.SCHEMA_ID);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.LOOK_UP.ToString(), objEntr.LOOK_UP);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.DESCRIPTN.ToString(), objEntr.DESCRIPTN);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.FIELD_TEXT.ToString(), objEntr.FIELD_TEXT);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.DB.ToString(), objEntr.DB);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.DAG.ToString(), objEntr.DAG);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.SCHEMA_STATUS.ToString(), objEntr.SCHEMA_STATUS);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.UPDATED.ToString(), objEntr.UPDATED);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.ENTER_BY.ToString(), objEntr.ENTER_BY);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.DEFAULT_CONN.ToString(), objEntr.DEFAULT_CONN);

            string sErr = "";
            try
            {
                excuteSPNonQuery();
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") ErrorLog.SetLog(sErr);
            return sErr;
        }

        public string Delete(
        String CONN_ID,
        String SCHEMA_ID
        )
        {
            connect();
            InitSPCommand(_strSPDeleteName);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.CONN_ID.ToString(), CONN_ID);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.SCHEMA_ID.ToString(), SCHEMA_ID);

            string sErr = "";
            try
            {
                excuteSPNonQuery();
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") ErrorLog.SetLog(sErr);
            return sErr;
        }

        public DataTableCollection Get_Page(IMPORT_SCHEMAInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            string whereClause = CreateWhereClause(obj);
            DataTableCollection dtList = null;
            connect();
            InitSPCommand(_strSPGetPages);

            AddParameter(prefix + "WhereClause", whereClause);
            AddParameter(prefix + "OrderBy", orderBy);
            AddParameter(prefix + "PageIndex", pageIndex);
            AddParameter(prefix + "PageSize", pageSize);

            try
            {
                dtList = executeCollectSelectSP();
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") ErrorLog.SetLog(sErr);
            return dtList;
        }

        public Boolean IsExist(
        String CONN_ID,
        String SCHEMA_ID
        )
        {
            connect();
            InitSPCommand(_strSPIsExist);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.CONN_ID.ToString(), CONN_ID);
           AddParameter(prefix +  IMPORT_SCHEMAInfo.Field.SCHEMA_ID.ToString(), SCHEMA_ID);

            string sErr = "";
            DataTable list = new DataTable();
            try
            {
                list = executeSelectSP();
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") ErrorLog.SetLog(sErr);
            if (list.Rows.Count == 1)
                return true;
            return false;
        }

        private string CreateWhereClause(IMPORT_SCHEMAInfo obj)
        {
            String result = "";

            return result;
        }

        public DataTable Search(string columnName, string columnValue, string condition, string tableName, ref string sErr)
        {
            string query = "select * from " + _strTableName + " where " + columnName + " " + condition + " " + columnValue;
            DataTable list = new DataTable();
            connect();
            try
            {
                list = executeSelectQuery(query);
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            //    if (sErr != "") ErrorLog.SetLog(sErr);
            return list;
        }
        public DataTable GetTransferOut(string dtb, object from, object to, ref string sErr)
        {
            connect();
            InitSPCommand(_strSPGetTransferOutName);
           AddParameter(prefix +  "DB", dtb);
           AddParameter(prefix +  "FROM", from);
           AddParameter(prefix +  "TO", to);
            DataTable list = new DataTable();
            try
            {
                list = executeSelectSP();
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();


            if (sErr != "") ErrorLog.SetLog(sErr);
            return list;
        }
        #endregion Method

    }
}
