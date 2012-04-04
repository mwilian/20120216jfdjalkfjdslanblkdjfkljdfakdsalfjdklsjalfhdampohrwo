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
    public class LIST_DADataAccess : Connection
    {
        #region Local Variable
        private string _strSPInsertName = "procLIST_DA_add";
        private string _strSPUpdateName = "procLIST_DA_update";
        private string _strSPDeleteName = "LIST_DA_delete";
        private string _strSPGetName = "procLIST_DA_get";
        private string _strSPGetAllName = "procLIST_DA_getall";
        private string _strSPGetPages = "procLIST_DA_getpaged";
        private string _strSPIsExist = "procLIST_DA_isexist";
        private string _strTableName = "LIST_DA";
        private string _strSPGetTransferOutName = "procLIST_DA_gettransferout";
        private string _strSPGetPermissionName = "LIST_DAGPermission";
        string _strSPGetPermissionByRoleName = "LIST_DAGPermissionByRole";
        string prefix = "param";
        #endregion Local Variable

        #region Method
        public LIST_DAInfo Get(
        String DAG_ID,
        ref string sErr)
        {
            LIST_DAInfo objEntr = new LIST_DAInfo();
            connect();
            InitSPCommand(_strSPGetName);
           AddParameter(prefix +  LIST_DAInfo.Field.DAG_ID.ToString(), DAG_ID);

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
                objEntr = (LIST_DAInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            LIST_DAInfo result = new LIST_DAInfo();
            result.DAG_ID = (dt.Rows[i][LIST_DAInfo.Field.DAG_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_DAInfo.Field.DAG_ID.ToString()]));
            result.NAME = (dt.Rows[i][LIST_DAInfo.Field.NAME.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_DAInfo.Field.NAME.ToString()]));
            result.EI = (dt.Rows[i][LIST_DAInfo.Field.EI.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_DAInfo.Field.EI.ToString()]));

            return result;
        }

        public DataTable GetAll(
        ref string sErr)
        {
            connect();
            InitSPCommand(_strSPGetAllName);
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
        public Int32 Add(LIST_DAInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
           AddParameter(prefix +  LIST_DAInfo.Field.DAG_ID.ToString(), objEntr.DAG_ID);
           AddParameter(prefix +  LIST_DAInfo.Field.NAME.ToString(), objEntr.NAME);
           AddParameter(prefix +  LIST_DAInfo.Field.EI.ToString(), objEntr.EI);

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

        public string Update(LIST_DAInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
           AddParameter(prefix +  LIST_DAInfo.Field.DAG_ID.ToString(), objEntr.DAG_ID);
           AddParameter(prefix +  LIST_DAInfo.Field.NAME.ToString(), objEntr.NAME);
           AddParameter(prefix +  LIST_DAInfo.Field.EI.ToString(), objEntr.EI);

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
        String DAG_ID
        )
        {
            connect();
            InitSPCommand(_strSPDeleteName);
           AddParameter(prefix +  LIST_DAInfo.Field.DAG_ID.ToString(), DAG_ID);

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

        public DataTableCollection Get_Page(LIST_DAInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
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
        String DAG_ID
        )
        {
            connect();
            InitSPCommand(_strSPIsExist);
           AddParameter(prefix +  LIST_DAInfo.Field.DAG_ID.ToString(), DAG_ID);

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

        private string CreateWhereClause(LIST_DAInfo obj)
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


        public DataTable GetPermission(string user, ref string sErr)
        {
            connect();
            InitSPCommand(_strSPGetPermissionName);
           AddParameter(prefix +  "USER_ID", user);
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
        public DataTable GetPermissionByRole(string role, ref string sErr)
        {

            connect();
            InitSPCommand(_strSPGetPermissionByRoleName);
           AddParameter(prefix +  "ROLE_ID", role);
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
    }
}
