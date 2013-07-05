using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace QueryBuilder
{
    public class CoreQDDDataAccess : CoreConnection
    {
        #region Local Variable
        private string _strSPInsertName = "dbo.[procLIST_QDD_add]";
        private string _strSPUpdateName = "dbo.[procLIST_QDD_update]";
        private string _strSPDeleteName = "dbo.[procLIST_QDD_delete]";
        private string _strSPGetName = "dbo.[procLIST_QDD_get]";
        private string _strSPGetAllName = "dbo.[procLIST_QDD_getall]";
        private string _strSPGetAllName_By_QD_ID = "dbo.[sp_procLIST_QDD_Select_By_QD_ID]";
        private string _strSPGetPages = "dbo.[procLIST_QDD_getpaged]";
        private string _strSPIsExist = "dbo.[procLIST_QDD_isexist]";
        private string _strTableName = "CoreQDD";
        #endregion Local Variable

        #region Method
        public CoreQDDInfo Get_CoreQDD(
            String DTB,
            String QD_ID,
            Int32 QDD_ID
        , ref string sErr)
        {
            CoreQDDInfo objEntr = new CoreQDDInfo();
            DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPGetName);
                AddParameter("DTB", DTB);
                AddParameter("QD_ID", QD_ID);
                AddParameter("QDD_ID", QDD_ID);


                list = executeSelectSP(command);
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();

            if (list.Rows.Count > 0)
                objEntr = (CoreQDDInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return objEntr;
        }

        public DataTable GetALL_CoreQDD_By_QD_ID(
            String DTB,
            String QD_ID

        , ref string sErr)
        {
            CoreQDDInfo objEntr = new CoreQDDInfo();
            DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPGetAllName_By_QD_ID);
                AddParameter("DTB", DTB);
                AddParameter("QD_ID", QD_ID);
                list = executeSelectSP(command);
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            finally
            {
                disconnect();
            }


            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return list;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            CoreQDDInfo result = new CoreQDDInfo();
            result.DTB = (dt.Rows[i]["DTB"] == DBNull.Value ? "" : (String)dt.Rows[i]["DTB"]);
            result.QD_ID = (dt.Rows[i]["QD_ID"] == DBNull.Value ? "" : (String)dt.Rows[i]["QD_ID"]);
            result.QDD_ID = (dt.Rows[i]["QDD_ID"] == DBNull.Value ? -1 : (Int32)dt.Rows[i]["QDD_ID"]);
            result.CODE = (dt.Rows[i]["CODE"] == DBNull.Value ? "" : (String)dt.Rows[i]["CODE"]);
            result.DESCRIPTN = (dt.Rows[i]["DESCRIPTN"] == DBNull.Value ? "" : (String)dt.Rows[i]["DESCRIPTN"]);
            result.F_TYPE = (dt.Rows[i]["F_TYPE"] == DBNull.Value ? "" : (String)dt.Rows[i]["F_TYPE"]);
            result.SORTING = (dt.Rows[i]["SORTING"] == DBNull.Value ? "" : (String)dt.Rows[i]["SORTING"]);
            result.AGREGATE = (dt.Rows[i]["AGREGATE"] == DBNull.Value ? "" : (String)dt.Rows[i]["AGREGATE"]);
            result.EXPRESSION = (dt.Rows[i]["EXPRESSION"] == DBNull.Value ? "" : (String)dt.Rows[i]["EXPRESSION"]);
            result.FILTER_FROM = (dt.Rows[i]["FILTER_FROM"] == DBNull.Value ? "" : (String)dt.Rows[i]["FILTER_FROM"]);
            result.FILTER_TO = (dt.Rows[i]["FILTER_TO"] == DBNull.Value ? "" : (String)dt.Rows[i]["FILTER_TO"]);
            result.IS_FILTER = (dt.Rows[i]["IS_FILTER"] == DBNull.Value ? true : (Boolean)dt.Rows[i]["IS_FILTER"]);

            return result;
        }

        public DataTable GetAll_CoreQDD(ref string sErr)
        {
            string query = "exec " + _strSPGetAllName;
            DataTable list = new DataTable();
            try
            {
                connect();

                list = executeSelectQuery(query);
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return list;
        }
        /// <summary>
        /// Return 1: Table is exist Identity Field
        /// Return 0: Table is not exist Identity Field
        /// Return -1: Erro
        /// </summary>
        /// <param name="tableName"></param>
        public Int32 Add_CoreQDD(CoreQDDInfo objEntr, ref string sErr)
        {
            int ret = -1;
            try
            {
                connect();
                InitSPCommand(_strSPInsertName);
                AddParameter("DTB", objEntr.DTB);
                AddParameter("QD_ID", objEntr.QD_ID);
                AddParameter("QDD_ID", objEntr.QDD_ID);
                AddParameter("CODE", objEntr.CODE);
                AddParameter("DESCRIPTN", objEntr.DESCRIPTN);
                AddParameter("F_TYPE", objEntr.F_TYPE);
                AddParameter("SORTING", objEntr.SORTING);
                AddParameter("AGREGATE", objEntr.AGREGATE);
                AddParameter("EXPRESSION", objEntr.EXPRESSION);
                AddParameter("FILTER_FROM", objEntr.FILTER_FROM);
                AddParameter("FILTER_TO", objEntr.FILTER_TO);
                AddParameter("IS_FILTER", objEntr.IS_FILTER);


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
            if (sErr != "") CoreErrorLog.SetLog(sErr);

            return ret;
        }

        public string Update_CoreQDD(CoreQDDInfo objEntr)
        {

            string sErr = "";
            try
            {
                connect();
                InitSPCommand(_strSPUpdateName);

                AddParameter("DTB", objEntr.DTB);
                AddParameter("QD_ID", objEntr.QD_ID);
                AddParameter("QDD_ID", objEntr.QDD_ID);
                AddParameter("CODE", objEntr.CODE);
                AddParameter("DESCRIPTN", objEntr.DESCRIPTN);
                AddParameter("F_TYPE", objEntr.F_TYPE);
                AddParameter("SORTING", objEntr.SORTING);
                AddParameter("AGREGATE", objEntr.AGREGATE);
                AddParameter("EXPRESSION", objEntr.EXPRESSION);
                AddParameter("FILTER_FROM", objEntr.FILTER_FROM);
                AddParameter("FILTER_TO", objEntr.FILTER_TO);
                AddParameter("IS_FILTER", objEntr.IS_FILTER);


                excuteSPNonQuery();
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return sErr;
        }

        public string Delete_CoreQDD(
            String DTB,
            String QD_ID,
            Int32 QDD_ID
        )
        {
            string sErr = "";
            try
            {
                connect();
                InitSPCommand(_strSPDeleteName);
                AddParameter("DTB", DTB);
                AddParameter("QD_ID", QD_ID);
                AddParameter("QDD_ID", QDD_ID);


                excuteSPNonQuery();
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return sErr;
        }

        public DataTableCollection Get_Page(CoreQDDInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            string whereClause = CreateWhereClause(obj);
            DataTableCollection dtList = null;
            try
            {
                connect();
                InitSPCommand(_strSPGetPages);

                AddParameter("WhereClause", whereClause);
                AddParameter("OrderBy", orderBy);
                AddParameter("PageIndex", pageIndex);
                AddParameter("PageSize", pageSize);


                dtList = executeCollectSelectSP();
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return dtList;
        }

        public Boolean IsExist_CoreQDD(
            String DTB,
            String QD_ID,
            Int32 QDD_ID
        )
        {
            string sErr = "";
            DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPIsExist);

                AddParameter("DTB", DTB);
                AddParameter("QD_ID", QD_ID);
                AddParameter("QDD_ID", QDD_ID);


                list = executeSelectSP();
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            if (list.Rows.Count == 1)
                return true;
            return false;
        }

        private string CreateWhereClause(CoreQDDInfo obj)
        {
            String result = "";

            return result;
        }

        public DataTable Search(string columnName, string columnValue, string condition, ref string sErr)
        {
            string query = "select * from " + _strTableName + " where " + columnName + " " + condition + " " + columnValue;
            DataTable list = new DataTable();
            try
            {
                connect();

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


        #endregion Method


        public void Delete_CoreQDD_By_QD_ID(string qdID, string dtb, ref string sErr)
        {
            string query = "Delete from " + _strTableName + " where QD_ID='" + qdID + "' and DTB='" + dtb + "'";
            try
            {
                connect();

                executeNonQuery(query);
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
        }
    }
}
