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
    public class DBADataAccess : Connection
    {
        #region Local Variable
        private string _strSPInsertName = "dbo.[procDBA_add]";
        private string _strSPUpdateName = "dbo.[procDBA_update]";
        private string _strSPDeleteName = "dbo.[procDBA_delete]";
        private string _strSPGetName = "dbo.[procDBA_get]";
        private string _strSPGetAllName = "dbo.[procDBA_getall]";
        private string _strSPGetPages = "dbo.[procDBA_getpaged]";
        private string _strSPIsExist = "dbo.[procDBA_isexist]";
        private string _strTableName = "[SSINSTAL]";
        private string _strSPGetTransferOutName = "dbo.[procDBA_gettransferout]";
        #endregion Local Variable

        #region Method
        public DBAInfo Get(
        String DB,
        ref string sErr)
        {
            DBAInfo objEntr = new DBAInfo();
            connect();
            InitSPCommand(_strSPGetName);
            AddParameter(DBAInfo.Field.DB.ToString(), DB);

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
                objEntr = (DBAInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            DBAInfo result = new DBAInfo();
            result.DB = (dt.Rows[i][DBAInfo.Field.DB.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.DB.ToString()]));
            result.DB1 = (dt.Rows[i][DBAInfo.Field.DB1.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.DB1.ToString()]));
            result.DB2 = (dt.Rows[i][DBAInfo.Field.DB2.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.DB2.ToString()]));
            result.DESCRIPTION = (dt.Rows[i][DBAInfo.Field.DESCRIPTION.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.DESCRIPTION.ToString()]));
            result.DATE_FORMAT = (dt.Rows[i][DBAInfo.Field.DATE_FORMAT.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.DATE_FORMAT.ToString()]));
            result.DECIMAL_PLACES_SUNACCOUNT = (dt.Rows[i][DBAInfo.Field.DECIMAL_PLACES_SUNACCOUNT.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.DECIMAL_PLACES_SUNACCOUNT.ToString()]));
            result.DECIMAL_SEPERATOR = (dt.Rows[i][DBAInfo.Field.DECIMAL_SEPERATOR.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.DECIMAL_SEPERATOR.ToString()]));
            result.THOUSAND_SEPERATOR = (dt.Rows[i][DBAInfo.Field.THOUSAND_SEPERATOR.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.THOUSAND_SEPERATOR.ToString()]));
            result.PRIMARY_BUDGET = (dt.Rows[i][DBAInfo.Field.PRIMARY_BUDGET.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.PRIMARY_BUDGET.ToString()]));
            result.DATA_ACCESS_GROUP = (dt.Rows[i][DBAInfo.Field.DATA_ACCESS_GROUP.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.DATA_ACCESS_GROUP.ToString()]));
            result.DECIMAL_PLACES_SUNBUSINESS = (dt.Rows[i][DBAInfo.Field.DECIMAL_PLACES_SUNBUSINESS.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.DECIMAL_PLACES_SUNBUSINESS.ToString()]));
            result.REPORT_TEMPLATE_DRIVER = (dt.Rows[i][DBAInfo.Field.REPORT_TEMPLATE_DRIVER.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.REPORT_TEMPLATE_DRIVER.ToString()]));
            if (dt.Columns.Contains(DBAInfo.Field.PARAM_1.ToString()))
                result.PARAM_1 = (dt.Rows[i][DBAInfo.Field.PARAM_1.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.PARAM_1.ToString()]));
            if (dt.Columns.Contains(DBAInfo.Field.PARAM_2.ToString()))
                result.PARAM_2 = (dt.Rows[i][DBAInfo.Field.PARAM_2.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.PARAM_2.ToString()]));
            if (dt.Columns.Contains(DBAInfo.Field.PARAM_3.ToString()))
                result.PARAM_3 = (dt.Rows[i][DBAInfo.Field.PARAM_3.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.PARAM_3.ToString()]));
            if (dt.Columns.Contains(DBAInfo.Field.PARAM_4.ToString()))
                result.PARAM_4 = (dt.Rows[i][DBAInfo.Field.PARAM_4.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.PARAM_4.ToString()]));
            if (dt.Columns.Contains(DBAInfo.Field.PARAM_5.ToString()))
                result.PARAM_5 = (dt.Rows[i][DBAInfo.Field.PARAM_5.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.PARAM_5.ToString()]));
            if (dt.Columns.Contains(DBAInfo.Field.PARAM_6.ToString()))
                result.PARAM_6 = (dt.Rows[i][DBAInfo.Field.PARAM_6.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][DBAInfo.Field.PARAM_6.ToString()]));

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
        public Int32 Add(DBAInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(DBAInfo.Field.DB.ToString(), objEntr.DB);
            AddParameter(DBAInfo.Field.DB1.ToString(), objEntr.DB1);
            AddParameter(DBAInfo.Field.DB2.ToString(), objEntr.DB2);
            AddParameter(DBAInfo.Field.DESCRIPTION.ToString(), objEntr.DESCRIPTION);
            AddParameter(DBAInfo.Field.DATE_FORMAT.ToString(), objEntr.DATE_FORMAT);
            AddParameter(DBAInfo.Field.DECIMAL_PLACES_SUNACCOUNT.ToString(), objEntr.DECIMAL_PLACES_SUNACCOUNT);
            AddParameter(DBAInfo.Field.DECIMAL_SEPERATOR.ToString(), objEntr.DECIMAL_SEPERATOR);
            AddParameter(DBAInfo.Field.THOUSAND_SEPERATOR.ToString(), objEntr.THOUSAND_SEPERATOR);
            AddParameter(DBAInfo.Field.PRIMARY_BUDGET.ToString(), objEntr.PRIMARY_BUDGET);
            AddParameter(DBAInfo.Field.DATA_ACCESS_GROUP.ToString(), objEntr.DATA_ACCESS_GROUP);
            AddParameter(DBAInfo.Field.DECIMAL_PLACES_SUNBUSINESS.ToString(), objEntr.DECIMAL_PLACES_SUNBUSINESS);
            AddParameter(DBAInfo.Field.REPORT_TEMPLATE_DRIVER.ToString(), objEntr.REPORT_TEMPLATE_DRIVER);
            AddParameter(DBAInfo.Field.PARAM_1.ToString(), objEntr.PARAM_1);
            AddParameter(DBAInfo.Field.PARAM_2.ToString(), objEntr.PARAM_2);
            AddParameter(DBAInfo.Field.PARAM_3.ToString(), objEntr.PARAM_3);
            AddParameter(DBAInfo.Field.PARAM_4.ToString(), objEntr.PARAM_4);
            AddParameter(DBAInfo.Field.PARAM_5.ToString(), objEntr.PARAM_5);
            AddParameter(DBAInfo.Field.PARAM_6.ToString(), objEntr.PARAM_6);

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

        public string Update(DBAInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter(DBAInfo.Field.DB.ToString(), objEntr.DB);
            AddParameter(DBAInfo.Field.DB1.ToString(), objEntr.DB1);
            AddParameter(DBAInfo.Field.DB2.ToString(), objEntr.DB2);
            AddParameter(DBAInfo.Field.DESCRIPTION.ToString(), objEntr.DESCRIPTION);
            AddParameter(DBAInfo.Field.DATE_FORMAT.ToString(), objEntr.DATE_FORMAT);
            AddParameter(DBAInfo.Field.DECIMAL_PLACES_SUNACCOUNT.ToString(), objEntr.DECIMAL_PLACES_SUNACCOUNT);
            AddParameter(DBAInfo.Field.DECIMAL_SEPERATOR.ToString(), objEntr.DECIMAL_SEPERATOR);
            AddParameter(DBAInfo.Field.THOUSAND_SEPERATOR.ToString(), objEntr.THOUSAND_SEPERATOR);
            AddParameter(DBAInfo.Field.PRIMARY_BUDGET.ToString(), objEntr.PRIMARY_BUDGET);
            AddParameter(DBAInfo.Field.DATA_ACCESS_GROUP.ToString(), objEntr.DATA_ACCESS_GROUP);
            AddParameter(DBAInfo.Field.DECIMAL_PLACES_SUNBUSINESS.ToString(), objEntr.DECIMAL_PLACES_SUNBUSINESS);
            AddParameter(DBAInfo.Field.REPORT_TEMPLATE_DRIVER.ToString(), objEntr.REPORT_TEMPLATE_DRIVER);
            AddParameter(DBAInfo.Field.PARAM_1.ToString(), objEntr.PARAM_1);
            AddParameter(DBAInfo.Field.PARAM_2.ToString(), objEntr.PARAM_2);
            AddParameter(DBAInfo.Field.PARAM_3.ToString(), objEntr.PARAM_3);
            AddParameter(DBAInfo.Field.PARAM_4.ToString(), objEntr.PARAM_4);
            AddParameter(DBAInfo.Field.PARAM_5.ToString(), objEntr.PARAM_5);
            AddParameter(DBAInfo.Field.PARAM_6.ToString(), objEntr.PARAM_6);

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
        String DB
        )
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter(DBAInfo.Field.DB.ToString(), DB);

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

        public DataTableCollection Get_Page(DBAInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            string whereClause = CreateWhereClause(obj);
            DataTableCollection dtList = null;
            connect();
            InitSPCommand(_strSPGetPages);

            AddParameter("WhereClause", whereClause);
            AddParameter("OrderBy", orderBy);
            AddParameter("PageIndex", pageIndex);
            AddParameter("PageSize", pageSize);

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
        String DB
        )
        {
            connect();
            InitSPCommand(_strSPIsExist);
            AddParameter(DBAInfo.Field.DB.ToString(), DB);

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

        private string CreateWhereClause(DBAInfo obj)
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
            AddParameter("DB", dtb);
            AddParameter("FROM", from);
            AddParameter("TO", to);
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
