using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace QueryBuilder
{
    public class CoreQDDataAccess : CoreConnection
    {
        #region Local Variable
        private string _strSPInsertName = "procLIST_QD_add";
        private string _strSPUpdateName = "procLIST_QD_update";
        private string _strSPDeleteName = "procLIST_QD_delete";
        private string _strSPGetName = "procLIST_QD_get";
        private string _strSPGetAllName = "procLIST_QD_getall";
        private string _strSPGetAllByUserName = "dbo].[procLIST_QDs_getuser";
        private string _strSPGetAllByCateName = "procLIST_QD_getcate";
        private string _strSPGetPages = "procLIST_QD_getpaged";
        private string _strSPIsExist = "procLIST_QD_isexist";
        private string _strTableName = "CoreQD";
        string prefix = "param";
        #endregion Local Variable

        #region Method
        public CoreQDInfo Get_CoreQD(
            String DTB,
            String QD_ID
        , ref string sErr)
        {
            CoreQDInfo objEntr = new CoreQDInfo();
            DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPGetName);
                AddParameter(prefix + "DTB", DTB);
                AddParameter(prefix + "QD_ID", QD_ID);



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

            if (list.Rows.Count > 0)
                objEntr = (CoreQDInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            CoreQDInfo result = new CoreQDInfo();
            result.DTB = (dt.Rows[i]["DTB"] == DBNull.Value ? "" : (String)dt.Rows[i]["DTB"]);
            result.QD_ID = (dt.Rows[i]["QD_ID"] == DBNull.Value ? "" : (String)dt.Rows[i]["QD_ID"]);
            result.DESCRIPTN = (dt.Rows[i]["DESCRIPTN"] == DBNull.Value ? "" : (String)dt.Rows[i]["DESCRIPTN"]);
            result.OWNER = (dt.Rows[i]["OWNER"] == DBNull.Value ? "" : (String)dt.Rows[i]["OWNER"]);
            result.SHARED = (dt.Rows[i]["SHARED"] == DBNull.Value ? true : (Boolean)dt.Rows[i]["SHARED"]);
            result.LAYOUT = (dt.Rows[i]["LAYOUT"] == DBNull.Value ? "" : (String)dt.Rows[i]["LAYOUT"]);
            result.ANAL_Q0 = (dt.Rows[i]["ANAL_Q0"] == DBNull.Value ? "" : (String)dt.Rows[i]["ANAL_Q0"]);
            result.ANAL_Q9 = (dt.Rows[i]["ANAL_Q9"] == DBNull.Value ? "" : (String)dt.Rows[i]["ANAL_Q9"]);
            result.ANAL_Q8 = (dt.Rows[i]["ANAL_Q8"] == DBNull.Value ? "" : (String)dt.Rows[i]["ANAL_Q8"]);
            result.ANAL_Q7 = (dt.Rows[i]["ANAL_Q7"] == DBNull.Value ? "" : (String)dt.Rows[i]["ANAL_Q7"]);
            result.ANAL_Q6 = (dt.Rows[i]["ANAL_Q6"] == DBNull.Value ? "" : (String)dt.Rows[i]["ANAL_Q6"]);
            result.ANAL_Q5 = (dt.Rows[i]["ANAL_Q5"] == DBNull.Value ? "" : (String)dt.Rows[i]["ANAL_Q5"]);
            result.ANAL_Q4 = (dt.Rows[i]["ANAL_Q4"] == DBNull.Value ? "" : (String)dt.Rows[i]["ANAL_Q4"]);
            result.ANAL_Q3 = (dt.Rows[i]["ANAL_Q3"] == DBNull.Value ? "" : (String)dt.Rows[i]["ANAL_Q3"]);
            result.ANAL_Q2 = (dt.Rows[i]["ANAL_Q2"] == DBNull.Value ? "" : (String)dt.Rows[i]["ANAL_Q2"]);
            result.ANAL_Q1 = (dt.Rows[i]["ANAL_Q1"] == DBNull.Value ? "" : (String)dt.Rows[i]["ANAL_Q1"]);
            result.SQL_TEXT = (dt.Rows[i]["SQL_TEXT"] == DBNull.Value ? "" : (String)dt.Rows[i]["SQL_TEXT"]);
            result.HEADER_TEXT = (dt.Rows[i]["HEADER_TEXT"] == DBNull.Value ? "" : (String)dt.Rows[i]["HEADER_TEXT"]);
            result.FOOTER_TEXT = (dt.Rows[i]["FOOTER_TEXT"] == DBNull.Value ? "" : (String)dt.Rows[i]["FOOTER_TEXT"]);

            return result;
        }

        public DataTable GetAll_CoreQD(String DTB, ref string sErr)
        {
            CoreQDInfo objEntr = new CoreQDInfo(); DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPGetAllName);
                AddParameter(prefix + "DTB", DTB);


                list = executeSelectSP(command);
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
        public Int32 Add_CoreQD(CoreQDInfo objEntr, ref string sErr)
        {
            int ret = -1; try
            {
                connect();
                InitSPCommand(_strSPInsertName);
                AddParameter(prefix + "DTB", objEntr.DTB);
                AddParameter(prefix + "QD_ID", objEntr.QD_ID);
                AddParameter(prefix + "DESCRIPTN", objEntr.DESCRIPTN);
                AddParameter(prefix + "OWNER", objEntr.OWNER);
                AddParameter(prefix + "SHARED", objEntr.SHARED);
                AddParameter(prefix + "LAYOUT", objEntr.LAYOUT);
                AddParameter(prefix + "ANAL_Q0", objEntr.ANAL_Q0);
                AddParameter(prefix + "ANAL_Q9", objEntr.ANAL_Q9);
                AddParameter(prefix + "ANAL_Q8", objEntr.ANAL_Q8);
                AddParameter(prefix + "ANAL_Q7", objEntr.ANAL_Q7);
                AddParameter(prefix + "ANAL_Q6", objEntr.ANAL_Q6);
                AddParameter(prefix + "ANAL_Q5", objEntr.ANAL_Q5);
                AddParameter(prefix + "ANAL_Q4", objEntr.ANAL_Q4);
                AddParameter(prefix + "ANAL_Q3", objEntr.ANAL_Q3);
                AddParameter(prefix + "ANAL_Q2", objEntr.ANAL_Q2);
                AddParameter(prefix + "ANAL_Q1", objEntr.ANAL_Q1);
                AddParameter(prefix + "SQL_TEXT", objEntr.SQL_TEXT);
                AddParameter(prefix + "HEADER_TEXT", objEntr.HEADER_TEXT);
                AddParameter(prefix + "FOOTER_TEXT", objEntr.FOOTER_TEXT);


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

        public string Update_CoreQD(CoreQDInfo objEntr)
        {
            string sErr = "";
            try
            {
                connect();
                InitSPCommand(_strSPUpdateName);

                AddParameter(prefix + "DTB", objEntr.DTB);
                AddParameter(prefix + "QD_ID", objEntr.QD_ID);
                AddParameter(prefix + "DESCRIPTN", objEntr.DESCRIPTN);
                AddParameter(prefix + "OWNER", objEntr.OWNER);
                AddParameter(prefix + "SHARED", objEntr.SHARED);
                AddParameter(prefix + "LAYOUT", objEntr.LAYOUT);
                AddParameter(prefix + "ANAL_Q0", objEntr.ANAL_Q0);
                AddParameter(prefix + "ANAL_Q9", objEntr.ANAL_Q9);
                AddParameter(prefix + "ANAL_Q8", objEntr.ANAL_Q8);
                AddParameter(prefix + "ANAL_Q7", objEntr.ANAL_Q7);
                AddParameter(prefix + "ANAL_Q6", objEntr.ANAL_Q6);
                AddParameter(prefix + "ANAL_Q5", objEntr.ANAL_Q5);
                AddParameter(prefix + "ANAL_Q4", objEntr.ANAL_Q4);
                AddParameter(prefix + "ANAL_Q3", objEntr.ANAL_Q3);
                AddParameter(prefix + "ANAL_Q2", objEntr.ANAL_Q2);
                AddParameter(prefix + "ANAL_Q1", objEntr.ANAL_Q1);
                AddParameter(prefix + "SQL_TEXT", objEntr.SQL_TEXT);
                AddParameter(prefix + "HEADER_TEXT", objEntr.HEADER_TEXT);
                AddParameter(prefix + "FOOTER_TEXT", objEntr.FOOTER_TEXT);


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

        public string Delete_CoreQD(
            String DTB,
            String QD_ID
        )
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter(prefix + "DTB", DTB);
            AddParameter(prefix + "QD_ID", QD_ID);

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
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return sErr;
        }

        public DataTableCollection Get_Page(CoreQDInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            string whereClause = CreateWhereClause(obj);
            DataTableCollection dtList = null; try
            {
                connect();
                InitSPCommand(_strSPGetPages);

                AddParameter(prefix + "WhereClause", whereClause);
                AddParameter(prefix + "OrderBy", orderBy);
                AddParameter(prefix + "PageIndex", pageIndex);
                AddParameter(prefix + "PageSize", pageSize);


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

        public Boolean IsExist_CoreQD(
            String DTB,
            String QD_ID
        )
        {
            string sErr = "";
            DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPIsExist);

                AddParameter(prefix + "DTB", DTB);
                AddParameter(prefix + "QD_ID", QD_ID);


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

        private string CreateWhereClause(CoreQDInfo obj)
        {
            String result = "";

            return result;
        }

        public DataTable Search(string columnName, string columnValue, string condition, ref string sErr)
        {
            string query = "select * from " + _strTableName + " where " + columnName + " " + condition + " " + columnValue;
            DataTable list = new DataTable(); try
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

        public String getMaxARJ(String kind, ref String sErr)
        {
            String kq = "000"; Object list = new Object();
            connect();
            try
            {
                string query = "SELECT MAX(QD_ID) FROM " + _strTableName + " WHERE [OWNER] LIKE '_____SYSTM' AND QD_ID LIKE '" + kind + "%'";

                list = executeScalar(query);
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (list != DBNull.Value)
            {
                String num = (String)list;
                int number = Convert.ToInt32(num.Substring(4, 6)) + 1;
                if (number > 9)
                    kq = "0" + number.ToString();
                else if (number > 99)
                    kq = number.ToString();
                else
                    kq = "00" + number.ToString();
            }
            return kq;
        }

        public DataTable getALL_CoreQD_By_ARJ(String DTB, String kind, ref String sErr)
        {
            string query = "SELECT * FROM " + _strTableName + " WHERE [OWNER] LIKE '_____SYSTM' AND QD_ID LIKE '" + kind + "%'";
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

            return list;
        }

        #endregion Method


        public DataTable GetAll_CoreQD_ByGroup(string dtb, string _strCategory, ref string sErr)
        {
            return Search("DTB ='" + dtb + "' AND ANAL_Q1 ", "'" + _strCategory + "'", "=", ref sErr);
        }

        public DataTable GetAll_CoreQD_ByCate(string dtb, string _strCategory, ref string sErr)
        {
            CoreQDInfo objEntr = new CoreQDInfo();
            DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPGetAllByCateName);
                AddParameter(prefix + "DTB", dtb);
                AddParameter(prefix + "CATEGORY", _strCategory);


                list = executeSelectSP(command);
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

        public DataTable ToTransferInStruct()
        {
            string sErr = "";
            string query = "SELECT   top 0  CoreQDD.QDD_ID, CoreQDD.CODE, CoreQDD.DESCRIPTN AS QDD_DESCRIPTN, CoreQDD.F_TYPE, CoreQDD.SORTING, CoreQDD.AGREGATE, " +
                            "CoreQDD.EXPRESSION, CoreQDD.FILTER_FROM, CoreQDD.FILTER_TO, CoreQDD.IS_FILTER, CoreQD.* " +
                            "FROM         CoreQD INNER JOIN " +
                            "CoreQDD ON CoreQD.DTB = CoreQDD.DTB AND CoreQD.QD_ID = CoreQDD.QD_ID";
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

        public DataTable GetTransferOut_CoreQD(string dtb, ref string sErr)
        {
            string query = "SELECT    CoreQDD.QDD_ID, CoreQDD.CODE, CoreQDD.DESCRIPTN AS QDD_DESCRIPTN, CoreQDD.F_TYPE, CoreQDD.SORTING, CoreQDD.AGREGATE, " +
                            "CoreQDD.EXPRESSION, CoreQDD.FILTER_FROM, CoreQDD.FILTER_TO, CoreQDD.IS_FILTER, CoreQD.* " +
                            "FROM         CoreQD INNER JOIN " +
                            "CoreQDD ON CoreQD.DTB = CoreQDD.DTB AND CoreQD.QD_ID = CoreQDD.QD_ID";
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

        public DataTable GetTransferOut_CoreQD(string DTB, string QD_CODE, ref string sErr)
        {
            string query = "SELECT    CoreQDD.QDD_ID, CoreQDD.CODE, CoreQDD.DESCRIPTN AS QDD_DESCRIPTN, CoreQDD.F_TYPE, CoreQDD.SORTING, CoreQDD.AGREGATE, " +
                            "CoreQDD.EXPRESSION, CoreQDD.FILTER_FROM, CoreQDD.FILTER_TO, CoreQDD.IS_FILTER, CoreQD.* " +
                            "FROM         CoreQD INNER JOIN " +
                            "CoreQDD ON CoreQD.DTB = CoreQDD.DTB AND CoreQD.QD_ID = CoreQDD.QD_ID AND CoreQDD.QD_ID='" + QD_CODE + "'";
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

        public DataTable GetAll_CoreQD_USER(string database, string user, ref string sErr)
        {
            CoreQDInfo objEntr = new CoreQDInfo();
            DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPGetAllByUserName);
                AddParameter(prefix + "DTB", database);
                AddParameter(prefix + "USER_ID", user);


                list = executeSelectSP(command);
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            return list;
        }
    }
}
