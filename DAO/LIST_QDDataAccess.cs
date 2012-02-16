using System;
using System.Collections.Generic;
using System.Text;
using DTO;
using System.Data;
using System.Data.SqlClient;

namespace DAO
{
    public class LIST_QDDataAccess : Connection
    {
        #region Local Variable
        private string _strSPInsertName = "dbo.[procLIST_QD_add]";
        private string _strSPUpdateName = "dbo.[procLIST_QD_update]";
        private string _strSPDeleteName = "dbo.[procLIST_QD_delete]";
        private string _strSPGetName = "dbo.[procLIST_QD_get]";
        private string _strSPGetAllName = "dbo.[procLIST_QD_getall]";
        private string _strSPGetAllByUserName = "[dbo].[procLIST_QDs_getuser]";
        private string _strSPGetAllByCateName = "dbo.[procLIST_QD_getcate]";
        private string _strSPGetPages = "dbo.[procLIST_QD_getpaged]";
        private string _strSPIsExist = "dbo.[procLIST_QD_isexist]";
        private string _strTableName = "LIST_QD";
        #endregion Local Variable

        #region Method
        public LIST_QDInfo Get_LIST_QD(
            String DTB,
            String QD_ID
        , ref string sErr)
        {
            LIST_QDInfo objEntr = new LIST_QDInfo();
            DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPGetName);
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

            if (list.Rows.Count > 0)
                objEntr = (LIST_QDInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            LIST_QDInfo result = new LIST_QDInfo();
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

        public DataTable GetAll_LIST_QD(String DTB, ref string sErr)
        {
            LIST_QDInfo objEntr = new LIST_QDInfo(); DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPGetAllName);
                AddParameter("DTB", DTB);


                list = executeSelectSP(command);
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();


            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return list;
        }
        /// <summary>
        /// Return 1: Table is exist Identity Field
        /// Return 0: Table is not exist Identity Field
        /// Return -1: Erro
        /// </summary>
        /// <param name="tableName"></param>
        public Int32 Add_LIST_QD(LIST_QDInfo objEntr, ref string sErr)
        {
            int ret = -1; try
            {
                connect();
                InitSPCommand(_strSPInsertName);
                AddParameter("DTB", objEntr.DTB);
                AddParameter("QD_ID", objEntr.QD_ID);
                AddParameter("DESCRIPTN", objEntr.DESCRIPTN);
                AddParameter("OWNER", objEntr.OWNER);
                AddParameter("SHARED", objEntr.SHARED);
                AddParameter("LAYOUT", objEntr.LAYOUT);
                AddParameter("ANAL_Q0", objEntr.ANAL_Q0);
                AddParameter("ANAL_Q9", objEntr.ANAL_Q9);
                AddParameter("ANAL_Q8", objEntr.ANAL_Q8);
                AddParameter("ANAL_Q7", objEntr.ANAL_Q7);
                AddParameter("ANAL_Q6", objEntr.ANAL_Q6);
                AddParameter("ANAL_Q5", objEntr.ANAL_Q5);
                AddParameter("ANAL_Q4", objEntr.ANAL_Q4);
                AddParameter("ANAL_Q3", objEntr.ANAL_Q3);
                AddParameter("ANAL_Q2", objEntr.ANAL_Q2);
                AddParameter("ANAL_Q1", objEntr.ANAL_Q1);
                AddParameter("SQL_TEXT", objEntr.SQL_TEXT);
                AddParameter("HEADER_TEXT", objEntr.HEADER_TEXT);
                AddParameter("FOOTER_TEXT", objEntr.FOOTER_TEXT);


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

        public string Update_LIST_QD(LIST_QDInfo objEntr)
        {
            string sErr = "";
            try
            {
                connect();
                InitSPCommand(_strSPUpdateName);

                AddParameter("DTB", objEntr.DTB);
                AddParameter("QD_ID", objEntr.QD_ID);
                AddParameter("DESCRIPTN", objEntr.DESCRIPTN);
                AddParameter("OWNER", objEntr.OWNER);
                AddParameter("SHARED", objEntr.SHARED);
                AddParameter("LAYOUT", objEntr.LAYOUT);
                AddParameter("ANAL_Q0", objEntr.ANAL_Q0);
                AddParameter("ANAL_Q9", objEntr.ANAL_Q9);
                AddParameter("ANAL_Q8", objEntr.ANAL_Q8);
                AddParameter("ANAL_Q7", objEntr.ANAL_Q7);
                AddParameter("ANAL_Q6", objEntr.ANAL_Q6);
                AddParameter("ANAL_Q5", objEntr.ANAL_Q5);
                AddParameter("ANAL_Q4", objEntr.ANAL_Q4);
                AddParameter("ANAL_Q3", objEntr.ANAL_Q3);
                AddParameter("ANAL_Q2", objEntr.ANAL_Q2);
                AddParameter("ANAL_Q1", objEntr.ANAL_Q1);
                AddParameter("SQL_TEXT", objEntr.SQL_TEXT);
                AddParameter("HEADER_TEXT", objEntr.HEADER_TEXT);
                AddParameter("FOOTER_TEXT", objEntr.FOOTER_TEXT);


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

        public string Delete_LIST_QD(
            String DTB,
            String QD_ID
        )
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter("DTB", DTB);
            AddParameter("QD_ID", QD_ID);

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

        public DataTableCollection Get_Page(LIST_QDInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            string whereClause = CreateWhereClause(obj);
            DataTableCollection dtList = null; try
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
            if (sErr != "") ErrorLog.SetLog(sErr);
            return dtList;
        }

        public Boolean IsExist_LIST_QD(
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

                AddParameter("DTB", DTB);
                AddParameter("QD_ID", QD_ID);


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

        private string CreateWhereClause(LIST_QDInfo obj)
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

        public DataTable getALL_LIST_QD_By_ARJ(String DTB, String kind, ref String sErr)
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


        public DataTable GetAll_LIST_QD_ByGroup(string dtb, string _strCategory, ref string sErr)
        {
            return Search("DTB ='" + dtb + "' AND ANAL_Q1 ", "'" + _strCategory + "'", "=", ref sErr);
        }

        public DataTable GetAll_LIST_QD_ByCate(string dtb, string _strCategory, ref string sErr)
        {
            LIST_QDInfo objEntr = new LIST_QDInfo();
            DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPGetAllByCateName);
                AddParameter("DTB", dtb);
                AddParameter("CATEGORY", _strCategory);


                list = executeSelectSP(command);
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();


            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return list;
        }

        public DataTable ToTransferInStruct()
        {
            string sErr = "";
            string query = "SELECT   top 0  LIST_QDD.QDD_ID, LIST_QDD.CODE, LIST_QDD.DESCRIPTN AS QDD_DESCRIPTN, LIST_QDD.F_TYPE, LIST_QDD.SORTING, LIST_QDD.AGREGATE,                       LIST_QDD.EXPRESSION, LIST_QDD.FILTER_FROM, LIST_QDD.FILTER_TO, LIST_QDD.IS_FILTER, LIST_QDD_FILTER.IS_NOT, LIST_QDD_FILTER.OPERATOR,                       LIST_QD.* FROM         LIST_QD INNER JOIN                      LIST_QDD ON LIST_QD.DTB = LIST_QDD.DTB AND LIST_QD.QD_ID = LIST_QDD.QD_ID LEFT OUTER JOIN                      LIST_QDD_FILTER ON LIST_QDD.DTB = LIST_QDD_FILTER.DTB AND LIST_QDD.QD_ID = LIST_QDD_FILTER.QD_ID AND                       LIST_QDD.QDD_ID = LIST_QDD_FILTER.QDD_ID";
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

        public DataTable GetTransferOut_LIST_QD(string dtb, ref string sErr)
        {
            string query = "SELECT     LIST_QDD.QDD_ID, LIST_QDD.CODE, LIST_QDD.DESCRIPTN AS QDD_DESCRIPTN, LIST_QDD.F_TYPE, LIST_QDD.SORTING, LIST_QDD.AGREGATE,                       LIST_QDD.EXPRESSION, LIST_QDD.FILTER_FROM, LIST_QDD.FILTER_TO, LIST_QDD.IS_FILTER, LIST_QDD_FILTER.IS_NOT, LIST_QDD_FILTER.OPERATOR,                       LIST_QD.* FROM         LIST_QD INNER JOIN                      LIST_QDD ON LIST_QD.DTB = LIST_QDD.DTB AND LIST_QD.QD_ID = LIST_QDD.QD_ID LEFT OUTER JOIN                      LIST_QDD_FILTER ON LIST_QDD.DTB = LIST_QDD_FILTER.DTB AND LIST_QDD.QD_ID = LIST_QDD_FILTER.QD_ID AND                       LIST_QDD.QDD_ID = LIST_QDD_FILTER.QDD_ID";
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

        public DataTable GetTransferOut_LIST_QD(string DTB, string QD_CODE, ref string sErr)
        {
            string query = "SELECT     LIST_QDD.QDD_ID, LIST_QDD.CODE, LIST_QDD.DESCRIPTN AS QDD_DESCRIPTN, LIST_QDD.F_TYPE, LIST_QDD.SORTING, LIST_QDD.AGREGATE,                       LIST_QDD.EXPRESSION, LIST_QDD.FILTER_FROM, LIST_QDD.FILTER_TO, LIST_QDD.IS_FILTER, LIST_QDD_FILTER.IS_NOT, LIST_QDD_FILTER.OPERATOR,                       LIST_QD.* FROM         LIST_QD INNER JOIN                      LIST_QDD ON LIST_QD.DTB = LIST_QDD.DTB AND LIST_QD.QD_ID = LIST_QDD.QD_ID AND LIST_QDD.QD_ID = '" + QD_CODE + "' LEFT OUTER JOIN                      LIST_QDD_FILTER ON LIST_QDD.DTB = LIST_QDD_FILTER.DTB AND LIST_QDD.QD_ID = LIST_QDD_FILTER.QD_ID AND                       LIST_QDD.QDD_ID = LIST_QDD_FILTER.QDD_ID";
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

        public DataTable GetAll_LIST_QD_USER(string database, string user, ref string sErr)
        {
            LIST_QDInfo objEntr = new LIST_QDInfo();
            DataTable list = new DataTable();
            try
            {
                connect();
                InitSPCommand(_strSPGetAllByUserName);
                AddParameter("DTB", database);
                AddParameter("USER_ID", user);


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
