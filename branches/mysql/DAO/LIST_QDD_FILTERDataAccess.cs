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
    public class LIST_QDD_FILTERDataAccess : Connection
    {
		#region Local Variable
        private string _strSPInsertName = "procLIST_QDD_FILTER_add";
        private string _strSPUpdateName = "procLIST_QDD_FILTER_update";
        private string _strSPDeleteName = "procLIST_QDD_FILTER_delete";
        private string _strSPGetName = "procLIST_QDD_FILTER_get";
        private string _strSPGetAllName = "procLIST_QDD_FILTER_getall";
		private string _strSPGetPages = "procLIST_QDD_FILTER_getpaged";
		private string _strSPIsExist = "procLIST_QDD_FILTER_isexist";
        private string _strTableName = "LIST_QDD_FILTER";
		private string _strSPGetTransferOutName = "procLIST_QDD_FILTER_gettransferout";
        private string _strSPDeleteByQDName = "procLIST_QDD_FILTER_deleteqd";
        string prefix = "param";
		#endregion Local Variable
		
		#region Method
        public LIST_QDD_FILTERInfo Get(
        String DTB,
        String QD_ID,
        Int32 QDD_ID,
		ref string sErr)
        {
			LIST_QDD_FILTERInfo objEntr = new LIST_QDD_FILTERInfo();
			connect();
			InitSPCommand(_strSPGetName);              
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.DTB.ToString(), DTB);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QD_ID.ToString(), QD_ID);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QDD_ID.ToString(), QDD_ID);
            
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
                objEntr = (LIST_QDD_FILTERInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            LIST_QDD_FILTERInfo result = new LIST_QDD_FILTERInfo();
            result.DTB = (dt.Rows[i][LIST_QDD_FILTERInfo.Field.DTB.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_QDD_FILTERInfo.Field.DTB.ToString()]));
            result.QD_ID = (dt.Rows[i][LIST_QDD_FILTERInfo.Field.QD_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_QDD_FILTERInfo.Field.QD_ID.ToString()]));
            result.QDD_ID = (dt.Rows[i][LIST_QDD_FILTERInfo.Field.QDD_ID.ToString()] == DBNull.Value ? 0 : Convert.ToInt32(dt.Rows[i][LIST_QDD_FILTERInfo.Field.QDD_ID.ToString()]));
            result.OPERATOR = (dt.Rows[i][LIST_QDD_FILTERInfo.Field.OPERATOR.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_QDD_FILTERInfo.Field.OPERATOR.ToString()]));
            result.IS_NOT = (dt.Rows[i][LIST_QDD_FILTERInfo.Field.IS_NOT.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_QDD_FILTERInfo.Field.IS_NOT.ToString()]));
           
            return result;
        }

        public DataTable GetAll(
        String DTB,
        ref string sErr)
        {
            connect();
            InitSPCommand(_strSPGetAllName);
			AddParameter(prefix + LIST_QDD_FILTERInfo.Field.DTB.ToString(), DTB);
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
		public DataTable GetByPos(
        String DTB,
        int pos, ref string sErr)
        {
            connect();
            InitSPCommand(_strSPGetAllName);
			AddParameter(prefix + "INX", pos);
			AddParameter(prefix + LIST_QDD_FILTERInfo.Field.DTB.ToString(), DTB);
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
		public int GetCountRecord(
        String DTB,
        ref string sErr)
        {
			int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.DTB.ToString(), DTB);
          
            try
            {
                //command.ExecuteNonQuery();
                object tmp = executeSPScalar();
                if(tmp != null && tmp != DBNull.Value)
					ret = Convert.ToInt32(tmp);
				else 
					ret=0;
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") ErrorLog.SetLog(sErr);
			
            return ret;			
        }
		/// <summary>
        /// Return 1: Table is exist Identity Field
        /// Return 0: Table is not exist Identity Field
        /// Return -1: Erro
        /// </summary>
        /// <param name="tableName"></param>
        public Int32 Add(LIST_QDD_FILTERInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.DTB.ToString(), objEntr.DTB);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QD_ID.ToString(), objEntr.QD_ID);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QDD_ID.ToString(), objEntr.QDD_ID);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.OPERATOR.ToString(), objEntr.OPERATOR);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.IS_NOT.ToString(), objEntr.IS_NOT);
          
            try
            {
                //command.ExecuteNonQuery();
                object tmp = executeSPScalar();
                if(tmp != null && tmp != DBNull.Value)
					ret = Convert.ToInt32(tmp);
				else 
					ret=0;
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            if (sErr != "") ErrorLog.SetLog(sErr);
			
            return ret;
        }

        public string Update(LIST_QDD_FILTERInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.DTB.ToString(), objEntr.DTB);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QD_ID.ToString(), objEntr.QD_ID);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QDD_ID.ToString(), objEntr.QDD_ID);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.OPERATOR.ToString(), objEntr.OPERATOR);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.IS_NOT.ToString(), objEntr.IS_NOT);
               
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
        String DTB,
        String QD_ID,
        Int32 QDD_ID
		)
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.DTB.ToString(), DTB);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QD_ID.ToString(), QD_ID);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QDD_ID.ToString(), QDD_ID);
              
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
		
		public DataTableCollection Get_Page(LIST_QDD_FILTERInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
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
        String DTB,
        String QD_ID,
        Int32 QDD_ID
		)
        {
            connect();
            InitSPCommand(_strSPIsExist);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.DTB.ToString(), DTB);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QD_ID.ToString(), QD_ID);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QDD_ID.ToString(), QDD_ID);
              
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
            if(list.Rows.Count==1)
				return true;
            return false;
        }
		
		private string CreateWhereClause(LIST_QDD_FILTERInfo obj)
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
			AddParameter(prefix + "DB", dtb);
			AddParameter(prefix + "FROM", from);
			AddParameter(prefix + "TO", to);
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


        public object DeleteByQD_ID(string QD_ID, string DTB, ref string sErr)
        {
            connect();
            InitSPCommand(_strSPDeleteByQDName);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.DTB.ToString(), DTB);
            AddParameter(prefix + LIST_QDD_FILTERInfo.Field.QD_ID.ToString(), QD_ID);

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
    }
}
