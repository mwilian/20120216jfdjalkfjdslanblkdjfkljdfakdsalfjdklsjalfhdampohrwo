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
    public class LIST_EMAILDataAccess : Connection
    {
		#region Local Variable
        private string _strSPInsertName = "dbo.[procLIST_EMAIL_add]";
        private string _strSPUpdateName = "dbo.[procLIST_EMAIL_update]";
        private string _strSPDeleteName = "dbo.[procLIST_EMAIL_delete]";
        private string _strSPGetName = "dbo.[procLIST_EMAIL_get]";
        private string _strSPGetAllName = "dbo.[procLIST_EMAIL_getall]";
		private string _strSPGetPages = "dbo.[procLIST_EMAIL_getpaged]";
		private string _strSPIsExist = "dbo.[procLIST_EMAIL_isexist]";
        private string _strTableName = "[LIST_EMAIL]";
		private string _strSPGetTransferOutName = "dbo.[procLIST_EMAIL_gettransferout]";
		string _strSPGetCountName = "procLIST_EMAIL_getcount";
        string _strSPGetByIndexName = "procLIST_EMAIL_getindex";
		#endregion Local Variable
		
		#region Method
        public LIST_EMAILInfo Get(
        String Mail,
		ref string sErr)
        {
			LIST_EMAILInfo objEntr = new LIST_EMAILInfo();
			connect();
			InitSPCommand(_strSPGetName);              
            AddParameter(LIST_EMAILInfo.Field.Mail.ToString(), Mail);
            
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
                objEntr = (LIST_EMAILInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            LIST_EMAILInfo result = new LIST_EMAILInfo();
            result.Mail = (dt.Rows[i][LIST_EMAILInfo.Field.Mail.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_EMAILInfo.Field.Mail.ToString()]));
            result.Name = (dt.Rows[i][LIST_EMAILInfo.Field.Name.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_EMAILInfo.Field.Name.ToString()]));
            result.Lookup = (dt.Rows[i][LIST_EMAILInfo.Field.Lookup.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_EMAILInfo.Field.Lookup.ToString()]));
           
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
		public DataTable GetByPos(
        int pos, ref string sErr)
        {
            connect();
            InitSPCommand(_strSPGetByIndexName);
			AddParameter("INX", pos);
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
        ref string sErr)
        {
			int ret = -1;
            connect();
            InitSPCommand(_strSPGetCountName);
          
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
        public Int32 Add(LIST_EMAILInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(LIST_EMAILInfo.Field.Mail.ToString(), objEntr.Mail);
            AddParameter(LIST_EMAILInfo.Field.Name.ToString(), objEntr.Name);
            AddParameter(LIST_EMAILInfo.Field.Lookup.ToString(), objEntr.Lookup);
          
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

        public string Update(LIST_EMAILInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter(LIST_EMAILInfo.Field.Mail.ToString(), objEntr.Mail);
            AddParameter(LIST_EMAILInfo.Field.Name.ToString(), objEntr.Name);
            AddParameter(LIST_EMAILInfo.Field.Lookup.ToString(), objEntr.Lookup);
               
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
        String Mail
		)
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter(LIST_EMAILInfo.Field.Mail.ToString(), Mail);
              
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
		
		public DataTableCollection Get_Page(LIST_EMAILInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
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
        String Mail
		)
        {
            connect();
            InitSPCommand(_strSPIsExist);
            AddParameter(LIST_EMAILInfo.Field.Mail.ToString(), Mail);
              
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
		
		private string CreateWhereClause(LIST_EMAILInfo obj)
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
