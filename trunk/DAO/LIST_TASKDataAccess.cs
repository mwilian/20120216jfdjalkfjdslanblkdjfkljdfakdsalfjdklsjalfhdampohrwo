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
    public class LIST_TASKDataAccess : Connection
    {
		#region Local Variable
        private string _strSPInsertName = "dbo.[procLIST_TASK_add]";
        private string _strSPUpdateName = "dbo.[procLIST_TASK_update]";
        private string _strSPDeleteName = "dbo.[procLIST_TASK_delete]";
        private string _strSPGetName = "dbo.[procLIST_TASK_get]";
        private string _strSPGetAllName = "dbo.[procLIST_TASK_getall]";
		private string _strSPGetPages = "dbo.[procLIST_TASK_getpaged]";
		private string _strSPIsExist = "dbo.[procLIST_TASK_isexist]";
        private string _strTableName = "[LIST_TASK]";
		private string _strSPGetTransferOutName = "dbo.[procLIST_TASK_gettransferout]";
		string _strSPGetCountName = "procLIST_TASK_getcount";
        string _strSPGetByIndexName = "procLIST_TASK_getindex";
		#endregion Local Variable
		
		#region Method
        public LIST_TASKInfo Get(
        String DTB,
        String Code,
		ref string sErr)
        {
			LIST_TASKInfo objEntr = new LIST_TASKInfo();
			connect();
			InitSPCommand(_strSPGetName);              
            AddParameter(LIST_TASKInfo.Field.DTB.ToString(), DTB);
            AddParameter(LIST_TASKInfo.Field.Code.ToString(), Code);
            
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
                objEntr = (LIST_TASKInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            LIST_TASKInfo result = new LIST_TASKInfo();
            result.DTB = (dt.Rows[i][LIST_TASKInfo.Field.DTB.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.DTB.ToString()]));
            result.Code = (dt.Rows[i][LIST_TASKInfo.Field.Code.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.Code.ToString()]));
            result.Description = (dt.Rows[i][LIST_TASKInfo.Field.Description.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.Description.ToString()]));
            result.Lookup = (dt.Rows[i][LIST_TASKInfo.Field.Lookup.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.Lookup.ToString()]));
            result.AttQD_ID = (dt.Rows[i][LIST_TASKInfo.Field.AttQD_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.AttQD_ID.ToString()]));
            result.AttTmp = (dt.Rows[i][LIST_TASKInfo.Field.AttTmp.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.AttTmp.ToString()]));
            result.ValidRange = (dt.Rows[i][LIST_TASKInfo.Field.ValidRange.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.ValidRange.ToString()]));
            result.CntQD_ID = (dt.Rows[i][LIST_TASKInfo.Field.CntQD_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.CntQD_ID.ToString()]));
            result.CntTmp = (dt.Rows[i][LIST_TASKInfo.Field.CntTmp.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.CntTmp.ToString()]));
            result.Emails = (dt.Rows[i][LIST_TASKInfo.Field.Emails.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.Emails.ToString()]));
            result.Server = (dt.Rows[i][LIST_TASKInfo.Field.Server.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.Server.ToString()]));
            result.Protocol = (dt.Rows[i][LIST_TASKInfo.Field.Protocol.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.Protocol.ToString()]));
            result.Port = (dt.Rows[i][LIST_TASKInfo.Field.Port.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.Port.ToString()]));
            result.UserID = (dt.Rows[i][LIST_TASKInfo.Field.UserID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.UserID.ToString()]));
            result.Password = (dt.Rows[i][LIST_TASKInfo.Field.Password.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.Password.ToString()]));
            result.Type = (dt.Rows[i][LIST_TASKInfo.Field.Type.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.Type.ToString()]));
            result.IsUse = (dt.Rows[i][LIST_TASKInfo.Field.IsUse.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TASKInfo.Field.IsUse.ToString()]));
           
            return result;
        }

        public DataTable GetAll(
        String DTB,
        ref string sErr)
        {
            connect();
            InitSPCommand(_strSPGetAllName);
			AddParameter(LIST_TASKInfo.Field.DTB.ToString(), DTB);
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
            InitSPCommand(_strSPGetByIndexName);
			AddParameter("INX", pos);
			AddParameter(LIST_TASKInfo.Field.DTB.ToString(), DTB);
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
            InitSPCommand(_strSPGetCountName);
            AddParameter(LIST_TASKInfo.Field.DTB.ToString(), DTB);
          
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
        public Int32 Add(LIST_TASKInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(LIST_TASKInfo.Field.DTB.ToString(), objEntr.DTB);
            AddParameter(LIST_TASKInfo.Field.Code.ToString(), objEntr.Code);
            AddParameter(LIST_TASKInfo.Field.Description.ToString(), objEntr.Description);
            AddParameter(LIST_TASKInfo.Field.Lookup.ToString(), objEntr.Lookup);
            AddParameter(LIST_TASKInfo.Field.AttQD_ID.ToString(), objEntr.AttQD_ID);
            AddParameter(LIST_TASKInfo.Field.AttTmp.ToString(), objEntr.AttTmp);
            AddParameter(LIST_TASKInfo.Field.ValidRange.ToString(), objEntr.ValidRange);
            AddParameter(LIST_TASKInfo.Field.CntQD_ID.ToString(), objEntr.CntQD_ID);
            AddParameter(LIST_TASKInfo.Field.CntTmp.ToString(), objEntr.CntTmp);
            AddParameter(LIST_TASKInfo.Field.Emails.ToString(), objEntr.Emails);
            AddParameter(LIST_TASKInfo.Field.Server.ToString(), objEntr.Server);
            AddParameter(LIST_TASKInfo.Field.Protocol.ToString(), objEntr.Protocol);
            AddParameter(LIST_TASKInfo.Field.Port.ToString(), objEntr.Port);
            AddParameter(LIST_TASKInfo.Field.UserID.ToString(), objEntr.UserID);
            AddParameter(LIST_TASKInfo.Field.Password.ToString(), objEntr.Password);
            AddParameter(LIST_TASKInfo.Field.Type.ToString(), objEntr.Type);
            AddParameter(LIST_TASKInfo.Field.IsUse.ToString(), objEntr.IsUse);
          
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

        public string Update(LIST_TASKInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter(LIST_TASKInfo.Field.DTB.ToString(), objEntr.DTB);
            AddParameter(LIST_TASKInfo.Field.Code.ToString(), objEntr.Code);
            AddParameter(LIST_TASKInfo.Field.Description.ToString(), objEntr.Description);
            AddParameter(LIST_TASKInfo.Field.Lookup.ToString(), objEntr.Lookup);
            AddParameter(LIST_TASKInfo.Field.AttQD_ID.ToString(), objEntr.AttQD_ID);
            AddParameter(LIST_TASKInfo.Field.AttTmp.ToString(), objEntr.AttTmp);
            AddParameter(LIST_TASKInfo.Field.ValidRange.ToString(), objEntr.ValidRange);
            AddParameter(LIST_TASKInfo.Field.CntQD_ID.ToString(), objEntr.CntQD_ID);
            AddParameter(LIST_TASKInfo.Field.CntTmp.ToString(), objEntr.CntTmp);
            AddParameter(LIST_TASKInfo.Field.Emails.ToString(), objEntr.Emails);
            AddParameter(LIST_TASKInfo.Field.Server.ToString(), objEntr.Server);
            AddParameter(LIST_TASKInfo.Field.Protocol.ToString(), objEntr.Protocol);
            AddParameter(LIST_TASKInfo.Field.Port.ToString(), objEntr.Port);
            AddParameter(LIST_TASKInfo.Field.UserID.ToString(), objEntr.UserID);
            AddParameter(LIST_TASKInfo.Field.Password.ToString(), objEntr.Password);
            AddParameter(LIST_TASKInfo.Field.Type.ToString(), objEntr.Type);
            AddParameter(LIST_TASKInfo.Field.IsUse.ToString(), objEntr.IsUse);
               
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
        String Code
		)
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter(LIST_TASKInfo.Field.DTB.ToString(), DTB);
            AddParameter(LIST_TASKInfo.Field.Code.ToString(), Code);
              
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
		
		public DataTableCollection Get_Page(LIST_TASKInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
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
        String DTB,
        String Code
		)
        {
            connect();
            InitSPCommand(_strSPIsExist);
            AddParameter(LIST_TASKInfo.Field.DTB.ToString(), DTB);
            AddParameter(LIST_TASKInfo.Field.Code.ToString(), Code);
              
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
		
		private string CreateWhereClause(LIST_TASKInfo obj)
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
