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
    public class LIST_GADGETDataAccess : Connection
    {
		#region Local Variable
        private string _strSPInsertName = "procLIST_GADGET_add";
        private string _strSPUpdateName = "procLIST_GADGET_update";
        private string _strSPDeleteName = "procLIST_GADGET_delete";
        private string _strSPGetName = "procLIST_GADGET_get";
        private string _strSPGetAllName = "procLIST_GADGET_getall";
		private string _strSPGetPages = "procLIST_GADGET_getpaged";
		private string _strSPIsExist = "procLIST_GADGET_isexist";
        private string _strTableName = "LIST_GADGET";
		private string _strSPGetTransferOutName = "procLIST_GADGET_gettransferout";
        string prefix = "param";
		#endregion Local Variable
		
		#region Method
        public LIST_GADGETInfo Get(
        Int32 ID,
		ref string sErr)
        {
			LIST_GADGETInfo objEntr = new LIST_GADGETInfo();
			connect();
			InitSPCommand(_strSPGetName);              
            AddParameter(prefix + LIST_GADGETInfo.Field.ID.ToString(), ID);
            
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
                objEntr = (LIST_GADGETInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            LIST_GADGETInfo result = new LIST_GADGETInfo();
            result.ID = (dt.Rows[i][LIST_GADGETInfo.Field.ID.ToString()] == DBNull.Value ? 0 : Convert.ToInt32(dt.Rows[i][LIST_GADGETInfo.Field.ID.ToString()]));
            result.Description = (dt.Rows[i][LIST_GADGETInfo.Field.Description.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_GADGETInfo.Field.Description.ToString()]));
            result.QDCode = (dt.Rows[i][LIST_GADGETInfo.Field.QDCode.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_GADGETInfo.Field.QDCode.ToString()]));
            result.ReportTmp = (dt.Rows[i][LIST_GADGETInfo.Field.ReportTmp.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_GADGETInfo.Field.ReportTmp.ToString()]));
            result.AutoUpdate = (dt.Rows[i][LIST_GADGETInfo.Field.AutoUpdate.ToString()] == DBNull.Value ? 0 : Convert.ToInt32(dt.Rows[i][LIST_GADGETInfo.Field.AutoUpdate.ToString()]));
            result.IsScroll = (dt.Rows[i][LIST_GADGETInfo.Field.IsScroll.ToString()] == DBNull.Value ? true : Convert.ToBoolean(dt.Rows[i][LIST_GADGETInfo.Field.IsScroll.ToString()]));
            result.Action = (dt.Rows[i][LIST_GADGETInfo.Field.Action.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_GADGETInfo.Field.Action.ToString()]));
            result.Argument = (dt.Rows[i][LIST_GADGETInfo.Field.Argument.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_GADGETInfo.Field.Argument.ToString()]));
            result.Image = (dt.Rows[i][LIST_GADGETInfo.Field.Image.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_GADGETInfo.Field.Image.ToString()]));
           
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
            InitSPCommand(_strSPGetAllName);
			AddParameter(prefix + "INX", pos);
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
            InitSPCommand(_strSPInsertName);
          
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
        public Int32 Add(LIST_GADGETInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(prefix + LIST_GADGETInfo.Field.Description.ToString(), objEntr.Description);
            AddParameter(prefix + LIST_GADGETInfo.Field.QDCode.ToString(), objEntr.QDCode);
            AddParameter(prefix + LIST_GADGETInfo.Field.ReportTmp.ToString(), objEntr.ReportTmp);
            AddParameter(prefix + LIST_GADGETInfo.Field.AutoUpdate.ToString(), objEntr.AutoUpdate);
            AddParameter(prefix + LIST_GADGETInfo.Field.IsScroll.ToString(), objEntr.IsScroll);
            AddParameter(prefix + LIST_GADGETInfo.Field.Action.ToString(), objEntr.Action);
            AddParameter(prefix + LIST_GADGETInfo.Field.Argument.ToString(), objEntr.Argument);
            AddParameter(prefix + LIST_GADGETInfo.Field.Image.ToString(), objEntr.Image);
          
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

        public string Update(LIST_GADGETInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter(prefix + LIST_GADGETInfo.Field.ID.ToString(), objEntr.ID);
            AddParameter(prefix + LIST_GADGETInfo.Field.Description.ToString(), objEntr.Description);
            AddParameter(prefix + LIST_GADGETInfo.Field.QDCode.ToString(), objEntr.QDCode);
            AddParameter(prefix + LIST_GADGETInfo.Field.ReportTmp.ToString(), objEntr.ReportTmp);
            AddParameter(prefix + LIST_GADGETInfo.Field.AutoUpdate.ToString(), objEntr.AutoUpdate);
            AddParameter(prefix + LIST_GADGETInfo.Field.IsScroll.ToString(), objEntr.IsScroll);
            AddParameter(prefix + LIST_GADGETInfo.Field.Action.ToString(), objEntr.Action);
            AddParameter(prefix + LIST_GADGETInfo.Field.Argument.ToString(), objEntr.Argument);
            AddParameter(prefix + LIST_GADGETInfo.Field.Image.ToString(), objEntr.Image);
               
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
        Int32 ID
		)
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter(prefix + LIST_GADGETInfo.Field.ID.ToString(), ID);
              
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
		
		public DataTableCollection Get_Page(LIST_GADGETInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
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
        Int32 ID
		)
        {
            connect();
            InitSPCommand(_strSPIsExist);
            AddParameter(prefix + LIST_GADGETInfo.Field.ID.ToString(), ID);
              
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
		
		private string CreateWhereClause(LIST_GADGETInfo obj)
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
     
    }
}
