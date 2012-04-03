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
    public class LIST_TEMPLATEDataAccess : Connection
    {
		#region Local Variable
        private string _strSPInsertName = "procLIST_TEMPLATE_add";
        private string _strSPUpdateName = "procLIST_TEMPLATE_update";
        private string _strSPDeleteName = "procLIST_TEMPLATE_delete";
        private string _strSPGetName = "procLIST_TEMPLATE_get";
        private string _strSPGetAllName = "procLIST_TEMPLATE_getall";
		private string _strSPGetPages = "procLIST_TEMPLATE_getpaged";
		private string _strSPIsExist = "procLIST_TEMPLATE_isexist";
        private string _strTableName = "LIST_TEMPLATE";
		private string _strSPGetTransferOutName = "procLIST_TEMPLATE_gettransferout";
		string _strSPGetCountName = "procLIST_TEMPLATE_getcount";
        string _strSPGetByIndexName = "procLIST_TEMPLATE_getindex";
        string prefix = "param";
		#endregion Local Variable
		
		#region Method
        public LIST_TEMPLATEInfo Get(
        String DTB,
        String Code,
		ref string sErr)
        {
			LIST_TEMPLATEInfo objEntr = new LIST_TEMPLATEInfo();
			connect();
			InitSPCommand(_strSPGetName);              
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.DTB.ToString(), DTB);
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.Code.ToString(), Code);
            
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
                objEntr = (LIST_TEMPLATEInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            LIST_TEMPLATEInfo result = new LIST_TEMPLATEInfo();
            result.DTB = (dt.Rows[i][LIST_TEMPLATEInfo.Field.DTB.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TEMPLATEInfo.Field.DTB.ToString()]));
            result.Code = (dt.Rows[i][LIST_TEMPLATEInfo.Field.Code.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][LIST_TEMPLATEInfo.Field.Code.ToString()]));
            result.Data = (dt.Rows[i][LIST_TEMPLATEInfo.Field.Data.ToString()] == DBNull.Value ? null : (Byte[])(dt.Rows[i][LIST_TEMPLATEInfo.Field.Data.ToString()]));
            result.Length = (dt.Rows[i][LIST_TEMPLATEInfo.Field.Length.ToString()] == DBNull.Value ? 0 : (int)(dt.Rows[i][LIST_TEMPLATEInfo.Field.Length.ToString()]));
            return result;
        }

        public DataTable GetAll(
        String DTB,
        ref string sErr)
        {
            connect();
            InitSPCommand(_strSPGetAllName);
			AddParameter(prefix  + LIST_TEMPLATEInfo.Field.DTB.ToString(), DTB);
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
			AddParameter(prefix  + "INX", pos);
			AddParameter(prefix  + LIST_TEMPLATEInfo.Field.DTB.ToString(), DTB);
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
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.DTB.ToString(), DTB);
          
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
        public Int32 Add(LIST_TEMPLATEInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.DTB.ToString(), objEntr.DTB);
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.Code.ToString(), objEntr.Code);
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.Data.ToString(), objEntr.Data);
            AddParameter(prefix + LIST_TEMPLATEInfo.Field.Length.ToString(), objEntr.Length);
          
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

        public string Update(LIST_TEMPLATEInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.DTB.ToString(), objEntr.DTB);
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.Code.ToString(), objEntr.Code);
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.Data.ToString(), objEntr.Data);
            AddParameter(prefix + LIST_TEMPLATEInfo.Field.Length.ToString(), objEntr.Length);
               
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
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.DTB.ToString(), DTB);
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.Code.ToString(), Code);
              
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
		
		public DataTableCollection Get_Page(LIST_TEMPLATEInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
			string whereClause = CreateWhereClause(obj);
            DataTableCollection dtList = null;
            connect();
            InitSPCommand(_strSPGetPages); 
          
            AddParameter(prefix  + "WhereClause", whereClause);
            AddParameter(prefix  + "OrderBy", orderBy);
            AddParameter(prefix  + "PageIndex", pageIndex);
            AddParameter(prefix  + "PageSize", pageSize);
            
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
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.DTB.ToString(), DTB);
            AddParameter(prefix  + LIST_TEMPLATEInfo.Field.Code.ToString(), Code);
              
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
		
		private string CreateWhereClause(LIST_TEMPLATEInfo obj)
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
			AddParameter(prefix  + "DB", dtb);
			AddParameter(prefix  + "FROM", from);
			AddParameter(prefix  + "TO", to);
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
