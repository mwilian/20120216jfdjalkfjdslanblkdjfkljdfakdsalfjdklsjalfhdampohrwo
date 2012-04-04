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
    public class POPDataAccess : Connection
    {
		#region Local Variable
        private string _strSPInsertName = "procPOP_add";
        private string _strSPUpdateName = "procPOP_update";
        private string _strSPDeleteName = "procPOP_delete";
        private string _strSPGetName = "procPOP_get";
        private string _strSPGetAllName = "procPOP_getall";
		private string _strSPGetPages = "procPOP_getpaged";
		private string _strSPIsExist = "procPOP_isexist";
        private string _strTableName = "SSINSTAL";
		private string _strSPGetTransferOutName = "procPOP_gettransferout";
        string prefix = "param";
		#endregion Local Variable
		
		#region Method
        public POPInfo Get(
        String ROLE_ID,
        String DB,
		ref string sErr)
        {
			POPInfo objEntr = new POPInfo();
			connect();
			InitSPCommand(_strSPGetName);              
            AddParameter(prefix + POPInfo.Field.ROLE_ID.ToString(), ROLE_ID);
            AddParameter(prefix + POPInfo.Field.DB.ToString(), DB);
            
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
                objEntr = (POPInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            POPInfo result = new POPInfo();
            result.ROLE_ID = (dt.Rows[i][POPInfo.Field.ROLE_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POPInfo.Field.ROLE_ID.ToString()]));
            result.DB = (dt.Rows[i][POPInfo.Field.DB.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POPInfo.Field.DB.ToString()]));
            result.DEFAULT_VALUE = (dt.Rows[i][POPInfo.Field.DEFAULT_VALUE.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POPInfo.Field.DEFAULT_VALUE.ToString()]));
            result.PERMISSION = (dt.Rows[i][POPInfo.Field.PERMISSION.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POPInfo.Field.PERMISSION.ToString()]));
           
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
        public Int32 Add(POPInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(prefix + POPInfo.Field.ROLE_ID.ToString(), objEntr.ROLE_ID);
            AddParameter(prefix + POPInfo.Field.DB.ToString(), objEntr.DB);
            AddParameter(prefix + POPInfo.Field.DEFAULT_VALUE.ToString(), objEntr.DEFAULT_VALUE);
            AddParameter(prefix + POPInfo.Field.PERMISSION.ToString(), objEntr.PERMISSION);
          
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

        public string Update(POPInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter(prefix + POPInfo.Field.ROLE_ID.ToString(), objEntr.ROLE_ID);
            AddParameter(prefix + POPInfo.Field.DB.ToString(), objEntr.DB);
            AddParameter(prefix + POPInfo.Field.DEFAULT_VALUE.ToString(), objEntr.DEFAULT_VALUE);
            AddParameter(prefix + POPInfo.Field.PERMISSION.ToString(), objEntr.PERMISSION);
               
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
        String ROLE_ID,
        String DB
		)
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter(prefix + POPInfo.Field.ROLE_ID.ToString(), ROLE_ID);
            AddParameter(prefix + POPInfo.Field.DB.ToString(), DB);
              
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
		
		public DataTableCollection Get_Page(POPInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
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
        String ROLE_ID,
        String DB
		)
        {
            connect();
            InitSPCommand(_strSPIsExist);
            AddParameter(prefix + POPInfo.Field.ROLE_ID.ToString(), ROLE_ID);
            AddParameter(prefix + POPInfo.Field.DB.ToString(), DB);
              
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
		
		private string CreateWhereClause(POPInfo obj)
        {
            String result = "";

            return result;
        }
        
        public DataTable Search(string columnName, string columnValue, string condition, ref string sErr)
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
