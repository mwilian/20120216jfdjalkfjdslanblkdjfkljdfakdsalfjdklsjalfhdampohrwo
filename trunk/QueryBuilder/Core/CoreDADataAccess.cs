using System;
using System.Collections.Generic;
using System.Text;
using DTO;
using System.Data;
using System.Data.SqlClient;
using QueryBuilder;

namespace DAO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
    public class CoreDADataAccess : CoreConnection
    {
		#region Local Variable
        private string _strSPInsertName = "dbo.[procLIST_DA_add]";
        private string _strSPUpdateName = "dbo.[procLIST_DA_update]";
        private string _strSPDeleteName = "dbo.[procLIST_DA_delete]";
        private string _strSPGetName = "dbo.[procLIST_DA_get]";
        private string _strSPGetAllName = "dbo.[procLIST_DA_getall]";
		private string _strSPGetPages = "dbo.[procLIST_DA_getpaged]";
		private string _strSPIsExist = "dbo.[procLIST_DA_isexist]";
        private string _strTableName = "[procLIST_DA]";
		private string _strSPGetTransferOutName = "dbo.[procLIST_DA_gettransferout]";
        private string _strSPGetPermissionName = "LIST_DAGPermission";
        string _strSPGetPermissionByRoleName = "LIST_DAGPermissionByRole";
		#endregion Local Variable
		
		#region Method
        public CoreDAInfo Get(
        String DAG_ID,
		ref string sErr)
        {
			CoreDAInfo objEntr = new CoreDAInfo();
			connect();
			InitSPCommand(_strSPGetName);              
            AddParameter(CoreDAInfo.Field.DAG_ID.ToString(), DAG_ID);
            
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
                objEntr = (CoreDAInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            CoreDAInfo result = new CoreDAInfo();
            result.DAG_ID = (dt.Rows[i][CoreDAInfo.Field.DAG_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][CoreDAInfo.Field.DAG_ID.ToString()]));
            result.NAME = (dt.Rows[i][CoreDAInfo.Field.NAME.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][CoreDAInfo.Field.NAME.ToString()]));
            result.EI = (dt.Rows[i][CoreDAInfo.Field.EI.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][CoreDAInfo.Field.EI.ToString()]));
           
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


            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return list;
        }
		/// <summary>
        /// Return 1: Table is exist Identity Field
        /// Return 0: Table is not exist Identity Field
        /// Return -1: Erro
        /// </summary>
        /// <param name="tableName"></param>
        public Int32 Add(CoreDAInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(CoreDAInfo.Field.DAG_ID.ToString(), objEntr.DAG_ID);
            AddParameter(CoreDAInfo.Field.NAME.ToString(), objEntr.NAME);
            AddParameter(CoreDAInfo.Field.EI.ToString(), objEntr.EI);
          
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
            if (sErr != "") CoreErrorLog.SetLog(sErr);
			
            return ret;
        }

        public string Update(CoreDAInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter(CoreDAInfo.Field.DAG_ID.ToString(), objEntr.DAG_ID);
            AddParameter(CoreDAInfo.Field.NAME.ToString(), objEntr.NAME);
            AddParameter(CoreDAInfo.Field.EI.ToString(), objEntr.EI);
               
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

        public string Delete(
        String DAG_ID
		)
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter(CoreDAInfo.Field.DAG_ID.ToString(), DAG_ID);
              
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
		
		public DataTableCollection Get_Page(CoreDAInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
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
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return dtList;
        }
        
        public Boolean IsExist(
        String DAG_ID
		)
        {
            connect();
            InitSPCommand(_strSPIsExist);
            AddParameter(CoreDAInfo.Field.DAG_ID.ToString(), DAG_ID);
              
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
            if (sErr != "") CoreErrorLog.SetLog(sErr);
            if(list.Rows.Count==1)
				return true;
            return false;
        }
		
		private string CreateWhereClause(CoreDAInfo obj)
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
            //    if (sErr != "") CoreErrorLog.SetLog(sErr);
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


            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return list;
        }
		#endregion Method


        public DataTable GetPermission(string user, ref string sErr)
        {
            connect();
            InitSPCommand(_strSPGetPermissionName);
            AddParameter("USER_ID", user);           
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


            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return list;
        }
        public DataTable GetPermissionByRole(string role, ref string sErr)
        {

            connect();
            InitSPCommand(_strSPGetPermissionByRoleName);
            AddParameter("ROLE_ID", role);
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


            if (sErr != "") CoreErrorLog.SetLog(sErr);
            return list;
        }
    }
}
