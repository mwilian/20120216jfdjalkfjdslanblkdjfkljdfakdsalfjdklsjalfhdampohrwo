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
    public class POSDataAccess : Connection
    {
		#region Local Variable
        private string _strSPInsertName = "dbo.[procPOS_add]";
        private string _strSPUpdateName = "dbo.[procPOS_update]";
        private string _strSPDeleteName = "dbo.[procPOS_delete]";
        private string _strSPGetName = "dbo.[procPOS_get]";
        private string _strSPGetAllName = "dbo.[procPOS_getall]";
		private string _strSPGetPages = "dbo.[procPOS_getpaged]";
		private string _strSPIsExist = "dbo.[procPOS_isexist]";
        private string _strTableName = "[SSINSTAL]";
		#endregion Local Variable
		
		#region Method
        public POSInfo Get(
        String USER_ID,
		ref string sErr)
        {
			POSInfo objEntr = new POSInfo();
			connect();
			InitSPCommand(_strSPGetName);              
            AddParameter("USER_ID", USER_ID);
            
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
                objEntr = (POSInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            POSInfo result = new POSInfo();
            result.USER_ID = (dt.Rows[i]["USER_ID"] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i]["USER_ID"]));
            result.CURRENT_DB = (dt.Rows[i]["CURRENT_DB"] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i]["CURRENT_DB"]));
            result.CURRENT_ACTIVITY = (dt.Rows[i]["CURRENT_ACTIVITY"] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i]["CURRENT_ACTIVITY"]));
            result.WORK_STATION = (dt.Rows[i]["WORK_STATION"] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i]["WORK_STATION"]));
            result.LOGIN_TIME = (dt.Rows[i]["LOGIN_TIME"] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i]["LOGIN_TIME"]));
           
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
        public Int32 Add(POSInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter("USER_ID", objEntr.USER_ID);
            AddParameter("CURRENT_DB", objEntr.CURRENT_DB);
            AddParameter("CURRENT_ACTIVITY", objEntr.CURRENT_ACTIVITY);
            AddParameter("WORK_STATION", objEntr.WORK_STATION);
            AddParameter("LOGIN_TIME", objEntr.LOGIN_TIME);
          
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

        public string Update(POSInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter("USER_ID", objEntr.USER_ID);
            AddParameter("CURRENT_DB", objEntr.CURRENT_DB);
            AddParameter("CURRENT_ACTIVITY", objEntr.CURRENT_ACTIVITY);
            AddParameter("WORK_STATION", objEntr.WORK_STATION);
            AddParameter("LOGIN_TIME", objEntr.LOGIN_TIME);
               
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
        String USER_ID
		)
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter("USER_ID", USER_ID);
              
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
		
		public DataTableCollection Get_Page(POSInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
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
        String USER_ID
		)
        {
            connect();
            InitSPCommand(_strSPIsExist);
            AddParameter("USER_ID", USER_ID);
              
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
            if(list.Rows.Count>=1)
				return true;
            return false;
        }
		
		private string CreateWhereClause(POSInfo obj)
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
		#endregion Method


        public int GetCount(ref string sErr)
        {
            int result = 0;
            string query = "select Count(*) from " + _strTableName + " where  INS_TB = 'POS'" ;
            DataTable list = new DataTable();
            connect();
            try
            {
                result = Convert.ToInt32(executeScalar(query));
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
            }
            disconnect();
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            //    if (sErr != "") ErrorLog.SetLog(sErr);
            return result;
        }
    }
}
