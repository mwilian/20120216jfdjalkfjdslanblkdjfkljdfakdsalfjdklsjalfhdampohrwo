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
    public class PODDataAccess : Connection
    {
		#region Local Variable
        private string _strSPInsertName = "dbo.[procPOD_add]";
        private string _strSPUpdateName = "dbo.[procPOD_update]";
        private string _strSPDeleteName = "dbo.[procPOD_delete]";
        private string _strSPGetName = "dbo.[procPOD_get]";
        private string _strSPGetAllName = "dbo.[procPOD_getall]";
		private string _strSPGetPages = "dbo.[procPOD_getpaged]";
		private string _strSPIsExist = "dbo.[procPOD_isexist]";
        private string _strTableName = "[SSINSTAL]";
		private string _strSPGetTransferOutName = "dbo.[procPOD_gettransferout]";
		#endregion Local Variable
		
		#region Method
        public PODInfo Get(
        String USER_ID,
		ref string sErr)
        {
			PODInfo objEntr = new PODInfo();
			connect();
			InitSPCommand(_strSPGetName);              
            AddParameter(PODInfo.Field.USER_ID.ToString(), USER_ID);
            
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
                objEntr = (PODInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            PODInfo result = new PODInfo();
            result.USER_ID = (dt.Rows[i][PODInfo.Field.USER_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][PODInfo.Field.USER_ID.ToString()]));
            result.TB = (dt.Rows[i][PODInfo.Field.TB.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][PODInfo.Field.TB.ToString()]));
            result.USER_ID1 = (dt.Rows[i][PODInfo.Field.USER_ID1.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][PODInfo.Field.USER_ID1.ToString()]));
            result.USER_NAME = (dt.Rows[i][PODInfo.Field.USER_NAME.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][PODInfo.Field.USER_NAME.ToString()]));
            result.DB_DEFAULT = (dt.Rows[i][PODInfo.Field.DB_DEFAULT.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][PODInfo.Field.DB_DEFAULT.ToString()]));
            result.LANGUAGE = (dt.Rows[i][PODInfo.Field.LANGUAGE.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][PODInfo.Field.LANGUAGE.ToString()]));
            result.ROLE_ID = (dt.Rows[i][PODInfo.Field.ROLE_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][PODInfo.Field.ROLE_ID.ToString()]));
            result.PASS = (dt.Rows[i][PODInfo.Field.PASS.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][PODInfo.Field.PASS.ToString()]));
           
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
        public Int32 Add(PODInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(PODInfo.Field.USER_ID.ToString(), objEntr.USER_ID);
            AddParameter(PODInfo.Field.TB.ToString(), objEntr.TB);
            AddParameter(PODInfo.Field.USER_ID1.ToString(), objEntr.USER_ID1);
            AddParameter(PODInfo.Field.USER_NAME.ToString(), objEntr.USER_NAME);
            AddParameter(PODInfo.Field.DB_DEFAULT.ToString(), objEntr.DB_DEFAULT);
            AddParameter(PODInfo.Field.LANGUAGE.ToString(), objEntr.LANGUAGE);
            AddParameter(PODInfo.Field.ROLE_ID.ToString(), objEntr.ROLE_ID);
            AddParameter(PODInfo.Field.PASS.ToString(), objEntr.PASS);
          
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

        public string Update(PODInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter(PODInfo.Field.USER_ID.ToString(), objEntr.USER_ID);
            AddParameter(PODInfo.Field.TB.ToString(), objEntr.TB);
            AddParameter(PODInfo.Field.USER_ID1.ToString(), objEntr.USER_ID1);
            AddParameter(PODInfo.Field.USER_NAME.ToString(), objEntr.USER_NAME);
            AddParameter(PODInfo.Field.DB_DEFAULT.ToString(), objEntr.DB_DEFAULT);
            AddParameter(PODInfo.Field.LANGUAGE.ToString(), objEntr.LANGUAGE);
            AddParameter(PODInfo.Field.ROLE_ID.ToString(), objEntr.ROLE_ID);
            AddParameter(PODInfo.Field.PASS.ToString(), objEntr.PASS);
               
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
            AddParameter(PODInfo.Field.USER_ID.ToString(), USER_ID);
              
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
		
		public DataTableCollection Get_Page(PODInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
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
            AddParameter(PODInfo.Field.USER_ID.ToString(), USER_ID);
              
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
		
		private string CreateWhereClause(PODInfo obj)
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
