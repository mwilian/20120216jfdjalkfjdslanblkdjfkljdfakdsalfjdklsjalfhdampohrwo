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
    public class POGDataAccess : Connection
    {
		#region Local Variable
        private string _strSPInsertName = "procPOG_add";
        private string _strSPUpdateName = "procPOG_update";
        private string _strSPDeleteName = "procPOG_delete";
        private string _strSPGetName = "procPOG_get";
        private string _strSPGetAllName = "procPOG_getall";
		private string _strSPGetPages = "procPOG_getpaged";
		private string _strSPIsExist = "procPOG_isexist";
        private string _strTableName = "SSINSTAL";
		private string _strSPGetTransferOutName = "procPOG_gettransferout";
        string prefix = "param";
		#endregion Local Variable
		
		#region Method
        public POGInfo Get(
        String ROLE_ID,
		ref string sErr)
        {
			POGInfo objEntr = new POGInfo();
			connect();
			InitSPCommand(_strSPGetName);              
            AddParameter(prefix + POGInfo.Field.ROLE_ID.ToString(), ROLE_ID);
            
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
                objEntr = (POGInfo)GetDataFromDataRow(list, 0);
            //if (dr != null) list = CBO.FillCollection(dr, ref list);
            if (sErr != "") ErrorLog.SetLog(sErr);
            return objEntr;
        }

        protected override object GetDataFromDataRow(DataTable dt, int i)
        {
            POGInfo result = new POGInfo();
            result.ROLE_ID = (dt.Rows[i][POGInfo.Field.ROLE_ID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POGInfo.Field.ROLE_ID.ToString()]));
            result.TB = (dt.Rows[i][POGInfo.Field.TB.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POGInfo.Field.TB.ToString()]));
            result.ROLE_ID1 = (dt.Rows[i][POGInfo.Field.ROLE_ID1.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POGInfo.Field.ROLE_ID1.ToString()]));
            result.ROLE_NAME = (dt.Rows[i][POGInfo.Field.ROLE_NAME.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POGInfo.Field.ROLE_NAME.ToString()]));
            result.PASS_MIN_LEN = (dt.Rows[i][POGInfo.Field.PASS_MIN_LEN.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POGInfo.Field.PASS_MIN_LEN.ToString()]));
            result.PASS_VALID = (dt.Rows[i][POGInfo.Field.PASS_VALID.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POGInfo.Field.PASS_VALID.ToString()]));
            result.RPT_CODE = (dt.Rows[i][POGInfo.Field.RPT_CODE.ToString()] == DBNull.Value ? "" : Convert.ToString(dt.Rows[i][POGInfo.Field.RPT_CODE.ToString()]));
           
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
        public Int32 Add(POGInfo objEntr, ref string sErr)
        {
            int ret = -1;
            connect();
            InitSPCommand(_strSPInsertName);
            AddParameter(prefix + POGInfo.Field.ROLE_ID.ToString(), objEntr.ROLE_ID);
            AddParameter(prefix + POGInfo.Field.TB.ToString(), objEntr.TB);
            AddParameter(prefix + POGInfo.Field.ROLE_ID1.ToString(), objEntr.ROLE_ID1);
            AddParameter(prefix + POGInfo.Field.ROLE_NAME.ToString(), objEntr.ROLE_NAME);
            AddParameter(prefix + POGInfo.Field.PASS_MIN_LEN.ToString(), objEntr.PASS_MIN_LEN);
            AddParameter(prefix + POGInfo.Field.PASS_VALID.ToString(), objEntr.PASS_VALID);
            AddParameter(prefix + POGInfo.Field.RPT_CODE.ToString(), objEntr.RPT_CODE);
          
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

        public string Update(POGInfo objEntr)
        {
            connect();
            InitSPCommand(_strSPUpdateName);
            AddParameter(prefix + POGInfo.Field.ROLE_ID.ToString(), objEntr.ROLE_ID);
            AddParameter(prefix + POGInfo.Field.TB.ToString(), objEntr.TB);
            AddParameter(prefix + POGInfo.Field.ROLE_ID1.ToString(), objEntr.ROLE_ID1);
            AddParameter(prefix + POGInfo.Field.ROLE_NAME.ToString(), objEntr.ROLE_NAME);
            AddParameter(prefix + POGInfo.Field.PASS_MIN_LEN.ToString(), objEntr.PASS_MIN_LEN);
            AddParameter(prefix + POGInfo.Field.PASS_VALID.ToString(), objEntr.PASS_VALID);
            AddParameter(prefix + POGInfo.Field.RPT_CODE.ToString(), objEntr.RPT_CODE);
               
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
        String ROLE_ID
		)
        {
            connect();
            InitSPCommand(_strSPDeleteName);
            AddParameter(prefix + POGInfo.Field.ROLE_ID.ToString(), ROLE_ID);
              
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
		
		public DataTableCollection Get_Page(POGInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
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
        String ROLE_ID
		)
        {
            connect();
            InitSPCommand(_strSPIsExist);
            AddParameter(prefix + POGInfo.Field.ROLE_ID.ToString(), ROLE_ID);
              
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
		
		private string CreateWhereClause(POGInfo obj)
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
