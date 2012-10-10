using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace QueryBuilder
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
    public class CoreQD_SCHEMAControl
    {
		#region Local Variable
        private CoreQD_SCHEMADataAccess _objDAO;
		#endregion Local Variable
		
		#region Method
        public CoreQD_SCHEMAControl()
        {
            _objDAO = new CoreQD_SCHEMADataAccess();
        }
		
        public CoreQD_SCHEMAInfo Get(
        String CONN_ID,
        String SCHEMA_ID,
		ref string sErr)
        {
            return _objDAO.Get(
            CONN_ID,
            SCHEMA_ID,
			ref sErr);
        }
      
		
        public DataTable GetAll(string conn,
        ref string sErr)
        {
            return _objDAO.GetAll(conn,
            ref sErr);
        }
		
        public Int32 Add(CoreQD_SCHEMAInfo obj, ref string sErr)
        {
            return _objDAO.Add(obj, ref sErr);
        }
		
        public string Update(CoreQD_SCHEMAInfo obj)
        {
            return _objDAO.Update(obj);
        }
		
        public string Delete(
        String CONN_ID,
        String SCHEMA_ID
		)
        {
            return _objDAO.Delete(
            CONN_ID,
            SCHEMA_ID
			);
        }  
        public Boolean IsExist(
        String CONN_ID,
        String SCHEMA_ID
		)
        {
            return _objDAO.IsExist(
            CONN_ID,
            SCHEMA_ID
			);
        } 
		      		
		public DataTableCollection Get_Page(CoreQD_SCHEMAInfo obj, string orderBy, int pageIndex, int pageSize,ref String sErr)
        {
            return _objDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }
        
        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {           
            return _objDAO.Search(columnName, columnValue, condition, tableName, ref  sErr);
        }
        public string InsertUpdate(CoreQD_SCHEMAInfo obj)
        {
            string sErr = "";
            if (IsExist(
            obj.CONN_ID,
            obj.SCHEMA_ID
			))
            {
                sErr = Update(obj);
            }
            else
                Add(obj, ref sErr);
            return sErr;
        }
		
        public DataTable GetTransferOut(string dtb, object from, object to, ref string sErr)
        {
            return _objDAO.GetTransferOut(dtb, from, to, ref sErr);
        }

        public DataTable ToTransferInStruct()
        {
			CoreQD_SCHEMAInfo inf = new CoreQD_SCHEMAInfo();
            return inf.ToDataTable();
        }
		
		public string TransferIn(DataRow row)
        {
            CoreQD_SCHEMAInfo inf = new CoreQD_SCHEMAInfo(row);
            return InsertUpdate(inf);
        }
		#endregion Method
        public string GetField(
      String CONN_ID,
      String SCHEMA_ID,
      ref string sErr)
        {
            return _objDAO.GetField(
            CONN_ID,
            SCHEMA_ID,
            ref sErr);
        }

        public DataTable GetJoins(string dtb, ref string sErr)
        {
            return _objDAO.GetJoins(
           dtb,           
           ref sErr);
        }
        public string GetDefaultDB(string db, string table, ref string sErr)
        {
            return _objDAO.GetDefaultDB(
            db, table,
            ref sErr);
        }
    }
}
