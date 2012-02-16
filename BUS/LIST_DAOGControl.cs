using System;
using System.Collections.Generic;
using System.Text;
using DTO;
using DAO;
using System.Data;

namespace BUS
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
    public class LIST_DAOGControl
    {
		#region Local Variable
        private LIST_DAOGDataAccess _objDAO;
		#endregion Local Variable
		
		#region Method
        public LIST_DAOGControl()
        {
            _objDAO = new LIST_DAOGDataAccess();
        }
		
        public LIST_DAOGInfo Get(
        String DAG_ID,
        String ROLE_ID,
		ref string sErr)
        {
            return _objDAO.Get(
            DAG_ID,
            ROLE_ID,
			ref sErr);
        }
		
        public DataTable GetAll(string DAG_ID,
        ref string sErr)
        {
            return _objDAO.GetAll(DAG_ID,
            ref sErr);
        }
		
        public Int32 Add(LIST_DAOGInfo obj, ref string sErr)
        {
            return _objDAO.Add(obj, ref sErr);
        }
		
        public string Update(LIST_DAOGInfo obj)
        {
            return _objDAO.Update(obj);
        }
		
        public string Delete(
        String DAG_ID,
        String ROLE_ID
		)
        {
            return _objDAO.Delete(
            DAG_ID,
            ROLE_ID
			);
        }  
        public Boolean IsExist(
        String DAG_ID,
        String ROLE_ID
		)
        {
            return _objDAO.IsExist(
            DAG_ID,
            ROLE_ID
			);
        } 
		      		
		public DataTableCollection Get_Page(LIST_DAOGInfo obj, string orderBy, int pageIndex, int pageSize,ref String sErr)
        {
            return _objDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }
        
        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {           
            return _objDAO.Search(columnName, columnValue, condition, tableName, ref  sErr);
        }
        public string InsertUpdate(LIST_DAOGInfo obj)
        {
            string sErr = "";
            if (IsExist(
            obj.DAG_ID,
            obj.ROLE_ID
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
			LIST_DAOGInfo inf = new LIST_DAOGInfo();
            return inf.ToDataTable();
        }
		
		public string TransferIn(DataRow row)
        {
            LIST_DAOGInfo inf = new LIST_DAOGInfo(row);
            return InsertUpdate(inf);
        }
		#endregion Method


        public string Deletes(string DAG_ID)
        {
            return _objDAO.Deletes(DAG_ID);
        }
    }
}
