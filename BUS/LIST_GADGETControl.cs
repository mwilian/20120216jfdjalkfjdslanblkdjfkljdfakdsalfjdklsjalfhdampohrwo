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
    public class LIST_GADGETControl
    {
		#region Local Variable
        private LIST_GADGETDataAccess _objDAO;
		#endregion Local Variable
		
		#region Method
        public LIST_GADGETControl()
        {
            _objDAO = new LIST_GADGETDataAccess();
        }
		
        public LIST_GADGETInfo Get(
        Int32 ID,
		ref string sErr)
        {
            return _objDAO.Get(
            ID,
			ref sErr);
        }
		
        public DataTable GetAll(
        ref string sErr)
        {
            return _objDAO.GetAll(
            ref sErr);
        }
		public DataTable GetByPos(
        int pos, ref string sErr)
        {
            return _objDAO.GetByPos(
            pos, ref sErr);
        }
		public int GetCountRecord(
        ref string sErr)
        {
            return _objDAO.GetCountRecord(
            ref sErr);
        }
		
        public Int32 Add(LIST_GADGETInfo obj, ref string sErr)
        {
            return _objDAO.Add(obj, ref sErr);
        }
		
        public string Update(LIST_GADGETInfo obj)
        {
            return _objDAO.Update(obj);
        }
		
        public string Delete(
        Int32 ID
		)
        {
            return _objDAO.Delete(
            ID
			);
        }  
        public Boolean IsExist(
        Int32 ID
		)
        {
            return _objDAO.IsExist(
            ID
			);
        } 
		      		
		public DataTableCollection Get_Page(LIST_GADGETInfo obj, string orderBy, int pageIndex, int pageSize,ref String sErr)
        {
            return _objDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }
        
        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {           
            return _objDAO.Search(columnName, columnValue, condition, tableName, ref  sErr);
        }
        public string InsertUpdate(LIST_GADGETInfo obj)
        {
            string sErr = "";
            if (IsExist(
            obj.ID
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
			return LIST_GADGETInfo.ToDataTable();             
        }
		
		public string TransferIn(DataRow row)
        {
            LIST_GADGETInfo inf = new LIST_GADGETInfo(row);
            return InsertUpdate(inf);
        }
		#endregion Method

    }
}
