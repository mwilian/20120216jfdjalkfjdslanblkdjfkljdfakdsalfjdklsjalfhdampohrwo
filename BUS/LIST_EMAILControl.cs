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
    public class LIST_EMAILControl
    {
		#region Local Variable
        private LIST_EMAILDataAccess _objDAO;
		#endregion Local Variable
		
		#region Method
        public LIST_EMAILControl()
        {
            _objDAO = new LIST_EMAILDataAccess();
        }
		
        public LIST_EMAILInfo Get(
        String Mail,
		ref string sErr)
        {
            return _objDAO.Get(
            Mail,
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
		
        public Int32 Add(LIST_EMAILInfo obj, ref string sErr)
        {
            return _objDAO.Add(obj, ref sErr);
        }
		
        public string Update(LIST_EMAILInfo obj)
        {
            return _objDAO.Update(obj);
        }
		
        public string Delete(
        String Mail
		)
        {
            return _objDAO.Delete(
            Mail
			);
        }  
        public Boolean IsExist(
        String Mail
		)
        {
            return _objDAO.IsExist(
            Mail
			);
        } 
		      		
		public DataTableCollection Get_Page(LIST_EMAILInfo obj, string orderBy, int pageIndex, int pageSize,ref String sErr)
        {
            return _objDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }
        
        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {           
            return _objDAO.Search(columnName, columnValue, condition, tableName, ref  sErr);
        }
        public string InsertUpdate(LIST_EMAILInfo obj)
        {
            string sErr = "";
            if (IsExist(
            obj.Mail
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
			return LIST_EMAILInfo.ToDataTable();             
        }
		
		public string TransferIn(DataRow row)
        {
            LIST_EMAILInfo inf = new LIST_EMAILInfo(row);
            return InsertUpdate(inf);
        }
		#endregion Method

    }
}
