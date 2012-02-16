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
    public class POGControl
    {
		#region Local Variable
        private POGDataAccess _objDAO;
		#endregion Local Variable
		
		#region Method
        public POGControl()
        {
            _objDAO = new POGDataAccess();
        }
		
        public POGInfo Get(
        String ROLE_ID,
		ref string sErr)
        {
            return _objDAO.Get(
            ROLE_ID,
			ref sErr);
        }
		
        public DataTable GetAll(
        ref string sErr)
        {
            return _objDAO.GetAll(
            ref sErr);
        }
		
        public Int32 Add(POGInfo obj, ref string sErr)
        {
            return _objDAO.Add(obj, ref sErr);
        }
		
        public string Update(POGInfo obj)
        {
            return _objDAO.Update(obj);
        }
		
        public string Delete(
        String ROLE_ID
		)
        {
            return _objDAO.Delete(
            ROLE_ID
			);
        }  
        public Boolean IsExist(
        String ROLE_ID
		)
        {
            return _objDAO.IsExist(
            ROLE_ID
			);
        } 
		      		
		public DataTableCollection Get_Page(POGInfo obj, string orderBy, int pageIndex, int pageSize,ref String sErr)
        {
            return _objDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }
        
        public DataTable Search(String columnName, String columnValue, String condition, ref String sErr)
        {           
            return _objDAO.Search(columnName, columnValue, condition, ref  sErr);
        }
        public string InsertUpdate(POGInfo obj)
        {
            string sErr = "";
            if (IsExist(
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
			POGInfo inf = new POGInfo();
            return inf.ToDataTable();
        }
		
		public string TransferIn(DataRow row)
        {
            POGInfo inf = new POGInfo(row);
            return InsertUpdate(inf);
        }
		#endregion Method

    }
}
