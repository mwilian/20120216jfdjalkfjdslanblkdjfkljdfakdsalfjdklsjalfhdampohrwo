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
    public class DBAControl
    {
		#region Local Variable
        private DBADataAccess _objDAO;
		#endregion Local Variable
		
		#region Method
        public DBAControl()
        {
            _objDAO = new DBADataAccess();
        }
		
        public DBAInfo Get(
        String DB,
		ref string sErr)
        {
            return _objDAO.Get(
            DB,
			ref sErr);
        }
		
        public DataTable GetAll(
        ref string sErr)
        {
            return _objDAO.GetAll(
            ref sErr);
        }
		
        public Int32 Add(DBAInfo obj, ref string sErr)
        {
            return _objDAO.Add(obj, ref sErr);
        }
		
        public string Update(DBAInfo obj)
        {
            return _objDAO.Update(obj);
        }
		
        public string Delete(
        String DB
		)
        {
            return _objDAO.Delete(
            DB
			);
        }  
        public Boolean IsExist(
        String DB
		)
        {
            return _objDAO.IsExist(
            DB
			);
        } 
		      		
		public DataTableCollection Get_Page(DBAInfo obj, string orderBy, int pageIndex, int pageSize,ref String sErr)
        {
            return _objDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }
        
        public DataTable Search(String columnName, String columnValue, String condition, ref String sErr)
        {           
            return _objDAO.Search(columnName, columnValue, condition, "", ref  sErr);
        }
        public string InsertUpdate(DBAInfo obj)
        {
            string sErr = "";
            if (IsExist(
            obj.DB
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
			DBAInfo inf = new DBAInfo();
            return inf.ToDataTable();
        }
		
		public string TransferIn(DataRow row)
        {
            DBAInfo inf = new DBAInfo(row);
            return InsertUpdate(inf);
        }
		#endregion Method

    }
}
