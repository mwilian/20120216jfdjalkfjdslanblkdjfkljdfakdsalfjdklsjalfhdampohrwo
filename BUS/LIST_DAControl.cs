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
    public class LIST_DAControl
    {
        #region Local Variable
        private LIST_DADataAccess _objDAO;
        #endregion Local Variable

        #region Method
        public LIST_DAControl()
        {
            _objDAO = new LIST_DADataAccess();
        }

        public LIST_DAInfo Get(
        String DAG_ID,
        ref string sErr)
        {
            return _objDAO.Get(
            DAG_ID,
            ref sErr);
        }

        public DataTable GetAll(
        ref string sErr)
        {
            return _objDAO.GetAll(
            ref sErr);
        }

        public Int32 Add(LIST_DAInfo obj, ref string sErr)
        {
            return _objDAO.Add(obj, ref sErr);
        }

        public string Update(LIST_DAInfo obj)
        {
            return _objDAO.Update(obj);
        }

        public string Delete(
        String DAG_ID
        )
        {
            return _objDAO.Delete(
            DAG_ID
            );
        }
        public Boolean IsExist(
        String DAG_ID
        )
        {
            return _objDAO.IsExist(
            DAG_ID
            );
        }

        public DataTableCollection Get_Page(LIST_DAInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            return _objDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }

        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {
            return _objDAO.Search(columnName, columnValue, condition, tableName, ref  sErr);
        }
        public string InsertUpdate(LIST_DAInfo obj)
        {
            string sErr = "";
            if (IsExist(
            obj.DAG_ID
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
            LIST_DAInfo inf = new LIST_DAInfo();
            return inf.ToDataTable();
        }

        public string TransferIn(DataRow row)
        {
            LIST_DAInfo inf = new LIST_DAInfo(row);
            return InsertUpdate(inf);
        }
        #endregion Method


        public DataTable GetPermission(string user, ref string sErr)
        {
            return _objDAO.GetPermission(user, ref  sErr);
        }
    }
}
