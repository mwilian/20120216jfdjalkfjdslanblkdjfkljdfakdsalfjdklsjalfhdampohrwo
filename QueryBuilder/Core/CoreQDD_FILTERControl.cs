using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace QueryBuilder
{
    /// <summary> 
    ///Author: nnamthach@gmail.com 
    /// <summary>
    public class CoreQDD_FILTERControl
    {
        #region Local Variable
        private CoreQDD_FILTERDataAccess _objDAO;
        #endregion Local Variable

        #region Method
        public CoreQDD_FILTERControl()
        {
            _objDAO = new CoreQDD_FILTERDataAccess();
        }

        public CoreQDD_FILTERInfo Get(
        String DTB,
        String QD_ID,
        Int32 QDD_ID,
        ref string sErr)
        {
            return _objDAO.Get(
            DTB,
            QD_ID,
            QDD_ID,
            ref sErr);
        }

        public DataTable GetAll(
        String DTB,
        ref string sErr)
        {
            return _objDAO.GetAll(
            DTB,
            ref sErr);
        }
        public DataTable GetByPos(
        String DTB,
        int pos, ref string sErr)
        {
            return _objDAO.GetByPos(DTB, pos, ref sErr);
        }
        public int GetCountRecord(
        String DTB,
        ref string sErr)
        {
            return _objDAO.GetCountRecord(
            DTB,
            ref sErr);
        }

        public Int32 Add(CoreQDD_FILTERInfo obj, ref string sErr)
        {
            return _objDAO.Add(obj, ref sErr);
        }

        public string Update(CoreQDD_FILTERInfo obj)
        {
            return _objDAO.Update(obj);
        }

        public string Delete(
        String DTB,
        String QD_ID,
        Int32 QDD_ID
        )
        {
            return _objDAO.Delete(
            DTB,
            QD_ID,
            QDD_ID
            );
        }
        public Boolean IsExist(
        String DTB,
        String QD_ID,
        Int32 QDD_ID
        )
        {
            return _objDAO.IsExist(
            DTB,
            QD_ID,
            QDD_ID
            );
        }

        public DataTableCollection Get_Page(CoreQDD_FILTERInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            return _objDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }

        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {
            return _objDAO.Search(columnName, columnValue, condition, tableName, ref  sErr);
        }
        public string InsertUpdate(CoreQDD_FILTERInfo obj)
        {
            string sErr = "";
            if (IsExist(
            obj.DTB,
            obj.QD_ID,
            obj.QDD_ID
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
            return CoreQDD_FILTERInfo.ToDataTable();
        }

        public string TransferIn(DataRow row)
        {
            CoreQDD_FILTERInfo inf = new CoreQDD_FILTERInfo(row);
            return InsertUpdate(inf);
        }
        #endregion Method

    }
}
