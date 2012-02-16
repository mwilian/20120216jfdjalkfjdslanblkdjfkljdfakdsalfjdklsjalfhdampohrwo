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
    public class LIST_QDD_FILTERControl
    {
        #region Local Variable
        private LIST_QDD_FILTERDataAccess _objDAO;
        #endregion Local Variable

        #region Method
        public LIST_QDD_FILTERControl()
        {
            _objDAO = new LIST_QDD_FILTERDataAccess();
        }

        public LIST_QDD_FILTERInfo Get(
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
            return _objDAO.GetByPos(
            DTB,
             pos, ref sErr);
        }
        public int GetCountRecord(
        String DTB,
        ref string sErr)
        {
            return _objDAO.GetCountRecord(
            DTB,
            ref sErr);
        }

        public Int32 Add(LIST_QDD_FILTERInfo obj, ref string sErr)
        {
            return _objDAO.Add(obj, ref sErr);
        }

        public string Update(LIST_QDD_FILTERInfo obj)
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

        public DataTableCollection Get_Page(LIST_QDD_FILTERInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            return _objDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }

        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {
            return _objDAO.Search(columnName, columnValue, condition, tableName, ref  sErr);
        }
        public string InsertUpdate(LIST_QDD_FILTERInfo obj)
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
            return LIST_QDD_FILTERInfo.ToDataTable();
        }

        public string TransferIn(DataRow row)
        {
            LIST_QDD_FILTERInfo inf = new LIST_QDD_FILTERInfo(row);
            return InsertUpdate(inf);
        }
        #endregion Method


        public void DeleteByQD_ID(string QD_ID, string DTB, ref string sErr)
        {
            _objDAO.DeleteByQD_ID(QD_ID, DTB, ref  sErr);
        }
    }
}
