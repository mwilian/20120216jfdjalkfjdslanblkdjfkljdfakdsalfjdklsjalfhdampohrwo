using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace QueryBuilder
{
    public class CoreQDControl
    {
        #region Local Variable
        private CoreQDDataAccess _CoreQDDAO;
        #endregion Local Variable

        #region Method
        public CoreQDControl()
        {
            _CoreQDDAO = new CoreQDDataAccess();
        }

        public CoreQDInfo Get_CoreQD(
            String DTB,
            String QD_ID
        , ref string sErr)
        {
            return _CoreQDDAO.Get_CoreQD(
            DTB,
            QD_ID
            , ref sErr);
        }

        public DataTable GetAll_CoreQD(String DTB, ref string sErr)
        {
            return _CoreQDDAO.GetAll_CoreQD(DTB, ref sErr);
        }

        public Int32 Add_CoreQD(CoreQDInfo obj, ref string sErr)
        {
            return _CoreQDDAO.Add_CoreQD(obj, ref sErr);
        }

        public string Update_CoreQD(CoreQDInfo obj)
        {
            return _CoreQDDAO.Update_CoreQD(obj);
        }

        public string Delete_CoreQD(
            String DTB,
            String QD_ID
        )
        {
            return _CoreQDDAO.Delete_CoreQD(
            DTB,
            QD_ID
            );
        }
        public Boolean IsExist_CoreQD(
            String DTB,
            String QD_ID
        )
        {
            return _CoreQDDAO.IsExist_CoreQD(
            DTB,
            QD_ID
            );
        }

        /*  public string Delete_CoreQD(String arrID)
          {
              string kq = "";
              string[] arrStrID = arrID.Split(',');
              foreach (string strID in arrStrID)
              {
                  if (strID != "")
                  {
                      int ID = Convert.ToInt32(strID);
                      kq += _CoreQDDAO.Delete_CoreQD(ID);
                  }
              }
              return kq;
          } */

        public DataTableCollection Get_Page(CoreQDInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            return _CoreQDDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }

        public DataTable Search(String columnName, String columnValue, String condition, ref String sErr)
        {
            return _CoreQDDAO.Search(columnName, columnValue, condition, ref  sErr);
        }

        public String getMaxARJ(String kind, ref String sErr)
        {
            return _CoreQDDAO.getMaxARJ(kind, ref sErr);
        }
        public DataTable getALL_CoreQD_By_ARJ(String DTB, String kind, ref String _sErr)
        {
            return _CoreQDDAO.getALL_CoreQD_By_ARJ(DTB, kind, ref _sErr);
        }
        #endregion Method


        public bool IsExist(string dtb, string qdID)
        {
            return _CoreQDDAO.IsExist_CoreQD(dtb, qdID);
        }

        public DataTable GetAll_CoreQD_ByGroup(string dtb, string _strCategory, ref string sErr)
        {
            return _CoreQDDAO.GetAll_CoreQD_ByGroup(dtb, _strCategory, ref  sErr);
        }

        public DataTable GetAll_CoreQD_ByCate(string dtb, string _strCategory, ref string sErr)
        {
            return _CoreQDDAO.GetAll_CoreQD_ByCate(dtb, _strCategory, ref  sErr);
        }

        public DataTable GetTransferOut_CoreQD(string dtb, ref string sErr)
        {
            return _CoreQDDAO.GetTransferOut_CoreQD(dtb, ref sErr);
        }

        public DataTable ToTransferInStruct()
        {
            return _CoreQDDAO.ToTransferInStruct();
        }
        public void TransferIn(DataRow row, ref string sErr)
        {
            CoreQDInfo qdInfo = new CoreQDInfo(row);
            InsertUpdate_CoreQD(qdInfo, ref sErr);
            CoreQDDInfo qddInfo = new CoreQDDInfo();
            qddInfo.GetTransferIn(row);
            CoreQDDControl qddCtr = new CoreQDDControl();
            qddCtr.InsertUpdate_CoreQD(qddInfo, ref sErr);
        }

        public void InsertUpdate_CoreQD(CoreQDInfo CoreQDInfo, ref string sErr)
        {
            if (IsExist_CoreQD(CoreQDInfo.DTB, CoreQDInfo.QD_ID))
                sErr = Update_CoreQD(CoreQDInfo);
            else Add_CoreQD(CoreQDInfo, ref sErr);
        }

        public DataTable GetTransferOut_CoreQD(string DTB, string QD_CODE, ref string sErr)
        {
            if (QD_CODE != "")
                return _CoreQDDAO.GetTransferOut_CoreQD(DTB, QD_CODE, ref sErr);
            else
                return _CoreQDDAO.GetTransferOut_CoreQD(DTB, ref sErr);
        }

        public DataTable GetAll_CoreQD_USER(string database, string user, ref string sErr)
        {
            if (user == "TVC")
                return GetAll_CoreQD(database, ref sErr);
            else
                return _CoreQDDAO.GetAll_CoreQD_USER(database, user, ref sErr);
        }
    }
}
