using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace QueryBuilder
{
    public class CoreQDDControl
    {
        #region Local Variable
        private CoreQDDDataAccess _CoreQDDDAO;
        #endregion Local Variable

        #region Method
        public CoreQDDControl()
        {
            _CoreQDDDAO = new CoreQDDDataAccess();
        }

        public CoreQDDInfo Get_CoreQDD(
            String DTB,
            String QD_ID,
            Int32 QDD_ID
        , ref string sErr)
        {
            return _CoreQDDDAO.Get_CoreQDD(
            DTB,
            QD_ID,
            QDD_ID
            , ref sErr);
        }

        public DataTable GetAll_CoreQDD(ref string sErr)
        {
            return _CoreQDDDAO.GetAll_CoreQDD(ref sErr);
        }

        public Int32 Add_CoreQDD(CoreQDDInfo obj, ref string sErr)
        {
            return _CoreQDDDAO.Add_CoreQDD(obj, ref sErr);
        }

        public string Update_CoreQDD(CoreQDDInfo obj)
        {
            return _CoreQDDDAO.Update_CoreQDD(obj);
        }

        public string Delete_CoreQDD(
            String DTB,
            String QD_ID,
            Int32 QDD_ID
        )
        {
            return _CoreQDDDAO.Delete_CoreQDD(
            DTB,
            QD_ID,
            QDD_ID
            );
        }
        public Boolean IsExist_CoreQDD(
            String DTB,
            String QD_ID,
            Int32 QDD_ID
        )
        {
            return _CoreQDDDAO.IsExist_CoreQDD(
            DTB,
            QD_ID,
            QDD_ID
            );
        }

        /*  public string Delete_CoreQDD(String arrID)
          {
              string kq = "";
              string[] arrStrID = arrID.Split(',');
              foreach (string strID in arrStrID)
              {
                  if (strID != "")
                  {
                      int ID = Convert.ToInt32(strID);
                      kq += _CoreQDDDAO.Delete_CoreQDD(ID);
                  }
              }
              return kq;
          } */

        public DataTableCollection Get_Page(CoreQDDInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            return _CoreQDDDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }

        public DataTable Search(String columnName, String columnValue, String condition, ref String sErr)
        {
            return _CoreQDDDAO.Search(columnName, columnValue, condition, ref  sErr);
        }

        public DataTable GetALL_CoreQDD_By_QD_ID(
            String DTB,
            String QD_ID

        , ref string sErr)
        {
            return _CoreQDDDAO.GetALL_CoreQDD_By_QD_ID(DTB, QD_ID, ref sErr);
        }

        public int get_QDDID(String DTB, String QD_ID, ref String sErr)
        {
            int numberline = 100;
            DataTable dt = GetALL_CoreQDD_By_QD_ID(DTB, QD_ID, ref sErr);
            if (dt.Rows.Count > 0)
            {
                numberline = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["QDD_ID"].ToString()) + 1;
            }
            return numberline;
        }

        #endregion Method


        public void Delete_CoreQDD_By_QD_ID(string qdID, string dtb, ref string sErr)
        {
            _CoreQDDDAO.Delete_CoreQDD_By_QD_ID(qdID, dtb, ref sErr);
        }

        public void InsertUpdate_CoreQD(CoreQDDInfo qddInfo,ref string sErr)
        {
            if (IsExist_CoreQDD(qddInfo.DTB, qddInfo.QD_ID, qddInfo.QDD_ID))
                Update_CoreQDD(qddInfo);
            else
                Add_CoreQDD(qddInfo, ref sErr);
        }
    }
}
