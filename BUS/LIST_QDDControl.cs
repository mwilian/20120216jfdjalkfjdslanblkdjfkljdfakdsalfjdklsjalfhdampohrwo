using System;
using System.Collections.Generic;
using System.Text;
using DTO;
using DAO;
using System.Data;

namespace BUS
{
    public class LIST_QDDControl
    {
        #region Local Variable
        private LIST_QDDDataAccess _LIST_QDDDAO;
        #endregion Local Variable

        #region Method
        public LIST_QDDControl()
        {
            _LIST_QDDDAO = new LIST_QDDDataAccess();
        }

        public LIST_QDDInfo Get_LIST_QDD(
            String DTB,
            String QD_ID,
            Int32 QDD_ID
        , ref string sErr)
        {
            return _LIST_QDDDAO.Get_LIST_QDD(
            DTB,
            QD_ID,
            QDD_ID
            , ref sErr);
        }

        public DataTable GetAll_LIST_QDD(ref string sErr)
        {
            return _LIST_QDDDAO.GetAll_LIST_QDD(ref sErr);
        }

        public Int32 Add_LIST_QDD(LIST_QDDInfo obj, ref string sErr)
        {
            return _LIST_QDDDAO.Add_LIST_QDD(obj, ref sErr);
        }

        public string Update_LIST_QDD(LIST_QDDInfo obj)
        {
            return _LIST_QDDDAO.Update_LIST_QDD(obj);
        }

        public string Delete_LIST_QDD(
            String DTB,
            String QD_ID,
            Int32 QDD_ID
        )
        {
            return _LIST_QDDDAO.Delete_LIST_QDD(
            DTB,
            QD_ID,
            QDD_ID
            );
        }
        public Boolean IsExist_LIST_QDD(
            String DTB,
            String QD_ID,
            Int32 QDD_ID
        )
        {
            return _LIST_QDDDAO.IsExist_LIST_QDD(
            DTB,
            QD_ID,
            QDD_ID
            );
        }

        /*  public string Delete_LIST_QDD(String arrID)
          {
              string kq = "";
              string[] arrStrID = arrID.Split(',');
              foreach (string strID in arrStrID)
              {
                  if (strID != "")
                  {
                      int ID = Convert.ToInt32(strID);
                      kq += _LIST_QDDDAO.Delete_LIST_QDD(ID);
                  }
              }
              return kq;
          } */

        public DataTableCollection Get_Page(LIST_QDDInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            return _LIST_QDDDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }

        public DataTable Search(String columnName, String columnValue, String condition, ref String sErr)
        {
            return _LIST_QDDDAO.Search(columnName, columnValue, condition, ref  sErr);
        }

        public DataTable GetALL_LIST_QDD_By_QD_ID(
            String DTB,
            String QD_ID

        , ref string sErr)
        {
            return _LIST_QDDDAO.GetALL_LIST_QDD_By_QD_ID(DTB, QD_ID, ref sErr);
        }

        public int get_QDDID(String DTB, String QD_ID, ref String sErr)
        {
            int numberline = 100;
            DataTable dt = GetALL_LIST_QDD_By_QD_ID(DTB, QD_ID, ref sErr);
            if (dt.Rows.Count > 0)
            {
                numberline = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["QDD_ID"].ToString()) + 1;
            }
            return numberline;
        }

        #endregion Method


        public void Delete_LIST_QDD_By_QD_ID(string qdID, string dtb, ref string sErr)
        {
            _LIST_QDDDAO.Delete_LIST_QDD_By_QD_ID(qdID, dtb, ref sErr);
        }

        public void InsertUpdate_LIST_QD(LIST_QDDInfo qddInfo,ref string sErr)
        {
            if (IsExist_LIST_QDD(qddInfo.DTB, qddInfo.QD_ID, qddInfo.QDD_ID))
                Update_LIST_QDD(qddInfo);
            else
                Add_LIST_QDD(qddInfo, ref sErr);
        }
    }
}
