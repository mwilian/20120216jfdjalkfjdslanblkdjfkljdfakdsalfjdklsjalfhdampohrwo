using System;
using System.Collections.Generic;
using System.Text;
using DTO;
using DAO;
using System.Data;

namespace BUS
{
    public class LIST_QDControl
    {
        #region Local Variable
        private LIST_QDDataAccess _LIST_QDDAO;
        #endregion Local Variable

        #region Method
        public LIST_QDControl()
        {
            _LIST_QDDAO = new LIST_QDDataAccess();
        }

        public LIST_QDInfo Get_LIST_QD(
            String DTB,
            String QD_ID
        , ref string sErr)
        {
            return _LIST_QDDAO.Get_LIST_QD(
            DTB,
            QD_ID
            , ref sErr);
        }

        public DataTable GetAll_LIST_QD(String DTB, ref string sErr)
        {
            return _LIST_QDDAO.GetAll_LIST_QD(DTB, ref sErr);
        }

        public Int32 Add_LIST_QD(LIST_QDInfo obj, ref string sErr)
        {
            return _LIST_QDDAO.Add_LIST_QD(obj, ref sErr);
        }

        public string Update_LIST_QD(LIST_QDInfo obj)
        {
            return _LIST_QDDAO.Update_LIST_QD(obj);
        }

        public string Delete_LIST_QD(
            String DTB,
            String QD_ID
        )
        {
            return _LIST_QDDAO.Delete_LIST_QD(
            DTB,
            QD_ID
            );
        }
        public Boolean IsExist_LIST_QD(
            String DTB,
            String QD_ID
        )
        {
            return _LIST_QDDAO.IsExist_LIST_QD(
            DTB,
            QD_ID
            );
        }

        /*  public string Delete_LIST_QD(String arrID)
          {
              string kq = "";
              string[] arrStrID = arrID.Split(',');
              foreach (string strID in arrStrID)
              {
                  if (strID != "")
                  {
                      int ID = Convert.ToInt32(strID);
                      kq += _LIST_QDDAO.Delete_LIST_QD(ID);
                  }
              }
              return kq;
          } */

        public DataTableCollection Get_Page(LIST_QDInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            return _LIST_QDDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }

        public DataTable Search(String columnName, String columnValue, String condition, ref String sErr)
        {
            return _LIST_QDDAO.Search(columnName, columnValue, condition, ref  sErr);
        }

        public String getMaxARJ(String kind, ref String sErr)
        {
            return _LIST_QDDAO.getMaxARJ(kind, ref sErr);
        }
        public DataTable getALL_LIST_QD_By_ARJ(String DTB, String kind, ref String _sErr)
        {
            return _LIST_QDDAO.getALL_LIST_QD_By_ARJ(DTB, kind, ref _sErr);
        }
        #endregion Method


        public bool IsExist(string dtb, string qdID)
        {
            return _LIST_QDDAO.IsExist_LIST_QD(dtb, qdID);
        }

        public DataTable GetAll_LIST_QD_ByGroup(string dtb, string _strCategory, ref string sErr)
        {
            return _LIST_QDDAO.GetAll_LIST_QD_ByGroup(dtb, _strCategory, ref  sErr);
        }

        public DataTable GetAll_LIST_QD_ByCate(string dtb, string _strCategory, ref string sErr)
        {
            return _LIST_QDDAO.GetAll_LIST_QD_ByCate(dtb, _strCategory, ref  sErr);
        }

        public DataTable GetTransferOut_LIST_QD(string dtb, ref string sErr)
        {
            return _LIST_QDDAO.GetTransferOut_LIST_QD(dtb, ref sErr);
        }

        public DataTable ToTransferInStruct()
        {
            return _LIST_QDDAO.ToTransferInStruct();
        }
        public void TransferIn(DataRow row, ref string sErr)
        {
            DTO.LIST_QDInfo qdInfo = new LIST_QDInfo(row);
            InsertUpdate_LIST_QD(qdInfo, ref sErr);
            DTO.LIST_QDDInfo qddInfo = new LIST_QDDInfo();
            qddInfo.GetTransferIn(row);
            DTO.LIST_QDD_FILTERInfo qddFInfo = new LIST_QDD_FILTERInfo(row);
            BUS.LIST_QDDControl qddCtr = new LIST_QDDControl();
            qddCtr.InsertUpdate_LIST_QD(qddInfo, ref sErr);
            BUS.LIST_QDD_FILTERControl filterCtr = new LIST_QDD_FILTERControl();
            if (qddFInfo.OPERATOR != "" && qddFInfo.OPERATOR != "-")
                filterCtr.InsertUpdate(qddFInfo);
        }

        public void InsertUpdate_LIST_QD(LIST_QDInfo lIST_QDInfo, ref string sErr)
        {
            if (IsExist_LIST_QD(lIST_QDInfo.DTB, lIST_QDInfo.QD_ID))
                sErr = Update_LIST_QD(lIST_QDInfo);
            else Add_LIST_QD(lIST_QDInfo, ref sErr);
        }

        public DataTable GetTransferOut_LIST_QD(string DTB, string QD_CODE, ref string sErr)
        {
            if (QD_CODE != "")
                return _LIST_QDDAO.GetTransferOut_LIST_QD(DTB, QD_CODE, ref sErr);
            else
                return _LIST_QDDAO.GetTransferOut_LIST_QD(DTB, ref sErr);
        }

        public DataTable GetAll_LIST_QD_USER(string database, string user, ref string sErr)
        {
            if (user == "TVC")
                return GetAll_LIST_QD(database, ref sErr);
            else
                return _LIST_QDDAO.GetAll_LIST_QD_USER(database, user, ref sErr);
        }
    }
}
