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
    public class CorePODControl
    {
        #region Local Variable
        private CorePODDataAccess _CorePODDAO;
        #endregion Local Variable

        #region Method
        public CorePODControl()
        {
            _CorePODDAO = new CorePODDataAccess();
        }

        public CorePODInfo Get(
            String USER_ID
        , ref string sErr)
        {
            return _CorePODDAO.Get(
            USER_ID
            , ref sErr);
        }

        public DataTable GetAll(
        ref string sErr)
        {
            return _CorePODDAO.GetAll(
            ref sErr);
        }

        public Int32 Add(CorePODInfo obj, ref string sErr)
        {
            return _CorePODDAO.Add(obj, ref sErr);
        }

        public string Update(CorePODInfo obj)
        {
            return _CorePODDAO.Update(obj);
        }

        public string Delete(
            String USER_ID
        )
        {
            return _CorePODDAO.Delete(
            USER_ID
            );
        }
        public Boolean IsExist(
            String USER_ID
        )
        {
            return _CorePODDAO.IsExist(
            USER_ID
            );
        }

        /*  public string Delete(String arrID)
          {
              string kq = "";
              string[] arrStrID = arrID.Split(',');
              foreach (string strID in arrStrID)
              {
                  if (strID != "")
                  {
                      int ID = Convert.ToInt32(strID);
                      kq += _CorePODDAO.Delete(ID);
                  }
              }
              return kq;
          } 
		
          public DataTableCollection Get_Page(CorePODInfo obj, string orderBy, int pageIndex, int pageSize,ref String sErr)
          {
              return _CorePODDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
          }*/

        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {
            return _CorePODDAO.Search(columnName, columnValue, condition, ref  sErr);
        }

        #endregion Method


        internal DataTable GetTransferOut(string dtb, string from, string to, ref string sErr)
        {
            throw new NotImplementedException();
        }

        internal DataTable ToTransferInStruct()
        {
            throw new NotImplementedException();
        }



        public string InsertUpdate(CorePODInfo pODInfo)
        {
            string sErr = "";
            if (IsExist(pODInfo.USER_ID))
                return Update(pODInfo);
            else
                Add(pODInfo, ref sErr);
            return sErr;
        }
    }
}
