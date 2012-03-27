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
    public class PODControl
    {
        #region Local Variable
        private PODDataAccess _PODDAO;
        #endregion Local Variable

        #region Method
        public PODControl()
        {
            _PODDAO = new PODDataAccess();
        }

        public PODInfo Get(
            String USER_ID
        , ref string sErr)
        {
            return _PODDAO.Get(
            USER_ID
            , ref sErr);
        }

        public DataTable GetAll(
        ref string sErr)
        {
            return _PODDAO.GetAll(
            ref sErr);
        }

        public Int32 Add(PODInfo obj, ref string sErr)
        {
            return _PODDAO.Add(obj, ref sErr);
        }

        public string Update(PODInfo obj)
        {
            return _PODDAO.Update(obj);
        }

        public string Delete(
            String USER_ID
        )
        {
            return _PODDAO.Delete(
            USER_ID
            );
        }
        public Boolean IsExist(
            String USER_ID
        )
        {
            return _PODDAO.IsExist(
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
                      kq += _PODDAO.Delete(ID);
                  }
              }
              return kq;
          } 
		
          public DataTableCollection Get_Page(PODInfo obj, string orderBy, int pageIndex, int pageSize,ref String sErr)
          {
              return _PODDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
          }*/

        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {
            return _PODDAO.Search(columnName, columnValue, condition, ref  sErr);
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



        public string InsertUpdate(PODInfo pODInfo)
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
