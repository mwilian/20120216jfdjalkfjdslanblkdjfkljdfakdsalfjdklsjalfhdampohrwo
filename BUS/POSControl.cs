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
    public class POSControl
    {
        #region Local Variable
        private POSDataAccess _POSDAO;
        #endregion Local Variable

        #region Method
        public POSControl()
        {
            _POSDAO = new POSDataAccess();
        }

        public POSInfo Get(
        String USER_ID,
        ref string sErr)
        {
            return _POSDAO.Get(
            USER_ID,
            ref sErr);
        }

        public DataTable GetAll(
        ref string sErr)
        {
            return _POSDAO.GetAll(
            ref sErr);
        }

        public Int32 Add(POSInfo obj, ref string sErr)
        {
            return _POSDAO.Add(obj, ref sErr);
        }

        public string Update(POSInfo obj)
        {
            return _POSDAO.Update(obj);
        }

        public string Delete(
        String USER_ID
        )
        {
            return _POSDAO.Delete(
            USER_ID
            );
        }
        public Boolean IsExist(
        String USER_ID
        )
        {
            return _POSDAO.IsExist(
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
                      kq += _POSDAO.Delete(ID);
                  }
              }
              return kq;
          } */

        public DataTableCollection Get_Page(POSInfo obj, string orderBy, int pageIndex, int pageSize, ref String sErr)
        {
            return _POSDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
        }

        public DataTable Search(String columnName, String columnValue, String condition, ref String sErr)
        {
            return _POSDAO.Search(columnName, columnValue, condition, ref  sErr);
        }
        public string InsertUpdate(POSInfo obj)
        {
            string sErr = "";
            if (IsExist(
            obj.USER_ID
            ))
            {
                sErr = Update(obj);
            }
            else
                Add(obj, ref sErr);
            return sErr;
        }
        public int GetCount(ref string sErr)
        {
            return _POSDAO.GetCount(ref  sErr);
        }
        #endregion Method

    }
}
