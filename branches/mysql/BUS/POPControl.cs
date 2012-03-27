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
    public class POPControl
    {
        #region Local Variable
        private POPDataAccess _POPDAO;
        #endregion Local Variable

        #region Method
        public POPControl()
        {
            _POPDAO = new POPDataAccess();
        }

        public POPInfo Get(
            String ROLE_ID,
            String DB
        , ref string sErr)
        {
            return _POPDAO.Get(
            ROLE_ID,
            DB
            , ref sErr);
        }

        public DataTable GetAll(
        ref string sErr)
        {
            return _POPDAO.GetAll(
            ref sErr);
        }

        public Int32 Add(POPInfo obj, ref string sErr)
        {
            return _POPDAO.Add(obj, ref sErr);
        }

        public string Update(POPInfo obj)
        {
            return _POPDAO.Update(obj);
        }

        public string Delete(
            String ROLE_ID,
            String DB
        )
        {
            return _POPDAO.Delete(
            ROLE_ID,
            DB
            );
        }
        public Boolean IsExist(
            String ROLE_ID,
            String DB
        )
        {
            return _POPDAO.IsExist(
            ROLE_ID,
            DB
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
                      kq += _POPDAO.Delete(ID);
                  }
              }
              return kq;
          } 
		
          public DataTableCollection Get_Page(POPInfo obj, string orderBy, int pageIndex, int pageSize,ref String sErr)
          {
              return _POPDAO.Get_Page(obj, orderBy, pageIndex, pageSize, ref sErr);
          }*/

        public DataTable Search(String columnName, String columnValue, String condition, String tableName, ref String sErr)
        {
            return _POPDAO.Search(columnName, columnValue, condition, ref  sErr);
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

        internal string TransferIn(DataRow row)
        {
            throw new NotImplementedException();
        }

        public string InsertUpdate(POPInfo pOPInfo)
        {
            string sErr = "";
            if (IsExist(pOPInfo.ROLE_ID, pOPInfo.DB))
                return Update(pOPInfo);
            else Add(pOPInfo, ref sErr);
            return sErr;
        }
    }
}
