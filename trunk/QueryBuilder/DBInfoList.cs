
using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Xml.Xsl;
//using pbs.Helper;
using System.ComponentModel;
using Microsoft.VisualBasic;
using System.Collections;
using System.Collections.Generic;

using System.Diagnostics;

using System.Configuration;
namespace QueryBuilder
{
    //  Class DBInfoList
    [Serializable()]
    public class DBInfoList : BindingList<DBInfo>
    {

        private static DBInfoList _list;

        #region '" Factory Methods "'

        private DBInfoList()
        {
            // require use of factory method
        }

        //  Method GetDBInfo
        public static DBInfo GetDBInfo(string DBCode)
        {
            DBInfo Info = DBInfo.EmptyDBInfo();
            ContainsCode(DBCode, ref Info);
            return Info;
        }

        //  Method GetDBInfoList
        public static DBInfoList GetDBInfoList()
        {
            try
            {
                if (_list == null)
                {
                    _list = new DBInfoList();
                    _list.DataPortal_Fetch();
                }
                return _list;
            }
            catch { return null; }

        }

        //  Method InvalidateCache
        public static void InvalidateCache()
        {
            _list = null;
        }


        //  Method ContainsCode
        public static bool ContainsCode(string Code, ref DBInfo RetInfo)
        {
            if (Code == "***")
            {
                return true;
            }
            foreach (DBInfo info in GetDBInfoList())
            {
                if (info.Code == Code)
                {
                    RetInfo = info;
                    return true;
                }
            }
            return false;
        }

        // TRANSWARNING: Automatically generated because of optional parameter(s) 
        //  Method ContainsCode
        public static bool ContainsCode(string Code)
        {
            DBInfo transTemp0 = null;
            return ContainsCode(Code, ref transTemp0);
        }

        #endregion //  Factory Methods

        #region '" Data Access "'

        #region '" Data Access - Fetch "'

        //  Method DataPortal_Fetch
        private void DataPortal_Fetch()
        {
            RaiseListChangedEvents = false;

            using (SqlConnection cn = new SqlConnection(CoreCommonControl.GetConnection()))
            {
                cn.Open();
                ExecuteFetch(cn);
            }


            RaiseListChangedEvents = true;
        }


        //  Method ExecuteFetch
        private void ExecuteFetch(SqlConnection cn)
        {
            using (SqlCommand cm = cn.CreateCommand())
            {
                cm.CommandType = CommandType.StoredProcedure;
                cm.CommandText = "dbo.[procDBA_getall]";

                using (SqlDataReader dr = cm.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        DBInfo _dbInfo = DBInfo.GetDBInfo(dr);
                        this.Add(DBInfo.GetDBInfo(dr));
                    }
                }

            }

        }


        #endregion //  Data Access - Fetch

        #endregion //  Data Access

    }

}


