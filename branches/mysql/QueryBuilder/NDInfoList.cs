using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Xml.Xsl;
using System.ComponentModel;


using Microsoft.VisualBasic;
using System.Collections;
using System.Collections.Generic;
//using System.Drawing;
using System.Diagnostics;
//using System.Windows.Forms;
using System.Configuration;
namespace QueryBuilder
{
    //  Class NDInfoList
    [Serializable()]
    public class NDInfoList : BindingList<NDInfo>
    {
        private static string _DTB = string.Empty;
        private static NDInfoList _list;

        #region '" Factory Methods "'

        private NDInfoList(string DB)
        {
            _DTB = DB;
        }

        //  Method GetNDInfoList
        public static NDInfoList GetNDInfoList(string DB)
        {
            if (_list == null & _DTB != DB)
            {
                _DTB = DB;
                _list = new NDInfoList(DB);
                _list.DataPortal_Fetch();
            }
            return _list;
        }


        //  Method InvalidateCache
        public static void InvalidateCache()
        {
            _list = null;
        }


        //  Method ContainsCode
        public bool ContainsCode(string Code, ref NDInfo RetInfo)
        {

            foreach (NDInfo info in GetNDInfoList(_DTB))
            {
                if (info.Code == Code)
                {
                    RetInfo = info;
                    /* L:40 */
                    return true;
                }
            }
            return false;
        }

        // TRANSWARNING: Automatically generated because of optional parameter(s) 
        //  Method ContainsCode
        public bool ContainsCode(string Code)
        {
            NDInfo transTemp0 = null;
            return ContainsCode(Code, ref transTemp0);
        }


        //  Method GetDescription
        public string GetDescription(string CategoryCode)
        {
            NDInfo ND = null;
            if (ContainsCode(CategoryCode, ref ND))
            {
                return ND.Description;
            }
            else
            {
                // return SatResources.ResourcesProxy.ResStr("ANALYSIS") + CategoryCode;
                return "";
            }
        }


        #endregion //  Factory Methods

        #region '" Data Access "'

        #region '" Data Access - Fetch "'

        //  Method DataPortal_Fetch
        private void DataPortal_Fetch()
        {
            RaiseListChangedEvents = false;

            using (SqlConnection cn = new SqlConnection(ConfigurationSettings.AppSettings["strConnect"].ToString()))
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
                cm.CommandText = "dbo].[procAND_getall";
                cm.Parameters.AddWithValue("@SUN_DB", _DTB);
                //    cm.Parameters.AddWithValue("@CATEGORY", String.Empty)
                using (SqlDataReader dr = cm.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        NDInfo _Info = NDInfo.GetNDInfo(dr);
                        this.Add(_Info);
                    }
                }

            }

        }


        #endregion //  Data Access - Fetch

        #endregion //  Data Access

    }




}
