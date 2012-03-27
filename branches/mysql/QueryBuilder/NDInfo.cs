
using System;
using System.Data;
using System.Data.SqlClient;

using System.Collections;
using System.Collections.Generic;

using System.Diagnostics;

namespace QueryBuilder
{
    //  Class NDInfo
    [Serializable()]
    public class NDInfo
    {

        #region '" Business Properties and Methods "'

        //  declare members
        private string _category = string.Empty;
        private string _subCategory = string.Empty;
        // TRANSNOTUSED: Private Member Variable _lookUp
        // private string _lookUp = string.Empty; 
        private string _description = string.Empty;
        private string _shortDesc = string.Empty;
        private string _nonValidated = string.Empty;
        private string _amendable = string.Empty;

        //  Property Category
        public string Category
        {
            get
            {
                return _category;
            }
        }
        //  Property Code
        public string Code
        {
            get
            {
                return _category;
            }
        }

        //  Property SubCategory
        public string SubCategory
        {
            get
            {
                return _subCategory;
            }
        }

        //  Property LookUp
        public string LookUp
        {
            get
            {
                return _subCategory + " " + _shortDesc;
            }
        }

        //  Property Description
        public string Description
        {
            get
            {
               // return VniConverter.Convert(_description);
                return _description;
            }
        }

        //  Property ShortDesc
        public string ShortDesc
        {
            get
            {
                return _shortDesc;
            }
        }

        //  Property NonValidated
        public bool NonValidated
        {
            get
            {
                return _nonValidated.ToUpper().Trim() == "Y";
            }
        }

        //  Property Amendable
        public bool Amendable
        {
            get
            {
                return _amendable.ToUpper().Trim() == "Y";
            }

        }

        //  Method GetIdValue
        protected object GetIdValue()
        {
            return _category + ":" + _subCategory;
        }


        #endregion //  Business Properties and Methods

        #region '" Factory Methods "'

        //  Method GetNDInfo
        public static NDInfo GetNDInfo(SqlDataReader dr)
        {
            return new NDInfo(dr);
        }


        //  Method EmptyNDInfo
        public static NDInfo EmptyNDInfo(string Cat)
        {
            return new NDInfo(Cat);
        }

        // TRANSWARNING: Automatically generated because of optional parameter(s) 
        //  Method EmptyNDInfo
        public static NDInfo EmptyNDInfo()
        {
            return EmptyNDInfo("");
        }


        private NDInfo(string cat)
        {
            _category = cat;
        }

        private NDInfo(SqlDataReader dr)
        {
            Fetch(dr);
        }

        #endregion //  Factory Methods

        #region '" Data Access "'

        #region '" Data Access - Fetch "'

        //  Method Fetch
        private void Fetch(SqlDataReader dr)
        {
            FetchObject(dr);

        }


        //  Method FetchObject
        private void FetchObject(SqlDataReader dr)
        {

            _category = Security.NZ(dr["Category"].ToString(), "");
            // _subCategory = NZ(dr("SubCategory").ToString)
            // _lookUp = NZ(dr("LookUp").ToString)
            _description = Security.NZ(dr["Description"].ToString(), "");
            _shortDesc = Security.NZ(dr["ShortDesc"].ToString(), "");
            _nonValidated = Security.NZ(dr["NonValidated"].ToString(), "");
            // _amendable = NZ(dr("Amendable").ToString)
        }



        #endregion //  Data Access - Fetch

        #endregion

    }



}
