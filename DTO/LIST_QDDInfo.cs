using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
    public class LIST_QDDInfo
    {
        #region Local Variable
        private String _DTB;
        public String DTB
        {
            get { return _DTB; }
            set { _DTB = value; }
        }
        private String _QD_ID;
        public String QD_ID
        {
            get { return _QD_ID; }
            set { _QD_ID = value; }
        }
        private Int32 _QDD_ID;
        public Int32 QDD_ID
        {
            get { return _QDD_ID; }
            set { _QDD_ID = value; }
        }
        private String _CODE;
        public String CODE
        {
            get { return _CODE; }
            set { _CODE = value; }
        }
        private String _DESCRIPTN;
        public String DESCRIPTN
        {
            get { return _DESCRIPTN; }
            set { _DESCRIPTN = value; }
        }
        private String _F_TYPE;
        public String F_TYPE
        {
            get { return _F_TYPE; }
            set { _F_TYPE = value; }
        }
        private String _SORTING;
        public String SORTING
        {
            get { return _SORTING; }
            set { _SORTING = value; }
        }
        private String _AGREGATE;
        public String AGREGATE
        {
            get { return _AGREGATE; }
            set { _AGREGATE = value; }
        }
        private String _EXPRESSION;
        public String EXPRESSION
        {
            get { return _EXPRESSION; }
            set { _EXPRESSION = value; }
        }
        private String _FILTER_FROM;
        public String FILTER_FROM
        {
            get { return _FILTER_FROM; }
            set { _FILTER_FROM = value; }
        }
        private String _FILTER_TO;
        public String FILTER_TO
        {
            get { return _FILTER_TO; }
            set { _FILTER_TO = value; }
        }
        private Boolean _IS_FILTER;
        public Boolean IS_FILTER
        {
            get { return _IS_FILTER; }
            set { _IS_FILTER = value; }
        }
        #endregion LocalVariable

        #region Constructor
        public LIST_QDDInfo()
        {
            _DTB = "";
            _QD_ID = "";
            _QDD_ID = -1;
            _CODE = "";
            _DESCRIPTN = "";
            _F_TYPE = "";
            _SORTING = "";
            _AGREGATE = "";
            _EXPRESSION = "";
            _FILTER_FROM = "";
            _FILTER_TO = "";
            _IS_FILTER = true;
        }
        public LIST_QDDInfo(
            String DTB,
            String QD_ID,
            Int32 QDD_ID,
            String CODE,
            String DESCRIPTN,
            String F_TYPE,
            String SORTING,
            String AGREGATE,
            String EXPRESSION,
            String FILTER_FROM,
            String FILTER_TO,
            Boolean IS_FILTER
            )
        {
            _DTB = DTB;
            _QD_ID = QD_ID;
            _QDD_ID = QDD_ID;
            _CODE = CODE;
            _DESCRIPTN = DESCRIPTN;
            _F_TYPE = F_TYPE;
            _SORTING = SORTING;
            _AGREGATE = AGREGATE;
            _EXPRESSION = EXPRESSION;
            _FILTER_FROM = FILTER_FROM;
            _FILTER_TO = FILTER_TO;
            _IS_FILTER = IS_FILTER;
        }
        public LIST_QDDInfo(DataRow dr)
        {
            if (dr != null)
            {
                _DTB = dr["DTB"] == DBNull.Value ? "" : Convert.ToString(dr["DTB"]);
                _QD_ID = dr["QD_ID"] == DBNull.Value ? "" : Convert.ToString(dr["QD_ID"]);
                _QDD_ID = dr["QDD_ID"] == DBNull.Value ? -1 : Convert.ToInt32(dr["QDD_ID"]);
                _CODE = dr["CODE"] == DBNull.Value ? "" : Convert.ToString(dr["CODE"]);
                _DESCRIPTN = dr["DESCRIPTN"] == DBNull.Value ? "" : Convert.ToString(dr["DESCRIPTN"]);
                _F_TYPE = dr["F_TYPE"] == DBNull.Value ? "" : Convert.ToString(dr["F_TYPE"]);
                _SORTING = dr["SORTING"] == DBNull.Value ? "" : Convert.ToString(dr["SORTING"]);
                _AGREGATE = dr["AGREGATE"] == DBNull.Value ? "" : Convert.ToString(dr["AGREGATE"]);
                _EXPRESSION = dr["EXPRESSION"] == DBNull.Value ? "" : Convert.ToString(dr["EXPRESSION"]);
                _FILTER_FROM = dr["FILTER_FROM"] == DBNull.Value ? "" : Convert.ToString(dr["FILTER_FROM"]);
                _FILTER_TO = dr["FILTER_TO"] == DBNull.Value ? "" : Convert.ToString(dr["FILTER_TO"]);
                _IS_FILTER = dr["IS_FILTER"] == DBNull.Value ? true : Convert.ToBoolean(dr["IS_FILTER"]);
            }
        }
        #endregion Constructor

        #region InitTable
        public DataTable ToDataTable()
        {
            DataTable dt = new DataTable("TblLIST_QDD");
            dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn("DTB", typeof(String)),
				new DataColumn("QD_ID", typeof(String)),
				new DataColumn("QDD_ID", typeof(Int32)),
				new DataColumn("CODE", typeof(String)),
				new DataColumn("DESCRIPTN", typeof(String)),
				new DataColumn("F_TYPE", typeof(String)),
				new DataColumn("SORTING", typeof(String)),
				new DataColumn("AGREGATE", typeof(String)),
				new DataColumn("EXPRESSION", typeof(String)),
				new DataColumn("FILTER_FROM", typeof(String)),
				new DataColumn("FILTER_TO", typeof(String)),
				new DataColumn("IS_FILTER", typeof(Boolean))
				});
            return dt;
        }
        public DataRow ToDataRow(DataTable dt)
        {
            DataRow row = dt.NewRow();
            row["DTB"] = _DTB;
            row["QD_ID"] = _QD_ID;
            row["QDD_ID"] = _QDD_ID;
            row["CODE"] = _CODE;
            row["DESCRIPTN"] = _DESCRIPTN;
            row["F_TYPE"] = _F_TYPE;
            row["SORTING"] = _SORTING;
            row["AGREGATE"] = _AGREGATE;
            row["EXPRESSION"] = _EXPRESSION;
            row["FILTER_FROM"] = _FILTER_FROM;
            row["FILTER_TO"] = _FILTER_TO;
            row["IS_FILTER"] = _IS_FILTER;
            return row;
        }
        #endregion InitTable

        public void GetTransferIn(DataRow dr)
        {
            if (dr != null)
            {
                _DTB = dr["DTB"] == DBNull.Value ? "" : Convert.ToString(dr["DTB"]);
                _QD_ID = dr["QD_ID"] == DBNull.Value ? "" : Convert.ToString(dr["QD_ID"]);
                _QDD_ID = dr["QDD_ID"] == DBNull.Value ? -1 : Convert.ToInt32(dr["QDD_ID"]);
                _CODE = dr["CODE"] == DBNull.Value ? "" : Convert.ToString(dr["CODE"]);
                _DESCRIPTN = dr["DESCRIPTN"] == DBNull.Value ? "" : Convert.ToString(dr["QDD_DESCRIPTN"]);
                _F_TYPE = dr["F_TYPE"] == DBNull.Value ? "" : Convert.ToString(dr["F_TYPE"]);
                _SORTING = dr["SORTING"] == DBNull.Value ? "" : Convert.ToString(dr["SORTING"]);
                _AGREGATE = dr["AGREGATE"] == DBNull.Value ? "" : Convert.ToString(dr["AGREGATE"]);
                _EXPRESSION = dr["EXPRESSION"] == DBNull.Value ? "" : Convert.ToString(dr["EXPRESSION"]);
                _FILTER_FROM = dr["FILTER_FROM"] == DBNull.Value ? "" : Convert.ToString(dr["FILTER_FROM"]);
                _FILTER_TO = dr["FILTER_TO"] == DBNull.Value ? "" : Convert.ToString(dr["FILTER_TO"]);
                _IS_FILTER = dr["IS_FILTER"] == DBNull.Value ? true : Convert.ToBoolean(dr["IS_FILTER"]);
            }
        }
    }
}
