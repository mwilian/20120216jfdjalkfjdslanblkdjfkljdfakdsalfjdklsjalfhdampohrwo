using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace QueryBuilder
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
    public class CoreQDInfo
    {
		#region Local Variable
		private String _DTB;
		private String _QD_ID;
		private String _DESCRIPTN;
		private String _OWNER;
		private Boolean _SHARED;
		private String _LAYOUT;
		private String _ANAL_Q0;
		private String _ANAL_Q9;
		private String _ANAL_Q8;
		private String _ANAL_Q7;
		private String _ANAL_Q6;
		private String _ANAL_Q5;
		private String _ANAL_Q4;
		private String _ANAL_Q3;
		private String _ANAL_Q2;
		private String _ANAL_Q1;
		private String _SQL_TEXT;
		private String _HEADER_TEXT;
		private String _FOOTER_TEXT;
		public String DTB
		{
			get { return _DTB; }
			set { _DTB = value; }
		}
		public String QD_ID
		{
			get { return _QD_ID; }
			set { _QD_ID = value; }
		}
		public String DESCRIPTN
		{
			get { return _DESCRIPTN; }
			set { _DESCRIPTN = value; }
		}
		public String OWNER
		{
			get { return _OWNER; }
			set { _OWNER = value; }
		}
		public Boolean SHARED
		{
			get { return _SHARED; }
			set { _SHARED = value; }
		}
		public String LAYOUT
		{
			get { return _LAYOUT; }
			set { _LAYOUT = value; }
		}
		public String ANAL_Q0
		{
			get { return _ANAL_Q0; }
			set { _ANAL_Q0 = value; }
		}
		public String ANAL_Q9
		{
			get { return _ANAL_Q9; }
			set { _ANAL_Q9 = value; }
		}
		public String ANAL_Q8
		{
			get { return _ANAL_Q8; }
			set { _ANAL_Q8 = value; }
		}
		public String ANAL_Q7
		{
			get { return _ANAL_Q7; }
			set { _ANAL_Q7 = value; }
		}
		public String ANAL_Q6
		{
			get { return _ANAL_Q6; }
			set { _ANAL_Q6 = value; }
		}
		public String ANAL_Q5
		{
			get { return _ANAL_Q5; }
			set { _ANAL_Q5 = value; }
		}
		public String ANAL_Q4
		{
			get { return _ANAL_Q4; }
			set { _ANAL_Q4 = value; }
		}
		public String ANAL_Q3
		{
			get { return _ANAL_Q3; }
			set { _ANAL_Q3 = value; }
		}
		public String ANAL_Q2
		{
			get { return _ANAL_Q2; }
			set { _ANAL_Q2 = value; }
		}
		public String ANAL_Q1
		{
			get { return _ANAL_Q1; }
			set { _ANAL_Q1 = value; }
		}
		public String SQL_TEXT
		{
			get { return _SQL_TEXT; }
			set { _SQL_TEXT = value; }
		}
		public String HEADER_TEXT
		{
			get { return _HEADER_TEXT; }
			set { _HEADER_TEXT = value; }
		}
		public String FOOTER_TEXT
		{
			get { return _FOOTER_TEXT; }
			set { _FOOTER_TEXT = value; }
		}
        #endregion LocalVariable
        
        #region Constructor
		public CoreQDInfo()
		{
			_DTB = "";
			_QD_ID = "";
			_DESCRIPTN = "";
			_OWNER = "";
			_SHARED = true;
			_LAYOUT = "";
			_ANAL_Q0 = "";
			_ANAL_Q9 = "";
			_ANAL_Q8 = "";
			_ANAL_Q7 = "";
			_ANAL_Q6 = "";
			_ANAL_Q5 = "";
			_ANAL_Q4 = "";
			_ANAL_Q3 = "";
			_ANAL_Q2 = "";
			_ANAL_Q1 = "";
			_SQL_TEXT = "";
			_HEADER_TEXT = "";
			_FOOTER_TEXT = "";
		}
		public CoreQDInfo(
			String DTB,
			String QD_ID,
			String DESCRIPTN,
			String OWNER,
			Boolean SHARED,
			String LAYOUT,
			String ANAL_Q0,
			String ANAL_Q9,
			String ANAL_Q8,
			String ANAL_Q7,
			String ANAL_Q6,
			String ANAL_Q5,
			String ANAL_Q4,
			String ANAL_Q3,
			String ANAL_Q2,
			String ANAL_Q1,
			String SQL_TEXT,
			String HEADER_TEXT,
			String FOOTER_TEXT
			)
		{
			_DTB = DTB;
			_QD_ID = QD_ID;
			_DESCRIPTN = DESCRIPTN;
			_OWNER = OWNER;
			_SHARED = SHARED;
			_LAYOUT = LAYOUT;
			_ANAL_Q0 = ANAL_Q0;
			_ANAL_Q9 = ANAL_Q9;
			_ANAL_Q8 = ANAL_Q8;
			_ANAL_Q7 = ANAL_Q7;
			_ANAL_Q6 = ANAL_Q6;
			_ANAL_Q5 = ANAL_Q5;
			_ANAL_Q4 = ANAL_Q4;
			_ANAL_Q3 = ANAL_Q3;
			_ANAL_Q2 = ANAL_Q2;
			_ANAL_Q1 = ANAL_Q1;
			_SQL_TEXT = SQL_TEXT;
			_HEADER_TEXT = HEADER_TEXT;
			_FOOTER_TEXT = FOOTER_TEXT;
		}
		public CoreQDInfo(DataRow dr)
		{
			if (dr != null)
			{
				_DTB = dr["DTB"] == DBNull.Value ? "" : Convert.ToString(dr["DTB"]);
				_QD_ID = dr["QD_ID"] == DBNull.Value ? "" : Convert.ToString(dr["QD_ID"]);
				_DESCRIPTN = dr["DESCRIPTN"] == DBNull.Value ? "" : Convert.ToString(dr["DESCRIPTN"]);
				_OWNER = dr["OWNER"] == DBNull.Value ? "" : Convert.ToString(dr["OWNER"]);
				_SHARED = dr["SHARED"] == DBNull.Value ? true : Convert.ToBoolean(dr["SHARED"]);
				_LAYOUT = dr["LAYOUT"] == DBNull.Value ? "" : Convert.ToString(dr["LAYOUT"]);
				_ANAL_Q0 = dr["ANAL_Q0"] == DBNull.Value ? "" : Convert.ToString(dr["ANAL_Q0"]);
				_ANAL_Q9 = dr["ANAL_Q9"] == DBNull.Value ? "" : Convert.ToString(dr["ANAL_Q9"]);
				_ANAL_Q8 = dr["ANAL_Q8"] == DBNull.Value ? "" : Convert.ToString(dr["ANAL_Q8"]);
				_ANAL_Q7 = dr["ANAL_Q7"] == DBNull.Value ? "" : Convert.ToString(dr["ANAL_Q7"]);
				_ANAL_Q6 = dr["ANAL_Q6"] == DBNull.Value ? "" : Convert.ToString(dr["ANAL_Q6"]);
				_ANAL_Q5 = dr["ANAL_Q5"] == DBNull.Value ? "" : Convert.ToString(dr["ANAL_Q5"]);
				_ANAL_Q4 = dr["ANAL_Q4"] == DBNull.Value ? "" : Convert.ToString(dr["ANAL_Q4"]);
				_ANAL_Q3 = dr["ANAL_Q3"] == DBNull.Value ? "" : Convert.ToString(dr["ANAL_Q3"]);
				_ANAL_Q2 = dr["ANAL_Q2"] == DBNull.Value ? "" : Convert.ToString(dr["ANAL_Q2"]);
				_ANAL_Q1 = dr["ANAL_Q1"] == DBNull.Value ? "" : Convert.ToString(dr["ANAL_Q1"]);
				_SQL_TEXT = dr["SQL_TEXT"] == DBNull.Value ? "" : Convert.ToString(dr["SQL_TEXT"]);
				_HEADER_TEXT = dr["HEADER_TEXT"] == DBNull.Value ? "" : Convert.ToString(dr["HEADER_TEXT"]);
				_FOOTER_TEXT = dr["FOOTER_TEXT"] == DBNull.Value ? "" : Convert.ToString(dr["FOOTER_TEXT"]);
			}
		}
        #endregion Constructor
        
        #region InitTable
		public DataTable ToDataTable()
		{
			DataTable dt = new DataTable("TblCoreQD");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn("DTB", typeof(String)),
				new DataColumn("QD_ID", typeof(String)),
				new DataColumn("DESCRIPTN", typeof(String)),
				new DataColumn("OWNER", typeof(String)),
				new DataColumn("SHARED", typeof(Boolean)),
				new DataColumn("LAYOUT", typeof(String)),
				new DataColumn("ANAL_Q0", typeof(String)),
				new DataColumn("ANAL_Q9", typeof(String)),
				new DataColumn("ANAL_Q8", typeof(String)),
				new DataColumn("ANAL_Q7", typeof(String)),
				new DataColumn("ANAL_Q6", typeof(String)),
				new DataColumn("ANAL_Q5", typeof(String)),
				new DataColumn("ANAL_Q4", typeof(String)),
				new DataColumn("ANAL_Q3", typeof(String)),
				new DataColumn("ANAL_Q2", typeof(String)),
				new DataColumn("ANAL_Q1", typeof(String)),
				new DataColumn("SQL_TEXT", typeof(String)),
				new DataColumn("HEADER_TEXT", typeof(String)),
				new DataColumn("FOOTER_TEXT", typeof(String))
				});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row["DTB"] = _DTB;
			row["QD_ID"] = _QD_ID;
			row["DESCRIPTN"] = _DESCRIPTN;
			row["OWNER"] = _OWNER;
			row["SHARED"] = _SHARED;
			row["LAYOUT"] = _LAYOUT;
			row["ANAL_Q0"] = _ANAL_Q0;
			row["ANAL_Q9"] = _ANAL_Q9;
			row["ANAL_Q8"] = _ANAL_Q8;
			row["ANAL_Q7"] = _ANAL_Q7;
			row["ANAL_Q6"] = _ANAL_Q6;
			row["ANAL_Q5"] = _ANAL_Q5;
			row["ANAL_Q4"] = _ANAL_Q4;
			row["ANAL_Q3"] = _ANAL_Q3;
			row["ANAL_Q2"] = _ANAL_Q2;
			row["ANAL_Q1"] = _ANAL_Q1;
			row["SQL_TEXT"] = _SQL_TEXT;
			row["HEADER_TEXT"] = _HEADER_TEXT;
			row["FOOTER_TEXT"] = _FOOTER_TEXT;
			return row;
		}
        #endregion InitTable
    }
}
