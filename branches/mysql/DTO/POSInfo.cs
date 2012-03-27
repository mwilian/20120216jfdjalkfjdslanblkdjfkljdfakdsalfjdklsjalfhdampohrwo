using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class POSInfo
    {
		#region Local Variable
		public enum Field
		{
			USER_ID,
			CURRENT_DB,
			CURRENT_ACTIVITY,
			WORK_STATION,
			LOGIN_TIME
		}
		private String _USER_ID;
		private String _CURRENT_DB;
		private String _CURRENT_ACTIVITY;
		private String _WORK_STATION;
		private String _LOGIN_TIME;
		
		public String USER_ID{	get{ return _USER_ID;} set{_USER_ID = value;} }
		public String CURRENT_DB{	get{ return _CURRENT_DB;} set{_CURRENT_DB = value;} }
		public String CURRENT_ACTIVITY{	get{ return _CURRENT_ACTIVITY;} set{_CURRENT_ACTIVITY = value;} }
		public String WORK_STATION{	get{ return _WORK_STATION;} set{_WORK_STATION = value;} }
		public String LOGIN_TIME{	get{ return _LOGIN_TIME;} set{_LOGIN_TIME = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public POSInfo()
		{
			_USER_ID = "";
			_CURRENT_DB = "";
			_CURRENT_ACTIVITY = "";
			_WORK_STATION = "";
			_LOGIN_TIME = "";
		}
		public POSInfo(
		String USER_ID,
		String CURRENT_DB,
		String CURRENT_ACTIVITY,
		String WORK_STATION,
		String LOGIN_TIME
		)
		{
			_USER_ID = USER_ID;
			_CURRENT_DB = CURRENT_DB;
			_CURRENT_ACTIVITY = CURRENT_ACTIVITY;
			_WORK_STATION = WORK_STATION;
			_LOGIN_TIME = LOGIN_TIME;
		}
		public POSInfo(DataRow dr)
		{
			if (dr != null)
			{
				_USER_ID = dr["USER_ID"] == DBNull.Value?"":Convert.ToString(dr["USER_ID"]);
				_CURRENT_DB = dr["CURRENT_DB"] == DBNull.Value?"":Convert.ToString(dr["CURRENT_DB"]);
				_CURRENT_ACTIVITY = dr["CURRENT_ACTIVITY"] == DBNull.Value?"":Convert.ToString(dr["CURRENT_ACTIVITY"]);
				_WORK_STATION = dr["WORK_STATION"] == DBNull.Value?"":Convert.ToString(dr["WORK_STATION"]);
				_LOGIN_TIME = dr["LOGIN_TIME"] == DBNull.Value?"":Convert.ToString(dr["LOGIN_TIME"]);
			}
		}
		public POSInfo(POSInfo objEntr)
		{			
			_USER_ID = objEntr.USER_ID;			
			_CURRENT_DB = objEntr.CURRENT_DB;			
			_CURRENT_ACTIVITY = objEntr.CURRENT_ACTIVITY;			
			_WORK_STATION = objEntr.WORK_STATION;			
			_LOGIN_TIME = objEntr.LOGIN_TIME;			
		}
        #endregion Constructor
        
        #region InitTable
		public DataTable ToDataTable()
		{
			DataTable dt = new DataTable("POS");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn("USER_ID", typeof(String)),
				new DataColumn("CURRENT_DB", typeof(String)),
				new DataColumn("CURRENT_ACTIVITY", typeof(String)),
				new DataColumn("WORK_STATION", typeof(String)),
				new DataColumn("LOGIN_TIME", typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row["USER_ID"] = _USER_ID;
			row["CURRENT_DB"] = _CURRENT_DB;
			row["CURRENT_ACTIVITY"] = _CURRENT_ACTIVITY;
			row["WORK_STATION"] = _WORK_STATION;
			row["LOGIN_TIME"] = _LOGIN_TIME;
			return row;
		}
        #endregion InitTable
    }
}
