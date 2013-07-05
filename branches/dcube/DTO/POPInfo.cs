using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class POPInfo
    {
		#region Local Variable
		public enum Field
		{
			ROLE_ID,
			DB,
			DEFAULT_VALUE,
			PERMISSION
		}
		private String _ROLE_ID;
		private String _DB;
		private String _DEFAULT_VALUE;
		private String _PERMISSION;
		
		public String ROLE_ID{	get{ return _ROLE_ID;} set{_ROLE_ID = value;} }
		public String DB{	get{ return _DB;} set{_DB = value;} }
		public String DEFAULT_VALUE{	get{ return _DEFAULT_VALUE;} set{_DEFAULT_VALUE = value;} }
		public String PERMISSION{	get{ return _PERMISSION;} set{_PERMISSION = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public POPInfo()
		{
			_ROLE_ID = "";
			_DB = "";
			_DEFAULT_VALUE = "";
			_PERMISSION = "";
		}
		public POPInfo(
		String ROLE_ID,
		String DB,
		String DEFAULT_VALUE,
		String PERMISSION
		)
		{
			_ROLE_ID = ROLE_ID;
			_DB = DB;
			_DEFAULT_VALUE = DEFAULT_VALUE;
			_PERMISSION = PERMISSION;
		}
		public POPInfo(DataRow dr)
		{
			if (dr != null)
			{
				_ROLE_ID = dr[Field.ROLE_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.ROLE_ID.ToString()]);
				_DB = dr[Field.DB.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DB.ToString()]);
				_DEFAULT_VALUE = dr[Field.DEFAULT_VALUE.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DEFAULT_VALUE.ToString()]);
				_PERMISSION = dr[Field.PERMISSION.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PERMISSION.ToString()]);
			}
		}
		public POPInfo(POPInfo objEntr)
		{			
			_ROLE_ID = objEntr.ROLE_ID;			
			_DB = objEntr.DB;			
			_DEFAULT_VALUE = objEntr.DEFAULT_VALUE;			
			_PERMISSION = objEntr.PERMISSION;			
		}
        #endregion Constructor
        
        #region InitTable
		public DataTable ToDataTable()
		{
			DataTable dt = new DataTable("POP");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.ROLE_ID.ToString(), typeof(String)),
				new DataColumn(Field.DB.ToString(), typeof(String)),
				new DataColumn(Field.DEFAULT_VALUE.ToString(), typeof(String)),
				new DataColumn(Field.PERMISSION.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.ROLE_ID.ToString()] = _ROLE_ID;
			row[Field.DB.ToString()] = _DB;
			row[Field.DEFAULT_VALUE.ToString()] = _DEFAULT_VALUE;
			row[Field.PERMISSION.ToString()] = _PERMISSION;
			return row;
		}
        #endregion InitTable
    }
}
