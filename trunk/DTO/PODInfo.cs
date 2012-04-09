using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class PODInfo
    {
		#region Local Variable
		public enum Field
		{
			USER_ID,
			TB,
			USER_ID1,
			USER_NAME,
			DB_DEFAULT,
			LANGUAGE,
			ROLE_ID,
			PASS
		}
		private String _USER_ID;
		private String _TB;
		private String _USER_ID1;
		private String _USER_NAME;
		private String _DB_DEFAULT;
		private String _LANGUAGE;
		private String _ROLE_ID;
		private String _PASS;
		
		public String USER_ID{	get{ return _USER_ID;} set{_USER_ID = value;} }
		public String TB{	get{ return _TB;} set{_TB = value;} }
		public String USER_ID1{	get{ return _USER_ID1;} set{_USER_ID1 = value;} }
		public String USER_NAME{	get{ return _USER_NAME;} set{_USER_NAME = value;} }
		public String DB_DEFAULT{	get{ return _DB_DEFAULT;} set{_DB_DEFAULT = value;} }
		public String LANGUAGE{	get{ return _LANGUAGE;} set{_LANGUAGE = value;} }
		public String ROLE_ID{	get{ return _ROLE_ID;} set{_ROLE_ID = value;} }
		public String PASS{	get{ return _PASS;} set{_PASS = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public PODInfo()
		{
			_USER_ID = "";
			_TB = "";
			_USER_ID1 = "";
			_USER_NAME = "";
			_DB_DEFAULT = "";
			_LANGUAGE = "";
			_ROLE_ID = "";
			_PASS = "";
		}
		public PODInfo(
		String USER_ID,
		String TB,
		String USER_ID1,
		String USER_NAME,
		String DB_DEFAULT,
		String LANGUAGE,
		String ROLE_ID,
		String PASS
		)
		{
			_USER_ID = USER_ID;
			_TB = TB;
			_USER_ID1 = USER_ID1;
			_USER_NAME = USER_NAME;
			_DB_DEFAULT = DB_DEFAULT;
			_LANGUAGE = LANGUAGE;
			_ROLE_ID = ROLE_ID;
			_PASS = PASS;
		}
		public PODInfo(DataRow dr)
		{
			if (dr != null)
			{
				_USER_ID = dr[Field.USER_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.USER_ID.ToString()]);
				_TB = dr[Field.TB.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.TB.ToString()]);
				_USER_ID1 = dr[Field.USER_ID1.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.USER_ID1.ToString()]);
				_USER_NAME = dr[Field.USER_NAME.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.USER_NAME.ToString()]);
				_DB_DEFAULT = dr[Field.DB_DEFAULT.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DB_DEFAULT.ToString()]);
				_LANGUAGE = dr[Field.LANGUAGE.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.LANGUAGE.ToString()]);
				_ROLE_ID = dr[Field.ROLE_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.ROLE_ID.ToString()]);
				_PASS = dr[Field.PASS.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PASS.ToString()]);
			}
		}
		public PODInfo(PODInfo objEntr)
		{			
			_USER_ID = objEntr.USER_ID;			
			_TB = objEntr.TB;			
			_USER_ID1 = objEntr.USER_ID1;			
			_USER_NAME = objEntr.USER_NAME;			
			_DB_DEFAULT = objEntr.DB_DEFAULT;			
			_LANGUAGE = objEntr.LANGUAGE;			
			_ROLE_ID = objEntr.ROLE_ID;			
			_PASS = objEntr.PASS;			
		}
        #endregion Constructor
        
        #region InitTable
		public static DataTable ToDataTable()
		{
			DataTable dt = new DataTable("POD");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.USER_ID.ToString(), typeof(String)),
				new DataColumn(Field.TB.ToString(), typeof(String)),
				new DataColumn(Field.USER_ID1.ToString(), typeof(String)),
				new DataColumn(Field.USER_NAME.ToString(), typeof(String)),
				new DataColumn(Field.DB_DEFAULT.ToString(), typeof(String)),
				new DataColumn(Field.LANGUAGE.ToString(), typeof(String)),
				new DataColumn(Field.ROLE_ID.ToString(), typeof(String)),
				new DataColumn(Field.PASS.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.USER_ID.ToString()] = _USER_ID;
			row[Field.TB.ToString()] = _TB;
			row[Field.USER_ID1.ToString()] = _USER_ID1;
			row[Field.USER_NAME.ToString()] = _USER_NAME;
			row[Field.DB_DEFAULT.ToString()] = _DB_DEFAULT;
			row[Field.LANGUAGE.ToString()] = _LANGUAGE;
			row[Field.ROLE_ID.ToString()] = _ROLE_ID;
			row[Field.PASS.ToString()] = _PASS;
			return row;
		}
        #endregion InitTable
    }
}
