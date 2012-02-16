using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class POGInfo
    {
		#region Local Variable
		public enum Field
		{
			ROLE_ID,
			TB,
			ROLE_ID1,
			ROLE_NAME,
			PASS_MIN_LEN,
			PASS_VALID,
			RPT_CODE
		}
		private String _ROLE_ID;
		private String _TB;
		private String _ROLE_ID1;
		private String _ROLE_NAME;
		private String _PASS_MIN_LEN;
		private String _PASS_VALID;
		private String _RPT_CODE;
		
		public String ROLE_ID{	get{ return _ROLE_ID;} set{_ROLE_ID = value;} }
		public String TB{	get{ return _TB;} set{_TB = value;} }
		public String ROLE_ID1{	get{ return _ROLE_ID1;} set{_ROLE_ID1 = value;} }
		public String ROLE_NAME{	get{ return _ROLE_NAME;} set{_ROLE_NAME = value;} }
		public String PASS_MIN_LEN{	get{ return _PASS_MIN_LEN;} set{_PASS_MIN_LEN = value;} }
		public String PASS_VALID{	get{ return _PASS_VALID;} set{_PASS_VALID = value;} }
		public String RPT_CODE{	get{ return _RPT_CODE;} set{_RPT_CODE = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public POGInfo()
		{
			_ROLE_ID = "";
			_TB = "";
			_ROLE_ID1 = "";
			_ROLE_NAME = "";
			_PASS_MIN_LEN = "";
			_PASS_VALID = "";
			_RPT_CODE = "";
		}
		public POGInfo(
		String ROLE_ID,
		String TB,
		String ROLE_ID1,
		String ROLE_NAME,
		String PASS_MIN_LEN,
		String PASS_VALID,
		String RPT_CODE
		)
		{
			_ROLE_ID = ROLE_ID;
			_TB = TB;
			_ROLE_ID1 = ROLE_ID1;
			_ROLE_NAME = ROLE_NAME;
			_PASS_MIN_LEN = PASS_MIN_LEN;
			_PASS_VALID = PASS_VALID;
			_RPT_CODE = RPT_CODE;
		}
		public POGInfo(DataRow dr)
		{
			if (dr != null)
			{
				_ROLE_ID = dr[Field.ROLE_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.ROLE_ID.ToString()]);
				_TB = dr[Field.TB.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.TB.ToString()]);
				_ROLE_ID1 = dr[Field.ROLE_ID1.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.ROLE_ID1.ToString()]);
				_ROLE_NAME = dr[Field.ROLE_NAME.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.ROLE_NAME.ToString()]);
				_PASS_MIN_LEN = dr[Field.PASS_MIN_LEN.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PASS_MIN_LEN.ToString()]);
				_PASS_VALID = dr[Field.PASS_VALID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PASS_VALID.ToString()]);
				_RPT_CODE = dr[Field.RPT_CODE.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.RPT_CODE.ToString()]);
			}
		}
		public POGInfo(POGInfo objEntr)
		{			
			_ROLE_ID = objEntr.ROLE_ID;			
			_TB = objEntr.TB;			
			_ROLE_ID1 = objEntr.ROLE_ID1;			
			_ROLE_NAME = objEntr.ROLE_NAME;			
			_PASS_MIN_LEN = objEntr.PASS_MIN_LEN;			
			_PASS_VALID = objEntr.PASS_VALID;			
			_RPT_CODE = objEntr.RPT_CODE;			
		}
        #endregion Constructor
        
        #region InitTable
		public DataTable ToDataTable()
		{
			DataTable dt = new DataTable("POG");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.ROLE_ID.ToString(), typeof(String)),
				new DataColumn(Field.TB.ToString(), typeof(String)),
				new DataColumn(Field.ROLE_ID1.ToString(), typeof(String)),
				new DataColumn(Field.ROLE_NAME.ToString(), typeof(String)),
				new DataColumn(Field.PASS_MIN_LEN.ToString(), typeof(String)),
				new DataColumn(Field.PASS_VALID.ToString(), typeof(String)),
				new DataColumn(Field.RPT_CODE.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.ROLE_ID.ToString()] = _ROLE_ID;
			row[Field.TB.ToString()] = _TB;
			row[Field.ROLE_ID1.ToString()] = _ROLE_ID1;
			row[Field.ROLE_NAME.ToString()] = _ROLE_NAME;
			row[Field.PASS_MIN_LEN.ToString()] = _PASS_MIN_LEN;
			row[Field.PASS_VALID.ToString()] = _PASS_VALID;
			row[Field.RPT_CODE.ToString()] = _RPT_CODE;
			return row;
		}
        #endregion InitTable
    }
}
