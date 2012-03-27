using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class LIST_QD_SCHEMAInfo
    {
		#region Local Variable
		public enum Field
		{
			CONN_ID,
			SCHEMA_ID,
			LOOK_UP,
			DESCRIPTN,
			FIELD_TEXT,
			FROM_TEXT,
			DAG,
			SCHEMA_STATUS,
			UPDATED,
			ENTER_BY,
			DEFAULT_CONN
		}
		private String _CONN_ID;
		private String _SCHEMA_ID;
		private String _LOOK_UP;
		private String _DESCRIPTN;
		private String _FIELD_TEXT;
		private String _FROM_TEXT;
		private String _DAG;
		private String _SCHEMA_STATUS;
		private Int32 _UPDATED;
		private String _ENTER_BY;
		private String _DEFAULT_CONN;
		
		public String CONN_ID{	get{ return _CONN_ID;} set{_CONN_ID = value;} }
		public String SCHEMA_ID{	get{ return _SCHEMA_ID;} set{_SCHEMA_ID = value;} }
		public String LOOK_UP{	get{ return _LOOK_UP;} set{_LOOK_UP = value;} }
		public String DESCRIPTN{	get{ return _DESCRIPTN;} set{_DESCRIPTN = value;} }
		public String FIELD_TEXT{	get{ return _FIELD_TEXT;} set{_FIELD_TEXT = value;} }
		public String FROM_TEXT{	get{ return _FROM_TEXT;} set{_FROM_TEXT = value;} }
		public String DAG{	get{ return _DAG;} set{_DAG = value;} }
		public String SCHEMA_STATUS{	get{ return _SCHEMA_STATUS;} set{_SCHEMA_STATUS = value;} }
		public Int32 UPDATED{	get{ return _UPDATED;} set{_UPDATED = value;} }
		public String ENTER_BY{	get{ return _ENTER_BY;} set{_ENTER_BY = value;} }
		public String DEFAULT_CONN{	get{ return _DEFAULT_CONN;} set{_DEFAULT_CONN = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public LIST_QD_SCHEMAInfo()
		{
			_CONN_ID = "";
			_SCHEMA_ID = "";
			_LOOK_UP = "";
			_DESCRIPTN = "";
			_FIELD_TEXT = "";
			_FROM_TEXT = "";
			_DAG = "";
			_SCHEMA_STATUS = "";
			_UPDATED = 0;
			_ENTER_BY = "";
			_DEFAULT_CONN = "";
		}
		public LIST_QD_SCHEMAInfo(
		String CONN_ID,
		String SCHEMA_ID,
		String LOOK_UP,
		String DESCRIPTN,
		String FIELD_TEXT,
		String FROM_TEXT,
		String DAG,
		String SCHEMA_STATUS,
		Int32 UPDATED,
		String ENTER_BY,
		String DEFAULT_CONN
		)
		{
			_CONN_ID = CONN_ID;
			_SCHEMA_ID = SCHEMA_ID;
			_LOOK_UP = LOOK_UP;
			_DESCRIPTN = DESCRIPTN;
			_FIELD_TEXT = FIELD_TEXT;
			_FROM_TEXT = FROM_TEXT;
			_DAG = DAG;
			_SCHEMA_STATUS = SCHEMA_STATUS;
			_UPDATED = UPDATED;
			_ENTER_BY = ENTER_BY;
			_DEFAULT_CONN = DEFAULT_CONN;
		}
		public LIST_QD_SCHEMAInfo(DataRow dr)
		{
			if (dr != null)
			{
				_CONN_ID = dr[Field.CONN_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.CONN_ID.ToString()]);
				_SCHEMA_ID = dr[Field.SCHEMA_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.SCHEMA_ID.ToString()]);
				_LOOK_UP = dr[Field.LOOK_UP.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.LOOK_UP.ToString()]);
				_DESCRIPTN = dr[Field.DESCRIPTN.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DESCRIPTN.ToString()]);
				_FIELD_TEXT = dr[Field.FIELD_TEXT.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.FIELD_TEXT.ToString()]);
				_FROM_TEXT = dr[Field.FROM_TEXT.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.FROM_TEXT.ToString()]);
				_DAG = dr[Field.DAG.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DAG.ToString()]);
				_SCHEMA_STATUS = dr[Field.SCHEMA_STATUS.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.SCHEMA_STATUS.ToString()]);
				_UPDATED = dr[Field.UPDATED.ToString()] == DBNull.Value?0:Convert.ToInt32(dr[Field.UPDATED.ToString()]);
				_ENTER_BY = dr[Field.ENTER_BY.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.ENTER_BY.ToString()]);
				_DEFAULT_CONN = dr[Field.DEFAULT_CONN.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DEFAULT_CONN.ToString()]);
			}
		}
		public LIST_QD_SCHEMAInfo(LIST_QD_SCHEMAInfo objEntr)
		{			
			_CONN_ID = objEntr.CONN_ID;			
			_SCHEMA_ID = objEntr.SCHEMA_ID;			
			_LOOK_UP = objEntr.LOOK_UP;			
			_DESCRIPTN = objEntr.DESCRIPTN;			
			_FIELD_TEXT = objEntr.FIELD_TEXT;			
			_FROM_TEXT = objEntr.FROM_TEXT;			
			_DAG = objEntr.DAG;			
			_SCHEMA_STATUS = objEntr.SCHEMA_STATUS;			
			_UPDATED = objEntr.UPDATED;			
			_ENTER_BY = objEntr.ENTER_BY;			
			_DEFAULT_CONN = objEntr.DEFAULT_CONN;			
		}
        #endregion Constructor
        
        #region InitTable
		public DataTable ToDataTable()
		{
			DataTable dt = new DataTable("LIST_QD_SCHEMA");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.CONN_ID.ToString(), typeof(String)),
				new DataColumn(Field.SCHEMA_ID.ToString(), typeof(String)),
				new DataColumn(Field.LOOK_UP.ToString(), typeof(String)),
				new DataColumn(Field.DESCRIPTN.ToString(), typeof(String)),
				new DataColumn(Field.FIELD_TEXT.ToString(), typeof(String)),
				new DataColumn(Field.FROM_TEXT.ToString(), typeof(String)),
				new DataColumn(Field.DAG.ToString(), typeof(String)),
				new DataColumn(Field.SCHEMA_STATUS.ToString(), typeof(String)),
				new DataColumn(Field.UPDATED.ToString(), typeof(Int32)),
				new DataColumn(Field.ENTER_BY.ToString(), typeof(String)),
				new DataColumn(Field.DEFAULT_CONN.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.CONN_ID.ToString()] = _CONN_ID;
			row[Field.SCHEMA_ID.ToString()] = _SCHEMA_ID;
			row[Field.LOOK_UP.ToString()] = _LOOK_UP;
			row[Field.DESCRIPTN.ToString()] = _DESCRIPTN;
			row[Field.FIELD_TEXT.ToString()] = _FIELD_TEXT;
			row[Field.FROM_TEXT.ToString()] = _FROM_TEXT;
			row[Field.DAG.ToString()] = _DAG;
			row[Field.SCHEMA_STATUS.ToString()] = _SCHEMA_STATUS;
			row[Field.UPDATED.ToString()] = _UPDATED;
			row[Field.ENTER_BY.ToString()] = _ENTER_BY;
			row[Field.DEFAULT_CONN.ToString()] = _DEFAULT_CONN;
			return row;
		}
        #endregion InitTable
    }
}
