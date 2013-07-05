using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class CoreDAInfo
    {
		#region Local Variable
		public enum Field
		{
			DAG_ID,
			NAME,
			EI
		}
		private String _DAG_ID;
		private String _NAME;
		private String _EI;
		
		public String DAG_ID{	get{ return _DAG_ID;} set{_DAG_ID = value;} }
		public String NAME{	get{ return _NAME;} set{_NAME = value;} }
		public String EI{	get{ return _EI;} set{_EI = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public CoreDAInfo()
		{
			_DAG_ID = "";
			_NAME = "";
			_EI = "";
		}
		public CoreDAInfo(
		String DAG_ID,
		String NAME,
		String EI
		)
		{
			_DAG_ID = DAG_ID;
			_NAME = NAME;
			_EI = EI;
		}
		public CoreDAInfo(DataRow dr)
		{
			if (dr != null)
			{
				_DAG_ID = dr[Field.DAG_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DAG_ID.ToString()]);
				_NAME = dr[Field.NAME.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.NAME.ToString()]);
				_EI = dr[Field.EI.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.EI.ToString()]);
			}
		}
		public CoreDAInfo(CoreDAInfo objEntr)
		{			
			_DAG_ID = objEntr.DAG_ID;			
			_NAME = objEntr.NAME;			
			_EI = objEntr.EI;			
		}
        #endregion Constructor
        
        #region InitTable
		public DataTable ToDataTable()
		{
			DataTable dt = new DataTable("CoreDA");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.DAG_ID.ToString(), typeof(String)),
				new DataColumn(Field.NAME.ToString(), typeof(String)),
				new DataColumn(Field.EI.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.DAG_ID.ToString()] = _DAG_ID;
			row[Field.NAME.ToString()] = _NAME;
			row[Field.EI.ToString()] = _EI;
			return row;
		}
        #endregion InitTable
    }
}
