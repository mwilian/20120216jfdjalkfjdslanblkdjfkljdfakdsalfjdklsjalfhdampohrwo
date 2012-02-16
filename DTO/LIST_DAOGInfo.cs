using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class LIST_DAOGInfo
    {
		#region Local Variable
		public enum Field
		{
			DAG_ID,
			ROLE_ID
		}
		private String _DAG_ID;
		private String _ROLE_ID;
		
		public String DAG_ID{	get{ return _DAG_ID;} set{_DAG_ID = value;} }
		public String ROLE_ID{	get{ return _ROLE_ID;} set{_ROLE_ID = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public LIST_DAOGInfo()
		{
			_DAG_ID = "";
			_ROLE_ID = "";
		}
		public LIST_DAOGInfo(
		String DAG_ID,
		String ROLE_ID
		)
		{
			_DAG_ID = DAG_ID;
			_ROLE_ID = ROLE_ID;
		}
		public LIST_DAOGInfo(DataRow dr)
		{
			if (dr != null)
			{
				_DAG_ID = dr[Field.DAG_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DAG_ID.ToString()]);
				_ROLE_ID = dr[Field.ROLE_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.ROLE_ID.ToString()]);
			}
		}
		public LIST_DAOGInfo(LIST_DAOGInfo objEntr)
		{			
			_DAG_ID = objEntr.DAG_ID;			
			_ROLE_ID = objEntr.ROLE_ID;			
		}
        #endregion Constructor
        
        #region InitTable
		public DataTable ToDataTable()
		{
			DataTable dt = new DataTable("LIST_DAOG");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.DAG_ID.ToString(), typeof(String)),
				new DataColumn(Field.ROLE_ID.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.DAG_ID.ToString()] = _DAG_ID;
			row[Field.ROLE_ID.ToString()] = _ROLE_ID;
			return row;
		}
        #endregion InitTable
    }
}
