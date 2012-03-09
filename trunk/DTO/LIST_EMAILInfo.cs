using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class LIST_EMAILInfo
    {
		#region Local Variable
		public enum Field
		{
			Mail,
			Name,
			Lookup
		}
		private String _Mail;
		private String _Name;
		private String _Lookup;
		
		public String Mail{	get{ return _Mail;} set{_Mail = value;} }
		public String Name{	get{ return _Name;} set{_Name = value;} }
		public String Lookup{	get{ return _Lookup;} set{_Lookup = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public LIST_EMAILInfo()
		{
			_Mail = "";
			_Name = "";
			_Lookup = "";
		}
		public LIST_EMAILInfo(
		String Mail,
		String Name,
		String Lookup
		)
		{
			_Mail = Mail;
			_Name = Name;
			_Lookup = Lookup;
		}
		public LIST_EMAILInfo(DataRow dr)
		{
			if (dr != null)
			{
				_Mail = dr[Field.Mail.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Mail.ToString()]);
				_Name = dr[Field.Name.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Name.ToString()]);
				_Lookup = dr[Field.Lookup.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.Lookup.ToString()]);
			}
		}
		public LIST_EMAILInfo(LIST_EMAILInfo objEntr)
		{			
			_Mail = objEntr.Mail;			
			_Name = objEntr.Name;			
			_Lookup = objEntr.Lookup;			
		}
        #endregion Constructor
        
        #region InitTable
		public static DataTable ToDataTable()
		{
			DataTable dt = new DataTable("LIST_EMAIL");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.Mail.ToString(), typeof(String)),
				new DataColumn(Field.Name.ToString(), typeof(String)),
				new DataColumn(Field.Lookup.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.Mail.ToString()] = _Mail;
			row[Field.Name.ToString()] = _Name;
			row[Field.Lookup.ToString()] = _Lookup;
			return row;
		}
        #endregion InitTable
    }
}
