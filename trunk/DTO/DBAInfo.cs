using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class DBAInfo
    {
		#region Local Variable
		public enum Field
		{
			DB,
			DB1,
			DB2,
			DESCRIPTION,
			DATE_FORMAT,
			DECIMAL_PLACES_SUNACCOUNT,
			DECIMAL_SEPERATOR,
			THOUSAND_SEPERATOR,
			PRIMARY_BUDGET,
			DATA_ACCESS_GROUP,
			DECIMAL_PLACES_SUNBUSINESS,
			REPORT_TEMPLATE_DRIVER,
			PARAM_1,
			PARAM_2,
			PARAM_3,
			PARAM_4,
			PARAM_5,
			PARAM_6
		}
		private String _DB;
		private String _DB1;
		private String _DB2;
		private String _DESCRIPTION;
		private String _DATE_FORMAT;
		private String _DECIMAL_PLACES_SUNACCOUNT;
		private String _DECIMAL_SEPERATOR;
		private String _THOUSAND_SEPERATOR;
		private String _PRIMARY_BUDGET;
		private String _DATA_ACCESS_GROUP;
		private String _DECIMAL_PLACES_SUNBUSINESS;
		private String _REPORT_TEMPLATE_DRIVER;
		private String _PARAM_1;
		private String _PARAM_2;
		private String _PARAM_3;
		private String _PARAM_4;
		private String _PARAM_5;
		private String _PARAM_6;
		
		public String DB{	get{ return _DB;} set{_DB = value;} }
		public String DB1{	get{ return _DB1;} set{_DB1 = value;} }
		public String DB2{	get{ return _DB2;} set{_DB2 = value;} }
		public String DESCRIPTION{	get{ return _DESCRIPTION;} set{_DESCRIPTION = value;} }
		public String DATE_FORMAT{	get{ return _DATE_FORMAT;} set{_DATE_FORMAT = value;} }
		public String DECIMAL_PLACES_SUNACCOUNT{	get{ return _DECIMAL_PLACES_SUNACCOUNT;} set{_DECIMAL_PLACES_SUNACCOUNT = value;} }
		public String DECIMAL_SEPERATOR{	get{ return _DECIMAL_SEPERATOR;} set{_DECIMAL_SEPERATOR = value;} }
		public String THOUSAND_SEPERATOR{	get{ return _THOUSAND_SEPERATOR;} set{_THOUSAND_SEPERATOR = value;} }
		public String PRIMARY_BUDGET{	get{ return _PRIMARY_BUDGET;} set{_PRIMARY_BUDGET = value;} }
		public String DATA_ACCESS_GROUP{	get{ return _DATA_ACCESS_GROUP;} set{_DATA_ACCESS_GROUP = value;} }
		public String DECIMAL_PLACES_SUNBUSINESS{	get{ return _DECIMAL_PLACES_SUNBUSINESS;} set{_DECIMAL_PLACES_SUNBUSINESS = value;} }
		public String REPORT_TEMPLATE_DRIVER{	get{ return _REPORT_TEMPLATE_DRIVER;} set{_REPORT_TEMPLATE_DRIVER = value;} }
		public String PARAM_1{	get{ return _PARAM_1;} set{_PARAM_1 = value;} }
		public String PARAM_2{	get{ return _PARAM_2;} set{_PARAM_2 = value;} }
		public String PARAM_3{	get{ return _PARAM_3;} set{_PARAM_3 = value;} }
		public String PARAM_4{	get{ return _PARAM_4;} set{_PARAM_4 = value;} }
		public String PARAM_5{	get{ return _PARAM_5;} set{_PARAM_5 = value;} }
		public String PARAM_6{	get{ return _PARAM_6;} set{_PARAM_6 = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public DBAInfo()
		{
			_DB = "";
			_DB1 = "";
			_DB2 = "";
			_DESCRIPTION = "";
			_DATE_FORMAT = "";
			_DECIMAL_PLACES_SUNACCOUNT = "";
			_DECIMAL_SEPERATOR = "";
			_THOUSAND_SEPERATOR = "";
			_PRIMARY_BUDGET = "";
			_DATA_ACCESS_GROUP = "";
			_DECIMAL_PLACES_SUNBUSINESS = "";
			_REPORT_TEMPLATE_DRIVER = "";
			_PARAM_1 = "";
			_PARAM_2 = "";
			_PARAM_3 = "";
			_PARAM_4 = "";
			_PARAM_5 = "";
			_PARAM_6 = "";
		}
		public DBAInfo(
		String DB,
		String DB1,
		String DB2,
		String DESCRIPTION,
		String DATE_FORMAT,
		String DECIMAL_PLACES_SUNACCOUNT,
		String DECIMAL_SEPERATOR,
		String THOUSAND_SEPERATOR,
		String PRIMARY_BUDGET,
		String DATA_ACCESS_GROUP,
		String DECIMAL_PLACES_SUNBUSINESS,
		String REPORT_TEMPLATE_DRIVER,
		String PARAM_1,
		String PARAM_2,
		String PARAM_3,
		String PARAM_4,
		String PARAM_5,
		String PARAM_6
		)
		{
			_DB = DB;
			_DB1 = DB1;
			_DB2 = DB2;
			_DESCRIPTION = DESCRIPTION;
			_DATE_FORMAT = DATE_FORMAT;
			_DECIMAL_PLACES_SUNACCOUNT = DECIMAL_PLACES_SUNACCOUNT;
			_DECIMAL_SEPERATOR = DECIMAL_SEPERATOR;
			_THOUSAND_SEPERATOR = THOUSAND_SEPERATOR;
			_PRIMARY_BUDGET = PRIMARY_BUDGET;
			_DATA_ACCESS_GROUP = DATA_ACCESS_GROUP;
			_DECIMAL_PLACES_SUNBUSINESS = DECIMAL_PLACES_SUNBUSINESS;
			_REPORT_TEMPLATE_DRIVER = REPORT_TEMPLATE_DRIVER;
			_PARAM_1 = PARAM_1;
			_PARAM_2 = PARAM_2;
			_PARAM_3 = PARAM_3;
			_PARAM_4 = PARAM_4;
			_PARAM_5 = PARAM_5;
			_PARAM_6 = PARAM_6;
		}
		public DBAInfo(DataRow dr)
		{
			if (dr != null)
			{
				_DB = dr[Field.DB.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DB.ToString()]);
				_DB1 = dr[Field.DB1.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DB1.ToString()]);
				_DB2 = dr[Field.DB2.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DB2.ToString()]);
				_DESCRIPTION = dr[Field.DESCRIPTION.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DESCRIPTION.ToString()]);
				_DATE_FORMAT = dr[Field.DATE_FORMAT.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DATE_FORMAT.ToString()]);
				_DECIMAL_PLACES_SUNACCOUNT = dr[Field.DECIMAL_PLACES_SUNACCOUNT.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DECIMAL_PLACES_SUNACCOUNT.ToString()]);
				_DECIMAL_SEPERATOR = dr[Field.DECIMAL_SEPERATOR.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DECIMAL_SEPERATOR.ToString()]);
				_THOUSAND_SEPERATOR = dr[Field.THOUSAND_SEPERATOR.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.THOUSAND_SEPERATOR.ToString()]);
				_PRIMARY_BUDGET = dr[Field.PRIMARY_BUDGET.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PRIMARY_BUDGET.ToString()]);
				_DATA_ACCESS_GROUP = dr[Field.DATA_ACCESS_GROUP.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DATA_ACCESS_GROUP.ToString()]);
				_DECIMAL_PLACES_SUNBUSINESS = dr[Field.DECIMAL_PLACES_SUNBUSINESS.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DECIMAL_PLACES_SUNBUSINESS.ToString()]);
				_REPORT_TEMPLATE_DRIVER = dr[Field.REPORT_TEMPLATE_DRIVER.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.REPORT_TEMPLATE_DRIVER.ToString()]);
				_PARAM_1 = dr[Field.PARAM_1.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PARAM_1.ToString()]);
				_PARAM_2 = dr[Field.PARAM_2.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PARAM_2.ToString()]);
				_PARAM_3 = dr[Field.PARAM_3.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PARAM_3.ToString()]);
				_PARAM_4 = dr[Field.PARAM_4.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PARAM_4.ToString()]);
				_PARAM_5 = dr[Field.PARAM_5.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PARAM_5.ToString()]);
				_PARAM_6 = dr[Field.PARAM_6.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.PARAM_6.ToString()]);
			}
		}
		public DBAInfo(DBAInfo objEntr)
		{			
			_DB = objEntr.DB;			
			_DB1 = objEntr.DB1;			
			_DB2 = objEntr.DB2;			
			_DESCRIPTION = objEntr.DESCRIPTION;			
			_DATE_FORMAT = objEntr.DATE_FORMAT;			
			_DECIMAL_PLACES_SUNACCOUNT = objEntr.DECIMAL_PLACES_SUNACCOUNT;			
			_DECIMAL_SEPERATOR = objEntr.DECIMAL_SEPERATOR;			
			_THOUSAND_SEPERATOR = objEntr.THOUSAND_SEPERATOR;			
			_PRIMARY_BUDGET = objEntr.PRIMARY_BUDGET;			
			_DATA_ACCESS_GROUP = objEntr.DATA_ACCESS_GROUP;			
			_DECIMAL_PLACES_SUNBUSINESS = objEntr.DECIMAL_PLACES_SUNBUSINESS;			
			_REPORT_TEMPLATE_DRIVER = objEntr.REPORT_TEMPLATE_DRIVER;			
			_PARAM_1 = objEntr.PARAM_1;			
			_PARAM_2 = objEntr.PARAM_2;			
			_PARAM_3 = objEntr.PARAM_3;			
			_PARAM_4 = objEntr.PARAM_4;			
			_PARAM_5 = objEntr.PARAM_5;			
			_PARAM_6 = objEntr.PARAM_6;			
		}
        #endregion Constructor
        
        #region InitTable
		public DataTable ToDataTable()
		{
			DataTable dt = new DataTable("DBA");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.DB.ToString(), typeof(String)),
				new DataColumn(Field.DB1.ToString(), typeof(String)),
				new DataColumn(Field.DB2.ToString(), typeof(String)),
				new DataColumn(Field.DESCRIPTION.ToString(), typeof(String)),
				new DataColumn(Field.DATE_FORMAT.ToString(), typeof(String)),
				new DataColumn(Field.DECIMAL_PLACES_SUNACCOUNT.ToString(), typeof(String)),
				new DataColumn(Field.DECIMAL_SEPERATOR.ToString(), typeof(String)),
				new DataColumn(Field.THOUSAND_SEPERATOR.ToString(), typeof(String)),
				new DataColumn(Field.PRIMARY_BUDGET.ToString(), typeof(String)),
				new DataColumn(Field.DATA_ACCESS_GROUP.ToString(), typeof(String)),
				new DataColumn(Field.DECIMAL_PLACES_SUNBUSINESS.ToString(), typeof(String)),
				new DataColumn(Field.REPORT_TEMPLATE_DRIVER.ToString(), typeof(String)),
				new DataColumn(Field.PARAM_1.ToString(), typeof(String)),
				new DataColumn(Field.PARAM_2.ToString(), typeof(String)),
				new DataColumn(Field.PARAM_3.ToString(), typeof(String)),
				new DataColumn(Field.PARAM_4.ToString(), typeof(String)),
				new DataColumn(Field.PARAM_5.ToString(), typeof(String)),
				new DataColumn(Field.PARAM_6.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.DB.ToString()] = _DB;
			row[Field.DB1.ToString()] = _DB1;
			row[Field.DB2.ToString()] = _DB2;
			row[Field.DESCRIPTION.ToString()] = _DESCRIPTION;
			row[Field.DATE_FORMAT.ToString()] = _DATE_FORMAT;
			row[Field.DECIMAL_PLACES_SUNACCOUNT.ToString()] = _DECIMAL_PLACES_SUNACCOUNT;
			row[Field.DECIMAL_SEPERATOR.ToString()] = _DECIMAL_SEPERATOR;
			row[Field.THOUSAND_SEPERATOR.ToString()] = _THOUSAND_SEPERATOR;
			row[Field.PRIMARY_BUDGET.ToString()] = _PRIMARY_BUDGET;
			row[Field.DATA_ACCESS_GROUP.ToString()] = _DATA_ACCESS_GROUP;
			row[Field.DECIMAL_PLACES_SUNBUSINESS.ToString()] = _DECIMAL_PLACES_SUNBUSINESS;
			row[Field.REPORT_TEMPLATE_DRIVER.ToString()] = _REPORT_TEMPLATE_DRIVER;
			row[Field.PARAM_1.ToString()] = _PARAM_1;
			row[Field.PARAM_2.ToString()] = _PARAM_2;
			row[Field.PARAM_3.ToString()] = _PARAM_3;
			row[Field.PARAM_4.ToString()] = _PARAM_4;
			row[Field.PARAM_5.ToString()] = _PARAM_5;
			row[Field.PARAM_6.ToString()] = _PARAM_6;
			return row;
		}
        #endregion InitTable
    }
}
