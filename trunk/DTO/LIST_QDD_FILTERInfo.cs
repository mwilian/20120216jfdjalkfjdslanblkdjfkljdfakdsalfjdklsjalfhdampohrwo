using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace DTO
{
	/// <summary> 
	///Author: nnamthach@gmail.com 
	/// <summary>
	
    public class LIST_QDD_FILTERInfo
    {
		#region Local Variable
		public enum Field
		{
			DTB,
			QD_ID,
			QDD_ID,
			OPERATOR,
			IS_NOT
		}
		private String _DTB;
		private String _QD_ID;
		private Int32 _QDD_ID;
		private String _OPERATOR;
		private String _IS_NOT;
		
		public String DTB{	get{ return _DTB;} set{_DTB = value;} }
		public String QD_ID{	get{ return _QD_ID;} set{_QD_ID = value;} }
		public Int32 QDD_ID{	get{ return _QDD_ID;} set{_QDD_ID = value;} }
		public String OPERATOR{	get{ return _OPERATOR;} set{_OPERATOR = value;} }
		public String IS_NOT{	get{ return _IS_NOT;} set{_IS_NOT = value;} }
		
        #endregion LocalVariable
        
        #region Constructor
		public LIST_QDD_FILTERInfo()
		{
			_DTB = "";
			_QD_ID = "";
			_QDD_ID = 0;
			_OPERATOR = "";
			_IS_NOT = "";
		}
		public LIST_QDD_FILTERInfo(
		String DTB,
		String QD_ID,
		Int32 QDD_ID,
		String OPERATOR,
		String IS_NOT
		)
		{
			_DTB = DTB;
			_QD_ID = QD_ID;
			_QDD_ID = QDD_ID;
			_OPERATOR = OPERATOR;
			_IS_NOT = IS_NOT;
		}
		public LIST_QDD_FILTERInfo(DataRow dr)
		{
			if (dr != null)
			{
				_DTB = dr[Field.DTB.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.DTB.ToString()]);
				_QD_ID = dr[Field.QD_ID.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.QD_ID.ToString()]);
				_QDD_ID = dr[Field.QDD_ID.ToString()] == DBNull.Value?0:Convert.ToInt32(dr[Field.QDD_ID.ToString()]);
				_OPERATOR = dr[Field.OPERATOR.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.OPERATOR.ToString()]);
				_IS_NOT = dr[Field.IS_NOT.ToString()] == DBNull.Value?"":Convert.ToString(dr[Field.IS_NOT.ToString()]);
			}
		}
		public LIST_QDD_FILTERInfo(LIST_QDD_FILTERInfo objEntr)
		{			
			_DTB = objEntr.DTB;			
			_QD_ID = objEntr.QD_ID;			
			_QDD_ID = objEntr.QDD_ID;			
			_OPERATOR = objEntr.OPERATOR;			
			_IS_NOT = objEntr.IS_NOT;			
		}
        #endregion Constructor
        
        #region InitTable
		public static DataTable ToDataTable()
		{
			DataTable dt = new DataTable("LIST_QDD_FILTER");
			dt.Columns.AddRange(new DataColumn[] { 
				new DataColumn(Field.DTB.ToString(), typeof(String)),
				new DataColumn(Field.QD_ID.ToString(), typeof(String)),
				new DataColumn(Field.QDD_ID.ToString(), typeof(Int32)),
				new DataColumn(Field.OPERATOR.ToString(), typeof(String)),
				new DataColumn(Field.IS_NOT.ToString(), typeof(String))
			});
			return dt;
		}
		public DataRow ToDataRow(DataTable dt)
		{
			DataRow row = dt.NewRow();
			row[Field.DTB.ToString()] = _DTB;
			row[Field.QD_ID.ToString()] = _QD_ID;
			row[Field.QDD_ID.ToString()] = _QDD_ID;
			row[Field.OPERATOR.ToString()] = _OPERATOR;
			row[Field.IS_NOT.ToString()] = _IS_NOT;
			return row;
		}
        #endregion InitTable
    }
}
